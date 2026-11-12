using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Read-only Jira REST adapter behind <see cref="IJiraTestCaseSource"/>.
/// Vendor mapping is configuration-driven; this class does not hardcode Xray/Zephyr APIs.
/// </summary>
public sealed class JiraRestTestCaseSource : IJiraTestCaseSource
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly IOptions<AgentOptions> _optionsAccessor;
    private readonly IWorkspaceManager _workspace;
    private readonly ILogger<JiraRestTestCaseSource> _logger;
    private readonly object _gate = new();
    private JiraHttpTransport? _transport;
    private JiraRawCache? _cache;
    private JiraOptions? _options;
    private JiraFieldMappingOptions? _mapping;
    private IReadOnlyList<string>? _issueFields;

    public JiraRestTestCaseSource(
        IHttpClientFactory httpClientFactory,
        IOptions<AgentOptions> options,
        IWorkspaceManager workspace,
        ILogger<JiraRestTestCaseSource> logger)
    {
        _httpClientFactory = httpClientFactory;
        _optionsAccessor = options;
        _workspace = workspace;
        _logger = logger;
    }

    /// <summary>Test helper with an explicit workspace root (avoids full workspace init).</summary>
    internal JiraRestTestCaseSource(
        IHttpClientFactory httpClientFactory,
        AgentOptions agentOptions,
        string workspaceRoot,
        ILogger<JiraRestTestCaseSource> logger)
        : this(
            httpClientFactory,
            Options.Create(agentOptions),
            new FixedWorkspaceRoot(workspaceRoot),
            logger)
    {
    }

    public async Task<JiraTestCaseFetchResult> FetchAsync(
        string jiraKey,
        CancellationToken cancellationToken = default)
    {
        if (!JiraKeyContract.IsCanonical(jiraKey))
        {
            return new JiraTestCaseFetchResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                JiraKey = jiraKey,
                Errors = [$"invalid jiraKey '{jiraKey}' — rejected before network."]
            };
        }

        EnsureReady();

        var conditional = _cache!.GetConditionalHeaders(jiraKey);
        var fieldsQuery = string.Join(',', _issueFields!);
        var path = $"rest/api/3/issue/{Uri.EscapeDataString(jiraKey)}?fields={Uri.EscapeDataString(fieldsQuery)}";

        _logger.LogInformation(
            "Fetching Jira issue {JiraKey} via read-only REST (fields={FieldCount}). Authorization redacted.",
            jiraKey,
            _issueFields!.Count);

        var response = await _transport!.GetAsync(path, conditional?.Headers, cancellationToken)
            .ConfigureAwait(false);

        if (response.NotModified && conditional is not null)
        {
            if (_cache.TryRead(jiraKey, conditional.Metadata.RevisionToken, out var cachedDoc, out var cachedMeta, out _, out var cacheErrors)
                && cachedDoc is not null
                && cachedMeta is not null)
            {
                return new JiraTestCaseFetchResult
                {
                    Success = true,
                    ExitCode = ExitCodes.Success,
                    JiraKey = jiraKey,
                    Document = cachedDoc,
                    Metadata = cachedMeta with
                    {
                        ETag = response.ETag ?? cachedMeta.ETag,
                        LastModified = response.LastModified ?? cachedMeta.LastModified
                    },
                    NotModified = true,
                    FromCache = true,
                    RateLimit = response.RateLimit
                };
            }

            return new JiraTestCaseFetchResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                JiraKey = jiraKey,
                Errors = cacheErrors.Count > 0
                    ? cacheErrors
                    : ["Jira returned 304 but local cache was missing/invalid."],
                RateLimit = response.RateLimit
            };
        }

        if (response.ExitCode != ExitCodes.Success || response.Body is null)
        {
            return new JiraTestCaseFetchResult
            {
                Success = false,
                ExitCode = response.ExitCode == 0 ? ExitCodes.SuiteOrWorkspace : response.ExitCode,
                JiraKey = jiraKey,
                Errors = response.Errors.Select(SecretRedactor.Redact).ToArray(),
                RateLimit = response.RateLimit
            };
        }

        JsonDocument jsonDoc;
        try
        {
            jsonDoc = JsonDocument.Parse(response.Body);
        }
        catch (JsonException ex)
        {
            return new JiraTestCaseFetchResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                JiraKey = jiraKey,
                Errors = [$"Jira response JSON parse failed: {SecretRedactor.Redact(ex.Message)}"],
                RateLimit = response.RateLimit
            };
        }

        using (jsonDoc)
        {
            if (!JiraVendorMapper.TryMap(
                    jsonDoc.RootElement,
                    _mapping!,
                    jiraKey,
                    out var document,
                    out var metadata,
                    out var mapErrors,
                    out _))
            {
                return new JiraTestCaseFetchResult
                {
                    Success = false,
                    ExitCode = ExitCodes.SuiteOrWorkspace,
                    JiraKey = jiraKey,
                    Errors = mapErrors.Select(SecretRedactor.Redact).ToArray(),
                    RateLimit = response.RateLimit
                };
            }

            metadata = metadata with
            {
                MappingProfile = (_options!.MappingProfile ?? "plain-fields").Trim().ToLowerInvariant(),
                ETag = response.ETag,
                LastModified = response.LastModified
            };

            try
            {
                _cache.WriteAtomic(jiraKey, document, metadata, out _, out metadata, out document);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException)
            {
                return new JiraTestCaseFetchResult
                {
                    Success = false,
                    ExitCode = ExitCodes.SuiteOrWorkspace,
                    JiraKey = jiraKey,
                    Document = document,
                    Metadata = metadata,
                    Errors = [$"failed to write jira-cache: {SecretRedactor.Redact(ex.Message)}"],
                    RateLimit = response.RateLimit
                };
            }

            return new JiraTestCaseFetchResult
            {
                Success = true,
                ExitCode = ExitCodes.Success,
                JiraKey = jiraKey,
                Document = document,
                Metadata = metadata,
                FromCache = false,
                RateLimit = response.RateLimit
            };
        }
    }

    public async Task<JiraKeySearchResult> SearchKeysAsync(
        string jql,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(jql))
        {
            return new JiraKeySearchResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                Errors = ["JQL is required."]
            };
        }

        EnsureReady();

        // Cloud enhanced JQL search (GET /rest/api/3/search/jql) — cursor via nextPageToken/isLast.
        // Legacy DC startAt/total offset strategy is intentionally not the default; set
        // Jira:SearchPagination=legacy-offset only if a Data Center adapter is required later.
        if (!string.Equals(_options!.SearchPagination, "jql-token", StringComparison.OrdinalIgnoreCase))
        {
            return new JiraKeySearchResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                Errors =
                [
                    $"Jira:SearchPagination '{_options.SearchPagination}' is not supported. " +
                    "Use 'jql-token' (/rest/api/3/search/jql)."
                ]
            };
        }

        var keys = new List<string>();
        JiraRateLimitInfo? lastRate = null;
        string? nextPageToken = null;
        var seenTokens = new HashSet<string>(StringComparer.Ordinal);
        var pageCount = 0;
        var complete = false;

        while (pageCount < _options.MaxPages)
        {
            pageCount++;
            var path =
                $"rest/api/3/search/jql?jql={Uri.EscapeDataString(jql)}" +
                $"&maxResults={_options.PageSize}&fields=key";
            if (!string.IsNullOrWhiteSpace(nextPageToken))
                path += $"&nextPageToken={Uri.EscapeDataString(nextPageToken)}";

            var response = await _transport!.GetAsync(path, null, cancellationToken).ConfigureAwait(false);
            lastRate = response.RateLimit;
            if (response.ExitCode != ExitCodes.Success || response.Body is null)
            {
                return new JiraKeySearchResult
                {
                    Success = false,
                    ExitCode = response.ExitCode == 0 ? ExitCodes.SuiteOrWorkspace : response.ExitCode,
                    Keys = keys,
                    PageCount = pageCount,
                    Errors = response.Errors.Select(SecretRedactor.Redact).ToArray(),
                    RateLimit = lastRate
                };
            }

            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(response.Body);
            }
            catch (JsonException ex)
            {
                return new JiraKeySearchResult
                {
                    Success = false,
                    ExitCode = ExitCodes.SuiteOrWorkspace,
                    Keys = keys,
                    PageCount = pageCount,
                    Errors = [$"Jira search page JSON parse failed: {SecretRedactor.Redact(ex.Message)}"],
                    RateLimit = lastRate
                };
            }

            using (doc)
            {
                if (!doc.RootElement.TryGetProperty("issues", out var issues)
                    || issues.ValueKind != JsonValueKind.Array)
                {
                    return new JiraKeySearchResult
                    {
                        Success = false,
                        ExitCode = ExitCodes.SuiteOrWorkspace,
                        Keys = keys,
                        PageCount = pageCount,
                        Errors = ["Jira search response missing issues array."],
                        RateLimit = lastRate
                    };
                }

                foreach (var issue in issues.EnumerateArray())
                {
                    var key = issue.TryGetProperty("key", out var keyEl) ? keyEl.GetString() : null;
                    if (!JiraKeyContract.IsCanonical(key ?? string.Empty))
                    {
                        return new JiraKeySearchResult
                        {
                            Success = false,
                            ExitCode = ExitCodes.SuiteOrWorkspace,
                            Keys = keys,
                            PageCount = pageCount,
                            Errors = [$"search returned non-canonical key '{key}'."],
                            RateLimit = lastRate
                        };
                    }

                    keys.Add(key!);
                }

                if (!TryReadIsLast(doc.RootElement, out var isLast, out var isLastError))
                {
                    return new JiraKeySearchResult
                    {
                        Success = false,
                        ExitCode = ExitCodes.SuiteOrWorkspace,
                        Keys = keys,
                        PageCount = pageCount,
                        Errors = [isLastError],
                        RateLimit = lastRate
                    };
                }

                if (isLast)
                {
                    complete = true;
                    break;
                }

                if (!TryReadNextPageToken(doc.RootElement, out var token, out var tokenError))
                {
                    return new JiraKeySearchResult
                    {
                        Success = false,
                        ExitCode = ExitCodes.SuiteOrWorkspace,
                        Keys = keys,
                        PageCount = pageCount,
                        Errors = [tokenError],
                        RateLimit = lastRate
                    };
                }

                if (!seenTokens.Add(token))
                {
                    return new JiraKeySearchResult
                    {
                        Success = false,
                        ExitCode = ExitCodes.SuiteOrWorkspace,
                        Keys = keys,
                        PageCount = pageCount,
                        Errors = ["Jira search returned a repeated nextPageToken while isLast=false."],
                        RateLimit = lastRate
                    };
                }

                nextPageToken = token;
            }
        }

        if (!complete)
        {
            return new JiraKeySearchResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                Keys = keys,
                PageCount = pageCount,
                Errors =
                [
                    $"Jira search truncated after MaxPages={_options.MaxPages} " +
                    $"(retrieved {keys.Count} keys; isLast was still false)."
                ],
                RateLimit = lastRate
            };
        }

        return new JiraKeySearchResult
        {
            Success = true,
            ExitCode = ExitCodes.Success,
            Keys = keys,
            PageCount = pageCount,
            RateLimit = lastRate
        };
    }

    private static bool TryReadIsLast(JsonElement root, out bool isLast, out string error)
    {
        isLast = false;
        error = string.Empty;
        if (!root.TryGetProperty("isLast", out var isLastEl))
        {
            error = "Jira search response missing required 'isLast' field.";
            return false;
        }

        if (isLastEl.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            error = "Jira search response 'isLast' must be a boolean.";
            return false;
        }

        isLast = isLastEl.GetBoolean();
        return true;
    }

    private static bool TryReadNextPageToken(JsonElement root, out string token, out string error)
    {
        token = string.Empty;
        error = string.Empty;
        if (!root.TryGetProperty("nextPageToken", out var tokenEl)
            || tokenEl.ValueKind == JsonValueKind.Null)
        {
            error = "Jira search response missing nextPageToken while isLast=false.";
            return false;
        }

        if (tokenEl.ValueKind != JsonValueKind.String)
        {
            error = "Jira search response 'nextPageToken' must be a string when isLast=false.";
            return false;
        }

        token = tokenEl.GetString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(token))
        {
            error = "Jira search response nextPageToken was blank while isLast=false.";
            return false;
        }

        return true;
    }

    private void EnsureReady()
    {
        if (_transport is not null)
            return;

        lock (_gate)
        {
            if (_transport is not null)
                return;

            var jira = _optionsAccessor.Value.Jira;
            AgentOptionsValidator.ValidateJiraOptions(jira);
            _options = jira;
            _mapping = JiraVendorMapper.ResolveMapping(jira);
            _issueFields = JiraVendorMapper.ResolveIssueFields(jira, _mapping);
            _cache = new JiraRawCache(_workspace.RootPath);
            _transport = new JiraHttpTransport(_httpClientFactory, jira, _logger);
        }
    }

    private sealed class FixedWorkspaceRoot : IWorkspaceManager
    {
        public FixedWorkspaceRoot(string root) => RootPath = Path.GetFullPath(root);
        public string RootPath { get; }
        public WorkspaceInitResult Initialize() => new() { RootPath = RootPath };
        public void EnsureInitialized() { }
        public string ResolveSafePath(string path) => Path.GetFullPath(Path.Combine(RootPath, path));
    }
}
