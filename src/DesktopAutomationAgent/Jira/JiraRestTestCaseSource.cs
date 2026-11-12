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
                _cache.WriteAtomic(jiraKey, document, metadata, out _);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
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

        var keys = new List<string>();
        JiraRateLimitInfo? lastRate = null;
        var startAt = 0;
        var pageCount = 0;

        while (pageCount < _options!.MaxPages)
        {
            pageCount++;
            var path =
                $"rest/api/3/search?jql={Uri.EscapeDataString(jql)}" +
                $"&startAt={startAt}&maxResults={_options.PageSize}&fields=key";

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

            using var doc = JsonDocument.Parse(response.Body);
            if (!doc.RootElement.TryGetProperty("issues", out var issues)
                || issues.ValueKind != JsonValueKind.Array)
            {
                return new JiraKeySearchResult
                {
                    Success = false,
                    ExitCode = ExitCodes.SuiteOrWorkspace,
                    PageCount = pageCount,
                    Errors = ["Jira search response missing issues array."],
                    RateLimit = lastRate
                };
            }

            var pageKeys = 0;
            foreach (var issue in issues.EnumerateArray())
            {
                pageKeys++;
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

            var total = doc.RootElement.TryGetProperty("total", out var totalEl) && totalEl.TryGetInt32(out var t)
                ? t
                : keys.Count;
            startAt += pageKeys;
            if (pageKeys == 0 || startAt >= total)
                break;
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
