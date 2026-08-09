using System.Diagnostics;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectVerificationService
{
    private static readonly HashSet<string> AllowedViews = new(StringComparer.Ordinal)
    {
        "control",
        "content",
        "raw"
    };

    private static readonly HashSet<string> AllowedRoots = new(StringComparer.Ordinal)
    {
        "activeWindow",
        "processWindows",
        "desktopChildren"
    };

    private readonly AgentOptions _options;
    private readonly ObjectRepositoryReader _repositoryReader;
    private readonly ObjectReferenceResolver _referenceResolver;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly IDriverUiClient _uiClient;
    private readonly IWorkspaceManager _workspace;
    private readonly ILogger<ObjectVerificationService> _logger;

    public ObjectVerificationService(
        IOptions<AgentOptions> options,
        ObjectRepositoryReader repositoryReader,
        ObjectReferenceResolver referenceResolver,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        IDriverUiClient uiClient,
        IWorkspaceManager workspace,
        ILogger<ObjectVerificationService> logger)
    {
        _options = options.Value;
        _repositoryReader = repositoryReader;
        _referenceResolver = referenceResolver;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _uiClient = uiClient;
        _workspace = workspace;
        _logger = logger;
    }

    public async Task<ObjectVerificationResult> VerifyAsync(
        string repositoryPath,
        ObjectVerificationOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(options);

        var stopwatch = Stopwatch.StartNew();

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.ObjectRepository);
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            _workspace.EnsureInitialized();
        }
        catch (AgentConfigurationException ex)
        {
            return Failure(ex.Message, ExitCodes.UsageOrConfiguration, stopwatch);
        }
        catch (WorkspaceException ex)
        {
            return Failure(ex.Message, ExitCodes.SuiteOrWorkspace, stopwatch);
        }

        var validation = _repositoryReader.Read(repositoryPath);
        if (!validation.IsValid || validation.Snapshot is null)
        {
            return Failure(
                validation.Errors.FirstOrDefault() ?? "Object repository validation failed.",
                ExitCodes.SuiteOrWorkspace,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot?.Manifest.RepositoryId,
                validation.AggregateSha256);
        }

        var references = SelectReferences(validation.Snapshot, options, out var selectionError);
        if (selectionError is not null)
        {
            return Failure(selectionError, ExitCodes.UsageOrConfiguration, stopwatch, validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId, validation.Snapshot.AggregateSha256);
        }

        if (!TryResolveOption(options.View, AllowedViews, "control", out var view, out var viewError))
        {
            return Failure(viewError!, ExitCodes.UsageOrConfiguration, stopwatch, validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId, validation.Snapshot.AggregateSha256);
        }

        if (!TryResolveOption(options.Root, AllowedRoots, "activeWindow", out var root, out var rootError))
        {
            return Failure(rootError!, ExitCodes.UsageOrConfiguration, stopwatch, validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId, validation.Snapshot.AggregateSha256);
        }

        var maxDepth = options.MaxDepth ?? 8;
        if (maxDepth is < 0 or > 20)
        {
            return Failure(
                "maxDepth must be between 0 and 20.",
                ExitCodes.UsageOrConfiguration,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId,
                validation.Snapshot.AggregateSha256);
        }

        var maxChildren = options.MaxChildren ?? 200;
        if (maxChildren is < 1 or > 1000)
        {
            return Failure(
                "maxChildren must be between 1 and 1000.",
                ExitCodes.UsageOrConfiguration,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId,
                validation.Snapshot.AggregateSha256);
        }

        var includeOffscreen = options.IncludeOffscreen ?? false;

        DriverConnection connection;
        try
        {
            connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            var status = await _catalogClient.GetStatusAsync(connection, cancellationToken).ConfigureAwait(false);
            if (!status.Ready)
            {
                return Failure(
                    "Driver /status reported ready=false.",
                    ExitCodes.DriverUnavailable,
                    stopwatch,
                    validation.RepositoryPath,
                    validation.Snapshot.Manifest.RepositoryId,
                    validation.Snapshot.AggregateSha256);
            }

            var catalog = await _catalogClient.GetOperationsAsync(connection, cancellationToken).ConfigureAwait(false);
            CatalogCompatibility.Validate(catalog, _options.Driver.ExpectedCatalogSchemaVersion);
            DiagnosticCatalogValidator.RequireDiagnosticOperation(catalog, "finduia");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return Failure(
                "Verification was cancelled.",
                ExitCodes.Cancelled,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId,
                validation.Snapshot.AggregateSha256);
        }
        catch (DriverConnectionException ex)
        {
            return Failure(
                ex.Message,
                ExitCodes.DriverUnavailable,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId,
                validation.Snapshot.AggregateSha256);
        }
        catch (DriverCatalogException ex)
        {
            return Failure(
                ex.Message,
                ExitCodes.AuthOrCatalog,
                stopwatch,
                validation.RepositoryPath,
                validation.Snapshot.Manifest.RepositoryId,
                validation.Snapshot.AggregateSha256);
        }

        var items = new List<ObjectVerificationItemResult>();
        var passed = 0;
        var missing = 0;
        var ambiguous = 0;
        var fragile = 0;
        var failed = 0;

        foreach (var reference in references)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return BuildSummary(
                    validation,
                    items,
                    passed,
                    missing,
                    ambiguous,
                    fragile,
                    failed,
                    stopwatch,
                    ExitCodes.Cancelled,
                    success: false,
                    error: "Verification was cancelled.");
            }

            var resolution = _referenceResolver.Resolve(validation.Snapshot, reference);
            if (!resolution.IsResolved || resolution.Locator is null)
            {
                failed++;
                items.Add(new ObjectVerificationItemResult
                {
                    Reference = reference,
                    Status = ObjectVerificationStatus.Failed,
                    MatchCount = 0,
                    Error = resolution.Errors.FirstOrDefault() ?? "Failed to resolve object reference."
                });
                continue;
            }

            ObjectVerificationItemResult itemResult;
            try
            {
                itemResult = await VerifyReferenceAsync(
                        connection,
                        reference,
                        resolution.Locator,
                        view,
                        root,
                        maxDepth,
                        maxChildren,
                        includeOffscreen,
                        cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return BuildSummary(
                    validation,
                    items,
                    passed,
                    missing,
                    ambiguous,
                    fragile,
                    failed,
                    stopwatch,
                    ExitCodes.Cancelled,
                    success: false,
                    error: "Verification was cancelled.");
            }
            catch (UiExecutionException ex) when (ex.Classification == UiFailureClassification.Cancelled)
            {
                return BuildSummary(
                    validation,
                    items,
                    passed,
                    missing,
                    ambiguous,
                    fragile,
                    failed,
                    stopwatch,
                    ExitCodes.Cancelled,
                    success: false,
                    error: SecretRedactor.Redact(ex.Message));
            }

            items.Add(itemResult);

            switch (itemResult.Status)
            {
                case ObjectVerificationStatus.Passed:
                    passed++;
                    break;
                case ObjectVerificationStatus.Missing:
                    missing++;
                    break;
                case ObjectVerificationStatus.Ambiguous:
                    ambiguous++;
                    break;
                case ObjectVerificationStatus.Fragile:
                    fragile++;
                    passed++;
                    break;
                default:
                    failed++;
                    break;
            }
        }

        var exitCode = failed > 0 || missing > 0 || ambiguous > 0
            ? ExitCodes.ExecutionFailure
            : ExitCodes.Success;

        return BuildSummary(
            validation,
            items,
            passed,
            missing,
            ambiguous,
            fragile,
            failed,
            stopwatch,
            exitCode,
            success: exitCode == ExitCodes.Success);
    }

    private async Task<ObjectVerificationItemResult> VerifyReferenceAsync(
        DriverConnection connection,
        string reference,
        ObjectLocator locator,
        string view,
        string root,
        int maxDepth,
        int maxChildren,
        bool includeOffscreen,
        CancellationToken cancellationToken)
    {
        var foundIndex = locator.FoundIndex;
        var requestLocator = foundIndex is null
            ? locator
            : new ObjectLocator
            {
                AutomationId = locator.AutomationId,
                Name = locator.Name,
                ClassName = locator.ClassName,
                ControlType = locator.ControlType,
                MatchMode = locator.MatchMode
            };

        var arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
        {
            ["locator"] = ObjectLocatorSerializer.ToJsonElement(requestLocator),
            ["root"] = JsonSerializer.SerializeToElement(root),
            ["view"] = JsonSerializer.SerializeToElement(view),
            ["maxDepth"] = JsonSerializer.SerializeToElement(maxDepth),
            ["maxChildren"] = JsonSerializer.SerializeToElement(maxChildren),
            ["includeOffscreen"] = JsonSerializer.SerializeToElement(includeOffscreen),
            ["includePath"] = JsonSerializer.SerializeToElement(true),
            ["timeoutMs"] = JsonSerializer.SerializeToElement(_options.ObjectRepository.DiagnosticTimeoutMilliseconds)
        };

        var step = new PlanStep
        {
            Id = $"verify-{reference}",
            Operation = "finduia",
            Arguments = arguments
        };

        UiExecutionResponse response;
        try
        {
            response = await _uiClient.ExecuteStepAsync(connection, step, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (UiExecutionException ex) when (ex.Classification == UiFailureClassification.Cancelled)
        {
            throw;
        }
        catch (UiExecutionException ex)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Failed,
                MatchCount = 0,
                Error = SecretRedactor.Redact(ex.Message)
            };
        }

        if (response.Value is not JsonElement value || value.ValueKind != JsonValueKind.Object)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Failed,
                MatchCount = 0,
                Error = "finduia response did not include a value object."
            };
        }

        if (IsTimeoutOrPartial(value))
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Failed,
                MatchCount = 0,
                Error = "finduia returned partial or timed out results."
            };
        }

        var matchCount = value.TryGetProperty("matchCount", out var countElement) && countElement.TryGetInt32(out var count)
            ? count
            : 0;

        if (matchCount == 0)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Missing,
                MatchCount = 0
            };
        }

        if (matchCount == 1)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Passed,
                MatchCount = 1
            };
        }

        if (foundIndex is null)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Ambiguous,
                MatchCount = matchCount,
                Error = $"finduia matched {matchCount} elements without foundIndex."
            };
        }

        if (foundIndex.Value >= 0 && foundIndex.Value < matchCount)
        {
            return new ObjectVerificationItemResult
            {
                Reference = reference,
                Status = ObjectVerificationStatus.Fragile,
                MatchCount = matchCount,
                Warning = $"foundIndex {foundIndex.Value} selected one of {matchCount} matches."
            };
        }

        return new ObjectVerificationItemResult
        {
            Reference = reference,
            Status = ObjectVerificationStatus.Failed,
            MatchCount = matchCount,
            Error = $"foundIndex {foundIndex.Value} is out of range for {matchCount} matches."
        };
    }

    private static bool TryResolveOption(
        string? value,
        HashSet<string> allowed,
        string fallback,
        out string resolved,
        out string? error)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            resolved = fallback;
            error = null;
            return true;
        }

        var normalized = value.Trim();
        if (allowed.Contains(normalized))
        {
            resolved = normalized;
            error = null;
            return true;
        }

        resolved = fallback;
        error = $"Invalid value '{value}'. Expected one of: {string.Join(", ", allowed.OrderBy(static x => x))}.";
        return false;
    }

    private static bool IsTimeoutOrPartial(JsonElement value)
    {
        if (value.TryGetProperty("partialResults", out var partial) && partial.ValueKind != JsonValueKind.Null)
            return true;

        if (value.TryGetProperty("reason", out var reason)
            && reason.ValueKind == JsonValueKind.String
            && string.Equals(reason.GetString(), "timeout", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return value.TryGetProperty("success", out var success)
               && success.ValueKind == JsonValueKind.False
               && value.TryGetProperty("reason", out var failureReason)
               && failureReason.ValueKind == JsonValueKind.String
               && !string.Equals(failureReason.GetString(), "element-not-found", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<string> SelectReferences(
        ObjectRepositorySnapshot snapshot,
        ObjectVerificationOptions options,
        out string? error)
    {
        error = null;
        var hasPage = !string.IsNullOrWhiteSpace(options.PageId);
        var hasRef = !string.IsNullOrWhiteSpace(options.ObjectRef);

        if (hasPage && hasRef)
        {
            error = "--page and --ref are mutually exclusive.";
            return Array.Empty<string>();
        }

        if (hasRef)
            return [options.ObjectRef!.Trim()];

        if (hasPage)
        {
            var pageId = options.PageId!.Trim();
            if (!snapshot.Pages.TryGetValue(pageId, out var page))
            {
                error = $"Page '{pageId}' was not found in the object repository.";
                return Array.Empty<string>();
            }

            if (!string.Equals(page.State, "active", StringComparison.Ordinal))
            {
                error = $"Page '{pageId}' is not active.";
                return Array.Empty<string>();
            }

            return page.Elements?.Keys
                       .OrderBy(static key => key, StringComparer.Ordinal)
                       .Select(elementId => $"{pageId}.{elementId}")
                       .ToArray()
                   ?? Array.Empty<string>();
        }

        return snapshot.Pages
            .Where(pair => string.Equals(pair.Value.State, "active", StringComparison.Ordinal))
            .OrderBy(static pair => pair.Key, StringComparer.Ordinal)
            .SelectMany(pair => pair.Value.Elements?.Keys
                                  .OrderBy(static key => key, StringComparer.Ordinal)
                                  .Select(elementId => $"{pair.Key}.{elementId}")
                              ?? Array.Empty<string>())
            .ToArray();
    }

    private static ObjectVerificationResult BuildSummary(
        ObjectRepositoryValidationResult validation,
        IReadOnlyList<ObjectVerificationItemResult> items,
        int passed,
        int missing,
        int ambiguous,
        int fragile,
        int failed,
        Stopwatch stopwatch,
        int exitCode,
        bool success,
        string? error = null) =>
        new()
        {
            Success = success,
            ExitCode = exitCode,
            Error = error is null ? null : SecretRedactor.Redact(error),
            RepositoryPath = validation.RepositoryPath,
            RepositoryId = validation.Snapshot!.Manifest.RepositoryId,
            RepositorySha256 = validation.Snapshot.AggregateSha256,
            Total = items.Count,
            Passed = passed,
            Missing = missing,
            Ambiguous = ambiguous,
            Fragile = fragile,
            Failed = failed,
            DurationMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
            Items = items
        };

    private static ObjectVerificationResult Failure(
        string message,
        int exitCode,
        Stopwatch stopwatch,
        string? repositoryPath = null,
        string? repositoryId = null,
        string? repositorySha256 = null) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = SecretRedactor.Redact(message),
            RepositoryPath = repositoryPath,
            RepositoryId = repositoryId,
            RepositorySha256 = repositorySha256,
            DurationMilliseconds = stopwatch.Elapsed.TotalMilliseconds
        };
}
