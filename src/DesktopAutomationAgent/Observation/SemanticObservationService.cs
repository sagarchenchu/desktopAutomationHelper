using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Observation;

public sealed class SemanticObservationService : ISemanticObservationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

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
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly IDriverUiClient _uiClient;
    private readonly SemanticSnapshotNormalizer _normalizer;
    private readonly IWorkspaceManager _workspace;
    private readonly ILogger<SemanticObservationService> _logger;

    public SemanticObservationService(
        IOptions<AgentOptions> options,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        IDriverUiClient uiClient,
        SemanticSnapshotNormalizer normalizer,
        IWorkspaceManager workspace,
        ILogger<SemanticObservationService> logger)
    {
        _options = options.Value;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _uiClient = uiClient;
        _normalizer = normalizer;
        _workspace = workspace;
        _logger = logger;
    }

    public async Task<ObservationResult> ObserveAsync(
        ObservationOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(options);

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Workspace);
            _workspace.EnsureInitialized();
        }
        catch (AgentConfigurationException ex)
        {
            return Failure(ex.Message, ExitCodes.UsageOrConfiguration);
        }
        catch (WorkspaceException ex)
        {
            return Failure(ex.Message, ExitCodes.SuiteOrWorkspace);
        }

        if (!TryResolveOption(options.View, AllowedViews, "control", out var view, out var viewError))
            return Failure(viewError!, ExitCodes.UsageOrConfiguration);

        if (!TryResolveOption(options.Root, AllowedRoots, "activeWindow", out var root, out var rootError))
            return Failure(rootError!, ExitCodes.UsageOrConfiguration);

        var maxDepth = options.MaxDepth ?? 8;
        var maxChildren = options.MaxChildren ?? 200;
        var maxNodes = options.MaxNodes ?? SemanticSnapshotNormalizer.DefaultMaxNodes;
        var includeOffscreen = options.IncludeOffscreen ?? false;

        if (maxDepth is < 0 or > 20)
            return Failure("maxDepth must be between 0 and 20.", ExitCodes.UsageOrConfiguration);

        if (maxChildren is < 1 or > 1000)
            return Failure("maxChildren must be between 1 and 1000.", ExitCodes.UsageOrConfiguration);

        if (maxNodes is < 1 or > 2000)
            return Failure("maxNodes must be between 1 and 2000.", ExitCodes.UsageOrConfiguration);

        DriverConnection connection;
        OperationsCatalogDto catalog;
        StatusValueDto status;
        try
        {
            connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            status = await _catalogClient.GetStatusAsync(connection, cancellationToken).ConfigureAwait(false);
            if (!status.Ready)
                return Failure("Driver /status reported ready=false.", ExitCodes.DriverUnavailable);

            catalog = await _catalogClient.GetOperationsAsync(connection, cancellationToken).ConfigureAwait(false);
            CatalogCompatibility.Validate(catalog, _options.Driver.ExpectedCatalogSchemaVersion);
            DiagnosticCatalogValidator.RequireDiagnosticOperation(catalog, "dumpuia");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return Failure("Observation was cancelled.", ExitCodes.Cancelled);
        }
        catch (DriverConnectionException ex)
        {
            return Failure(ex.Message, ExitCodes.DriverUnavailable);
        }
        catch (DriverCatalogException ex)
        {
            return Failure(ex.Message, ExitCodes.AuthOrCatalog);
        }

        var arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
        {
            ["view"] = JsonSerializer.SerializeToElement(view),
            ["root"] = JsonSerializer.SerializeToElement(root),
            ["maxDepth"] = JsonSerializer.SerializeToElement(maxDepth),
            ["maxChildren"] = JsonSerializer.SerializeToElement(maxChildren),
            ["maxNodes"] = JsonSerializer.SerializeToElement(maxNodes),
            ["includeOffscreen"] = JsonSerializer.SerializeToElement(includeOffscreen),
            ["includePath"] = JsonSerializer.SerializeToElement(false),
            ["includeBounds"] = JsonSerializer.SerializeToElement(options.IncludeBounds),
            ["includeHelpText"] = JsonSerializer.SerializeToElement(options.IncludeHelpText),
            ["includeFocusState"] = JsonSerializer.SerializeToElement(options.IncludeFocusState),
            ["includeRuntimeId"] = JsonSerializer.SerializeToElement(options.IncludeRuntimeId),
            ["includeTreeLinks"] = JsonSerializer.SerializeToElement(true),
            ["includeWindowState"] = JsonSerializer.SerializeToElement(true),
            ["timeoutMs"] = JsonSerializer.SerializeToElement(
                Math.Clamp(_options.ObjectRepository.DiagnosticTimeoutMilliseconds, 500, 15000))
        };

        // Observe-only: exactly one diagnostic dumpuia POST /ui. No action operations.
        var step = new PlanStep
        {
            Id = "observe-dumpuia",
            Operation = "dumpuia",
            Arguments = arguments
        };

        UiExecutionResponse response;
        try
        {
            response = await _uiClient.ExecuteStepAsync(connection, step, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return Failure("Observation was cancelled.", ExitCodes.Cancelled);
        }
        catch (UiExecutionException ex)
        {
            return Failure(
                SecretRedactor.Redact(ex.Message),
                ex.Classification is UiFailureClassification.Cancelled
                    ? ExitCodes.Cancelled
                    : ExitCodes.ExecutionFailure);
        }

        if (response.Value is not JsonElement value || value.ValueKind != JsonValueKind.Object)
            return Failure("dumpuia response did not include a value object.", ExitCodes.ExecutionFailure);

        if (IsFailureResponse(value, out var failureReason))
        {
            return Failure(
                $"dumpuia failed ({failureReason}).",
                failureReason == "cancelled" ? ExitCodes.Cancelled : ExitCodes.ExecutionFailure);
        }

        if (!value.TryGetProperty("nodes", out var nodesElement) || nodesElement.ValueKind != JsonValueKind.Array)
            return Failure("dumpuia response did not include nodes.", ExitCodes.ExecutionFailure);

        if (nodesElement.GetArrayLength() == 0)
            return Failure("dumpuia returned zero nodes.", ExitCodes.ExecutionFailure);

        var observationId = DeterministicPlanRunner.GenerateRunId();
        SemanticSnapshot snapshot;
        try
        {
            snapshot = _normalizer.Normalize(
                value,
                options with { MaxNodes = maxNodes },
                observationId,
                DateTimeOffset.UtcNow,
                catalog.DriverVersion ?? status.Build?.Version ?? string.Empty);
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException)
        {
            return Failure(SecretRedactor.Redact(ex.Message), ExitCodes.ExecutionFailure);
        }

        // Drop runtime IDs from persisted artifact while keeping them on the in-memory result.
        var persistSnapshot = CloneWithoutRuntimeIds(snapshot);
        string? relativePath = null;

        try
        {
            relativePath = await WriteArtifactAsync(observationId, persistSnapshot, options.OutputPath)
                .ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or System.Security.SecurityException
                                       or RepositoryPathException or WorkspaceException or InvalidOperationException)
        {
            return Failure(SecretRedactor.Redact(ex.Message), ExitCodes.SuiteOrWorkspace);
        }

        _logger.LogInformation(
            "Semantic observation {ObservationId} captured {NodeCount} nodes truncated={Truncated}.",
            observationId,
            snapshot.Nodes.Count,
            snapshot.Truncated);

        return new ObservationResult
        {
            Success = true,
            ExitCode = ExitCodes.Success,
            ObservationId = observationId,
            SnapshotFilePath = relativePath,
            SnapshotHash = snapshot.SnapshotHash,
            NodeCount = snapshot.Nodes.Count,
            Truncated = snapshot.Truncated,
            Snapshot = snapshot
        };
    }

    private async Task<string> WriteArtifactAsync(
        string observationId,
        SemanticSnapshot snapshot,
        string? outputPath)
    {
        string fullPath;
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            fullPath = Path.Combine(
                _workspace.RootPath,
                "observations",
                observationId,
                "snapshot.json");
        }
        else
        {
            fullPath = _workspace.ResolveSafePath(outputPath);
        }

        ObjectRepositoryPathSafety.EnsureWritablePathInside(fullPath, _workspace.RootPath);

        if (File.Exists(fullPath))
            throw new InvalidOperationException("Observation artifact already exists and must not be overwritten.");

        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("Observation path has no directory.");
        Directory.CreateDirectory(directory);

        var json = JsonSerializer.Serialize(snapshot, JsonOptions);
        json = SecretRedactor.Redact(json);
        await File.WriteAllTextAsync(fullPath, json.EndsWith('\n') ? json : json + "\n").ConfigureAwait(false);

        return ToDisplayRelative(fullPath);
    }

    private static SemanticSnapshot CloneWithoutRuntimeIds(SemanticSnapshot snapshot) =>
        new()
        {
            SnapshotSchemaVersion = snapshot.SnapshotSchemaVersion,
            SnapshotId = snapshot.SnapshotId,
            CapturedAtUtc = snapshot.CapturedAtUtc,
            DriverVersion = snapshot.DriverVersion,
            SnapshotHash = snapshot.SnapshotHash,
            Application = snapshot.Application,
            WindowsSession = snapshot.WindowsSession,
            Window = snapshot.Window,
            Nodes = snapshot.Nodes.Select(static n => new SemanticNode
            {
                Ref = n.Ref,
                ParentRef = n.ParentRef,
                Ordinal = n.Ordinal,
                Depth = n.Depth,
                ControlType = n.ControlType,
                Role = n.Role,
                Name = n.Name,
                AutomationId = n.AutomationId,
                ClassName = n.ClassName,
                HelpText = n.HelpText,
                Enabled = n.Enabled,
                Offscreen = n.Offscreen,
                Focusable = n.Focusable,
                HasKeyboardFocus = n.HasKeyboardFocus,
                Patterns = n.Patterns,
                Capabilities = n.Capabilities,
                Bounds = n.Bounds
            }).ToList(),
            Truncated = snapshot.Truncated,
            Warnings = snapshot.Warnings,
            Ordering = snapshot.Ordering
        };

    private string ToDisplayRelative(string fullPath)
    {
        var root = _workspace.RootPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                   + Path.DirectorySeparatorChar;
        if (fullPath.StartsWith(root, OperatingSystem.IsWindows()
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal))
        {
            return fullPath[root.Length..].Replace('\\', '/');
        }

        return fullPath.Replace('\\', '/');
    }

    private static bool TryResolveOption(
        string? value,
        HashSet<string> allowed,
        string fallback,
        out string resolved,
        out string? error)
    {
        resolved = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        if (!allowed.Contains(resolved))
        {
            error = $"Unsupported value '{resolved}'. Allowed: {string.Join(", ", allowed.OrderBy(static x => x))}.";
            return false;
        }

        error = null;
        return true;
    }

    private static bool IsFailureResponse(JsonElement value, out string reason)
    {
        reason = "unknown";
        if (value.TryGetProperty("success", out var success)
            && success.ValueKind == JsonValueKind.False)
        {
            if (value.TryGetProperty("reason", out var reasonElement)
                && reasonElement.ValueKind == JsonValueKind.String)
            {
                reason = reasonElement.GetString() ?? reason;
            }

            return true;
        }

        return false;
    }

    private static ObservationResult Failure(string message, int exitCode) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = SecretRedactor.Redact(message)
        };
}
