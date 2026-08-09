using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectCaptureService
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
    private readonly ObjectRepositoryReader _repositoryReader;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly IDriverUiClient _uiClient;
    private readonly ObjectCandidateGenerator _candidateGenerator;
    private readonly ObjectArtifactWriter _artifactWriter;
    private readonly IWorkspaceManager _workspace;
    private readonly ILogger<ObjectCaptureService> _logger;

    public ObjectCaptureService(
        IOptions<AgentOptions> options,
        ObjectRepositoryReader repositoryReader,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        IDriverUiClient uiClient,
        ObjectCandidateGenerator candidateGenerator,
        ObjectArtifactWriter artifactWriter,
        IWorkspaceManager workspace,
        ILogger<ObjectCaptureService> logger)
    {
        _options = options.Value;
        _repositoryReader = repositoryReader;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _uiClient = uiClient;
        _candidateGenerator = candidateGenerator;
        _artifactWriter = artifactWriter;
        _workspace = workspace;
        _logger = logger;
    }

    public async Task<ObjectCaptureResult> CaptureAsync(
        string repositoryPath,
        string pageId,
        string pageName,
        ObjectCaptureOptions captureOptions,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(captureOptions);

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.ObjectRepository);
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
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

        var validation = _repositoryReader.Read(repositoryPath);
        if (!validation.IsValid || validation.Snapshot is null)
        {
            return Failure(
                validation.Errors.FirstOrDefault() ?? "Object repository validation failed.",
                ExitCodes.SuiteOrWorkspace);
        }

        var view = NormalizeOption(captureOptions.View, AllowedViews, "control");
        var root = NormalizeOption(captureOptions.Root, AllowedRoots, "activeWindow");
        var maxDepth = captureOptions.MaxDepth ?? 8;
        var maxChildren = captureOptions.MaxChildren ?? 200;
        var includeOffscreen = captureOptions.IncludeOffscreen ?? false;

        if (maxDepth is < 0 or > 20)
            return Failure("maxDepth must be between 0 and 20.", ExitCodes.UsageOrConfiguration);

        if (maxChildren is < 1 or > 1000)
            return Failure("maxChildren must be between 1 and 1000.", ExitCodes.UsageOrConfiguration);

        DriverConnection connection;
        OperationsCatalogDto catalog;
        try
        {
            connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            var status = await _catalogClient.GetStatusAsync(connection, cancellationToken).ConfigureAwait(false);
            if (!status.Ready)
                return Failure("Driver /status reported ready=false.", ExitCodes.DriverUnavailable);

            catalog = await _catalogClient.GetOperationsAsync(connection, cancellationToken).ConfigureAwait(false);
            CatalogCompatibility.Validate(catalog, _options.Driver.ExpectedCatalogSchemaVersion);
            DiagnosticCatalogValidator.RequireDiagnosticOperation(catalog, "dumpuia");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return Failure("Capture was cancelled.", ExitCodes.Cancelled);
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
            ["includeOffscreen"] = JsonSerializer.SerializeToElement(includeOffscreen),
            ["includePath"] = JsonSerializer.SerializeToElement(true),
            ["timeoutMs"] = JsonSerializer.SerializeToElement(_options.ObjectRepository.DiagnosticTimeoutMilliseconds)
        };

        var step = new PlanStep
        {
            Id = "capture-dumpuia",
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
            return Failure("Capture was cancelled.", ExitCodes.Cancelled);
        }
        catch (UiExecutionException ex)
        {
            return Failure(
                SecretRedactor.Redact(ex.Message),
                ex.Classification == UiFailureClassification.Cancelled
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

        var nodes = nodesElement.EnumerateArray().ToList();
        if (nodes.Count == 0)
            return Failure("dumpuia returned zero usable nodes.", ExitCodes.ExecutionFailure);

        var captureId = DeterministicPlanRunner.GenerateRunId();
        var repositoryRoot = Path.GetDirectoryName(_workspace.ResolveSafePath(repositoryPath))
            ?? throw new InvalidOperationException("Repository path has no directory.");

        var captureFileName = $"{captureId}.capture.json";
        var candidateFileName = $"{captureId}.page.json";
        var captureFullPath = Path.Combine(repositoryRoot, "captures", pageId, captureFileName);
        var candidateFullPath = Path.Combine(repositoryRoot, "candidates", pageId, candidateFileName);
        var captureRelative = ToDisplayRelative(captureFullPath);
        var candidateRelative = ToDisplayRelative(candidateFullPath);

        if (File.Exists(captureFullPath) || File.Exists(candidateFullPath))
            return Failure("Capture artifacts already exist and must not be overwritten.", ExitCodes.ExecutionFailure);

        var candidatePage = _candidateGenerator.Generate(nodes, pageId, pageName, captureId);
        var unresolvedCount = candidatePage.Unresolved is JsonElement unresolved
                              && unresolved.ValueKind == JsonValueKind.Array
            ? unresolved.GetArrayLength()
            : 0;

        var captureDocument = new Dictionary<string, object?>
        {
            ["captureId"] = captureId,
            ["pageId"] = pageId,
            ["pageName"] = pageName,
            ["capturedAtUtc"] = DateTimeOffset.UtcNow.ToString("O"),
            ["repositoryId"] = validation.Snapshot.Manifest.RepositoryId,
            ["repositorySha256"] = validation.Snapshot.AggregateSha256,
            ["request"] = new
            {
                view,
                root,
                maxDepth,
                maxChildren,
                includeOffscreen,
                includePath = true,
                timeoutMs = _options.ObjectRepository.DiagnosticTimeoutMilliseconds
            },
            ["response"] = JsonSerializer.Deserialize<object>(value.GetRawText())
        };

        var captureJson = JsonSerializer.Serialize(captureDocument, JsonOptions);
        var candidateJson = JsonSerializer.Serialize(candidatePage, JsonOptions);

        try
        {
            _artifactWriter.WriteAtomic(captureFullPath, captureJson);
            try
            {
                _artifactWriter.WriteAtomic(candidateFullPath, candidateJson);
            }
            catch
            {
                TryDelete(captureFullPath);
                throw;
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Failure($"Failed to write capture artifacts: {ex.Message}", ExitCodes.ExecutionFailure);
        }

        _logger.LogInformation(
            "Capture {CaptureId} for page {PageId} wrote {ElementCount} elements and {UnresolvedCount} unresolved nodes.",
            captureId,
            pageId,
            candidatePage.Elements?.Count ?? 0,
            unresolvedCount);

        return new ObjectCaptureResult
        {
            Success = true,
            ExitCode = ExitCodes.Success,
            RepositoryPath = validation.RepositoryPath,
            RepositoryId = validation.Snapshot.Manifest.RepositoryId,
            RepositorySha256 = validation.Snapshot.AggregateSha256,
            CaptureId = captureId,
            PageId = pageId,
            CaptureFilePath = captureRelative,
            CandidateFilePath = candidateRelative,
            CaptureSha256 = ComputeSha256(captureJson),
            CandidateSha256 = ComputeSha256(candidateJson),
            NodeCount = nodes.Count,
            ElementCount = candidatePage.Elements?.Count ?? 0,
            UnresolvedCount = unresolvedCount
        };
    }

    private static bool IsFailureResponse(JsonElement value, out string reason)
    {
        reason = "unknown";
        if (value.TryGetProperty("success", out var success) && success.ValueKind == JsonValueKind.False)
        {
            reason = value.TryGetProperty("reason", out var reasonElement) && reasonElement.ValueKind == JsonValueKind.String
                ? reasonElement.GetString() ?? "unknown"
                : "unknown";
            return true;
        }

        if (value.TryGetProperty("partialResults", out var partial) && partial.ValueKind != JsonValueKind.Null)
        {
            reason = value.TryGetProperty("reason", out var reasonElement) && reasonElement.ValueKind == JsonValueKind.String
                ? reasonElement.GetString() ?? "partial"
                : "partial";
            return true;
        }

        if (value.TryGetProperty("reason", out var reasonProp)
            && reasonProp.ValueKind == JsonValueKind.String
            && string.Equals(reasonProp.GetString(), "timeout", StringComparison.OrdinalIgnoreCase))
        {
            reason = "timeout";
            return true;
        }

        if (value.TryGetProperty("reason", out var noRoot)
            && noRoot.ValueKind == JsonValueKind.String
            && string.Equals(noRoot.GetString(), "no-root", StringComparison.OrdinalIgnoreCase))
        {
            reason = "no-root";
            return true;
        }

        return false;
    }

    private static string NormalizeOption(string? value, HashSet<string> allowed, string fallback)
    {
        if (string.IsNullOrWhiteSpace(value))
            return fallback;

        var normalized = value.Trim();
        return allowed.Contains(normalized) ? normalized : fallback;
    }

    private static string ComputeSha256(string content) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(content))).ToLowerInvariant();

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // best-effort cleanup
        }
    }

    private string ToDisplayRelative(string fullPath)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var relative = Path.GetRelativePath(root, normalized);
        return relative.Replace(Path.DirectorySeparatorChar, '/');
    }

    private static ObjectCaptureResult Failure(string message, int exitCode) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = SecretRedactor.Redact(message)
        };
}
