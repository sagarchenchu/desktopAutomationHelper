using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Jira;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Readiness;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Mcp;

public sealed class DesktopMcpFacade : IDesktopMcpFacade
{
    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;
    private readonly IAgentReadinessService _readiness;
    private readonly ISemanticObservationService _observation;
    private readonly IExplorationCoordinator _exploration;
    private readonly KnowledgeService _knowledge;
    private readonly PlanManifestReader _planReader;
    private readonly PlanObjectRepositoryIntegrator _planIntegrator;
    private readonly BddCompilationService _compiler;
    private readonly JiraSuiteOrchestrator _jira;
    private readonly IDeterministicPlanRunner _runner;
    private readonly McpOpaqueRefStore _refs;

    public DesktopMcpFacade(
        IOptions<AgentOptions> options,
        IWorkspaceManager workspace,
        IAgentReadinessService readiness,
        ISemanticObservationService observation,
        IExplorationCoordinator exploration,
        KnowledgeService knowledge,
        PlanManifestReader planReader,
        PlanObjectRepositoryIntegrator planIntegrator,
        BddCompilationService compiler,
        JiraSuiteOrchestrator jira,
        IDeterministicPlanRunner runner,
        McpOpaqueRefStore refs)
    {
        _options = options.Value;
        _workspace = workspace;
        _readiness = readiness;
        _observation = observation;
        _exploration = exploration;
        _knowledge = knowledge;
        _planReader = planReader;
        _planIntegrator = planIntegrator;
        _compiler = compiler;
        _jira = jira;
        _runner = runner;
        _refs = refs;
    }

    public async Task<object> DoctorAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var report = await _readiness.RunDoctorAsync(cancellationToken).ConfigureAwait(false);
        return CompactDoctor(report);
    }

    public async Task<object> ObserveAsync(
        string? view,
        string? root,
        int? maxDepth,
        int? maxChildren,
        int? maxNodes,
        bool includeOffscreen,
        bool includeBounds,
        bool includeHelpText,
        bool? includeFocusState,
        bool? actionableOnly,
        string? outputPath,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var depth = ClampOptional(maxDepth, McpResponseBudget.DefaultMaxDepth, 0, McpResponseBudget.AbsoluteMaxDepth);
        var children = ClampOptional(maxChildren, McpResponseBudget.DefaultMaxChildren, 1, McpResponseBudget.AbsoluteMaxChildren);
        var nodes = ClampOptional(maxNodes, McpResponseBudget.DefaultMaxNodes, 1, McpResponseBudget.AbsoluteMaxNodes);
        var safeOutput = McpPathGuard.OptionalWorkspacePath(_workspace, outputPath, "outputPath");

        var result = await _observation.ObserveAsync(
                new ObservationOptions
                {
                    View = string.IsNullOrWhiteSpace(view) ? "control" : view,
                    Root = string.IsNullOrWhiteSpace(root) ? "activeWindow" : root,
                    MaxDepth = depth,
                    MaxChildren = children,
                    MaxNodes = nodes,
                    IncludeOffscreen = includeOffscreen,
                    IncludeBounds = includeBounds,
                    IncludeHelpText = includeHelpText,
                    IncludeFocusState = includeFocusState ?? true,
                    ActionableOnly = actionableOnly ?? true,
                    OutputPath = safeOutput,
                    IncludeRuntimeId = false
                },
                cancellationToken)
            .ConfigureAwait(false);

        if (!result.Success)
            throw MapObservationFailure(result);

        if (result.Snapshot is not null)
            _refs.ReplaceSnapshot(result.Snapshot);

        return CompactObservation(result);
    }

    public async Task<object> ExploreAsync(string policyPath, bool dryRun, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var safePolicy = McpPathGuard.RequireWorkspacePath(_workspace, policyPath, "policyPath");

        var result = await _exploration.ExploreAsync(
                new ExplorationRequest
                {
                    PolicyPath = safePolicy,
                    DryRun = dryRun
                },
                cancellationToken)
            .ConfigureAwait(false);

        if (!result.Success)
            throw MapExplorationFailure(result);

        return CompactExploration(result, dryRun);
    }

    public object SearchKnowledge(string knowledgePath, string query, int? limit)
    {
        if (string.IsNullOrWhiteSpace(query) || query.Length > 256)
            throw new McpToolException(McpErrorCode.Validation, "query is required and must be at most 256 characters.");

        var safe = McpPathGuard.RequireWorkspacePath(_workspace, knowledgePath, "knowledgePath");
        var hits = _knowledge.Search(safe, query);
        var take = Math.Clamp(limit ?? McpResponseBudget.DefaultSearchLimit, 1, McpResponseBudget.AbsoluteSearchLimit);
        return new
        {
            success = true,
            category = "observation",
            query,
            count = Math.Min(hits.Count, take),
            hits = hits.Take(take).Select(h => new
            {
                h.Kind,
                h.Id,
                h.Score,
                snippet = Truncate(h.Text, 240)
            })
        };
    }

    public object ValidatePlan(string planPath)
    {
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Runner);
        var safe = McpPathGuard.RequireWorkspacePath(_workspace, planPath, "planPath");
        var result = _planIntegrator.Integrate(_planReader.Read(safe));
        return new
        {
            success = result.IsValid,
            category = "validation",
            planPath = ToRelative(result.PlanPath),
            result.PlanId,
            result.Name,
            result.StepCount,
            result.OnFailureStepCount,
            result.TotalStepCount,
            result.Sha256,
            objectRepositoryPath = ToRelative(result.ObjectRepositoryPath),
            result.ObjectRepositoryId,
            result.ObjectRepositorySha256,
            result.ResolvedObjectReferences,
            errors = (result.Errors ?? []).Select(SecretRedactor.Redact).ToArray(),
            warnings = (result.Warnings ?? []).Select(SecretRedactor.Redact).ToArray()
        };
    }

    public object CompileKey(
        string key,
        string? file,
        bool fromJira,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        bool explain)
    {
        if (fromJira)
            throw new McpToolException(McpErrorCode.Validation, "Use CompileKeyFromJiraAsync for --from-jira compiles.");

        var result = _compiler.CompileKey(new CompileKeyRequest
        {
            JiraKey = key,
            TestCasePath = McpPathGuard.OptionalWorkspacePath(_workspace, file, "file"),
            RegistryPath = McpPathGuard.OptionalWorkspacePath(_workspace, registry, "registry"),
            ObjectRepositoryPath = McpPathGuard.OptionalWorkspacePath(_workspace, objectRepository, "objectRepository"),
            KnowledgePath = McpPathGuard.OptionalWorkspacePath(_workspace, knowledge, "knowledge"),
            CatalogPath = McpPathGuard.OptionalWorkspacePath(_workspace, catalog, "catalog"),
            Explain = explain
        });

        return CompactCompilation(result);
    }

    public async Task<object> CompileKeyFromJiraAsync(
        string key,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        bool explain,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = await _jira.CompileKeyFromJiraAsync(
                new CompileKeyFromJiraRequest
                {
                    JiraKey = key,
                    RegistryPath = McpPathGuard.OptionalWorkspacePath(_workspace, registry, "registry"),
                    ObjectRepositoryPath = McpPathGuard.OptionalWorkspacePath(_workspace, objectRepository, "objectRepository"),
                    KnowledgePath = McpPathGuard.OptionalWorkspacePath(_workspace, knowledge, "knowledge"),
                    CatalogPath = McpPathGuard.OptionalWorkspacePath(_workspace, catalog, "catalog"),
                    Explain = explain
                },
                cancellationToken)
            .ConfigureAwait(false);
        return CompactCompilation(result);
    }

    public async Task<object> RunPlanAsync(string planPath, bool dryRun, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Runner);
        var safe = McpPathGuard.RequireWorkspacePath(_workspace, planPath, "planPath");
        var report = await _runner.RunAsync(safe, dryRun, cancellationToken).ConfigureAwait(false);
        return CompactRunReport(report);
    }

    public async Task<object> RunKeyAsync(
        string key,
        bool dryRun,
        bool fromJira,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = await _jira.RunKeyAsync(
                new JiraRunKeyRequest
                {
                    JiraKey = key,
                    DryRun = dryRun,
                    FromJira = fromJira,
                    RegistryPath = McpPathGuard.OptionalWorkspacePath(_workspace, registry, "registry"),
                    ObjectRepositoryPath = McpPathGuard.OptionalWorkspacePath(_workspace, objectRepository, "objectRepository"),
                    KnowledgePath = McpPathGuard.OptionalWorkspacePath(_workspace, knowledge, "knowledge"),
                    CatalogPath = McpPathGuard.OptionalWorkspacePath(_workspace, catalog, "catalog")
                },
                cancellationToken)
            .ConfigureAwait(false);

        return new
        {
            success = result.Success,
            category = dryRun ? "validation" : "execution",
            result.ExitCode,
            result.JiraKey,
            result.CompilationFingerprint,
            result.PlanHash,
            result.RunId,
            result.Result,
            result.DryRun,
            result.DurationMilliseconds,
            runReport = result.RunReport is null ? null : CompactRunReport(result.RunReport),
            errors = result.Errors.Select(SecretRedactor.Redact).ToArray()
        };
    }

    public async Task<object> RunSuiteAsync(
        string suitePath,
        bool dryRun,
        bool fromJira,
        bool continueOnFailure,
        int? maxFailures,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var safe = McpPathGuard.RequireWorkspacePath(_workspace, suitePath, "suitePath");
        var result = await _jira.RunSuiteAsync(
                new JiraSuiteRunRequest
                {
                    SuiteFile = safe,
                    DryRun = dryRun,
                    FromJira = fromJira,
                    ContinueOnFailure = continueOnFailure,
                    MaxFailures = maxFailures ?? 0,
                    RegistryPath = McpPathGuard.OptionalWorkspacePath(_workspace, registry, "registry"),
                    ObjectRepositoryPath = McpPathGuard.OptionalWorkspacePath(_workspace, objectRepository, "objectRepository"),
                    KnowledgePath = McpPathGuard.OptionalWorkspacePath(_workspace, knowledge, "knowledge"),
                    CatalogPath = McpPathGuard.OptionalWorkspacePath(_workspace, catalog, "catalog")
                },
                cancellationToken)
            .ConfigureAwait(false);

        return new
        {
            success = result.Success,
            category = dryRun ? "validation" : "execution",
            result.ExitCode,
            suiteFile = ToRelative(result.SuiteFile),
            result.SuiteName,
            result.SuiteRunId,
            result.DryRun,
            result.ContinueOnFailure,
            result.StopReason,
            result.DurationMilliseconds,
            keys = result.Keys.Select(k => new
            {
                k.JiraKey,
                k.CompilationFingerprint,
                k.PlanHash,
                k.RunId,
                k.Result,
                k.DurationMilliseconds,
                k.ExitCode,
                errors = k.Errors.Select(SecretRedactor.Redact).ToArray()
            }),
            errors = result.Errors.Select(SecretRedactor.Redact).ToArray()
        };
    }

    public object GetRunSummary(string runId)
    {
        if (string.IsNullOrWhiteSpace(runId) || runId.Length > 128)
            throw new McpToolException(McpErrorCode.Validation, "runId is required and must be at most 128 characters.");

        if (runId.Contains("..", StringComparison.Ordinal) || runId.Contains('/') || runId.Contains('\\') || runId.Contains(':'))
            throw new McpToolException(McpErrorCode.PathTraversal, "runId must be a single path segment.");

        var runDir = McpPathGuard.RequireWorkspacePath(_workspace, Path.Combine("runs", runId), "runId");
        var runJson = Path.Combine(runDir, "run.json");
        var suiteJson = Path.Combine(runDir, "suite-run.json");

        if (File.Exists(runJson))
        {
            McpPathGuard.EnsureReadablePathInside(runJson, _workspace.RootPath);
            var report = JsonSerializer.Deserialize<RunReport>(File.ReadAllText(runJson), McpJson.Options)
                ?? throw new McpToolException(McpErrorCode.Validation, "run.json could not be parsed.");
            return CompactRunReport(RunArtifactWriter.RedactReport(report));
        }

        if (File.Exists(suiteJson))
        {
            McpPathGuard.EnsureReadablePathInside(suiteJson, _workspace.RootPath);
            using var doc = JsonDocument.Parse(File.ReadAllText(suiteJson));
            var redacted = SecretRedactor.Redact(doc.RootElement.GetRawText());
            using var cleaned = JsonDocument.Parse(redacted);
            return new
            {
                success = true,
                category = "observation",
                kind = "suite-run",
                runId,
                summary = cleaned.RootElement.Clone()
            };
        }

        throw new McpToolException(McpErrorCode.Validation, $"No run summary found for runId '{runId}'.");
    }

    public object ResolveOpaqueRef(string snapshotId, string nodeRef)
    {
        var node = _refs.Resolve(snapshotId, nodeRef);
        return new
        {
            success = true,
            category = "observation",
            snapshotId,
            nodeRef,
            node = new
            {
                node.Ref,
                node.ParentRef,
                node.ControlType,
                node.Role,
                node.Name,
                node.AutomationId,
                node.Enabled,
                node.Capabilities
            }
        };
    }

    private static object CompactDoctor(ReadinessReport report) =>
        new
        {
            success = report.Success,
            category = "observation",
            report.ExitCode,
            workspaceRoot = report.WorkspaceRoot,
            // Username omitted from MCP by default (PII / same-user assumption doc elsewhere).
            driverBaseUrl = report.DriverBaseUrl,
            report.DiscoveryMethod,
            report.DriverVersion,
            report.CatalogSchemaVersion,
            report.OperationCount,
            checks = report.Checks.Select(c => new
            {
                c.Name,
                c.Status,
                detail = SecretRedactor.Redact(c.Detail)
            }),
            errors = report.Errors.Select(SecretRedactor.Redact).ToArray()
        };

    private object CompactObservation(ObservationResult result)
    {
        var snapshot = result.Snapshot;
        var nodes = snapshot?.Nodes
            .Take(McpResponseBudget.AbsoluteMaxNodes)
            .Select(n => new
            {
                n.Ref,
                n.ParentRef,
                n.Ordinal,
                n.Depth,
                n.ControlType,
                n.Role,
                name = Truncate(n.Name, 120),
                n.AutomationId,
                n.Enabled,
                n.Offscreen,
                n.Focusable,
                n.HasKeyboardFocus,
                capabilities = n.Capabilities?.Take(12).ToArray(),
                patterns = n.Patterns?.Take(12).ToArray()
            })
            .ToArray();

        return new
        {
            success = result.Success,
            category = "observation",
            result.ExitCode,
            result.ObservationId,
            snapshotId = snapshot?.SnapshotId,
            snapshotHash = result.SnapshotHash,
            nodeCount = result.NodeCount,
            truncated = result.Truncated || (snapshot?.Nodes.Count ?? 0) > McpResponseBudget.AbsoluteMaxNodes,
            artifactPath = ToRelative(result.SnapshotFilePath),
            refLifetimeMinutes = (int)_refs.Lifetime.TotalMinutes,
            nodes,
            warnings = snapshot?.Warnings?.Select(SecretRedactor.Redact).Take(20).ToArray()
        };
    }

    private object CompactExploration(ExplorationResult result, bool dryRun) =>
        new
        {
            success = result.Success,
            category = dryRun ? "observation" : "navigation",
            result.ExitCode,
            result.ExplorationId,
            artifactDirectory = ToRelative(result.ArtifactDirectory),
            summary = result.Summary is null
                ? null
                : new
                {
                    result.Summary.Status,
                    result.Summary.Mode,
                    result.Summary.DryRun,
                    result.Summary.StatesVisited,
                    result.Summary.ActionsAttempted,
                    result.Summary.TransitionsRecorded,
                    result.Summary.MaxDepthReached,
                    result.Summary.StopReason
                },
            dryRunCandidates = result.DryRunCandidates.Take(50).Select(c => new
            {
                c.ActionKey,
                kind = c.Kind.ToString(),
                c.Category,
                c.DriverOperation,
                c.ElementKey,
                c.AutomationId,
                name = Truncate(c.Name, 80),
                c.ControlType,
                c.Depth,
                c.Priority
            })
        };

    private static object CompactCompilation(CompilationResult result) =>
        new
        {
            success = result.Success,
            category = "validation",
            result.ExitCode,
            result.JiraKey,
            result.Fingerprint,
            outputDirectory = result.OutputDirectory,
            planPath = result.PlanPath,
            compilationPath = result.CompilationPath,
            result.PlanSha256,
            result.CacheHit,
            explain = result.Explain?.Take(40).Select(e => new
            {
                e.Ordinal,
                e.Keyword,
                text = Truncate(e.Text, 160),
                e.MappingId,
                e.Outcome,
                details = e.Details.Take(8).Select(SecretRedactor.Redact).ToArray()
            }),
            errors = result.Errors.Select(SecretRedactor.Redact).ToArray(),
            warnings = result.Warnings.Select(SecretRedactor.Redact).ToArray()
        };

    private static object CompactRunReport(RunReport report)
    {
        var redacted = RunArtifactWriter.RedactReport(report);
        return new
        {
            success = redacted.ExitCode == ExitCodes.Success,
            category = redacted.DryRun ? "validation" : "execution",
            redacted.RunId,
            redacted.Status,
            redacted.ExitCode,
            planPath = redacted.PlanPath,
            redacted.PlanId,
            redacted.PlanName,
            redacted.PlanSha256,
            redacted.DryRun,
            redacted.DriverBaseUrl,
            redacted.DiscoveryMethod,
            redacted.DriverVersion,
            redacted.CatalogSchemaVersion,
            redacted.StartedAtUtc,
            redacted.FinishedAtUtc,
            redacted.DurationMilliseconds,
            stepCount = redacted.Steps.Count,
            onFailureStepCount = redacted.OnFailureSteps.Count,
            steps = redacted.Steps.Take(100).Select(s => new
            {
                s.Sequence,
                s.Id,
                s.Operation,
                s.Phase,
                s.Status,
                s.Success,
                s.Skipped,
                skipReason = SecretRedactor.Redact(s.SkipReason),
                s.HttpStatusCode
            }),
            failure = redacted.Failure is null
                ? null
                : new
                {
                    redacted.Failure.Classification,
                    message = SecretRedactor.Redact(redacted.Failure.Message)
                },
            redacted.ArtifactWriteStatus
        };
    }

    private static McpToolException MapObservationFailure(ObservationResult result)
    {
        var message = SecretRedactor.Redact(result.Error ?? "observation failed");
        var code = result.ExitCode switch
        {
            ExitCodes.DriverUnavailable => McpErrorCode.DriverUnavailable,
            ExitCodes.UsageOrConfiguration => McpErrorCode.Validation,
            ExitCodes.Cancelled => McpErrorCode.Cancelled,
            _ => McpErrorCode.Internal
        };
        return new McpToolException(code, message);
    }

    private static McpToolException MapExplorationFailure(ExplorationResult result)
    {
        var message = SecretRedactor.Redact(result.Error ?? "exploration failed");
        var lower = message.ToLowerInvariant();
        if (lower.Contains("deny") || lower.Contains("denied") || lower.Contains("policy"))
            return new McpToolException(McpErrorCode.Denied, message);
        if (result.ExitCode == ExitCodes.DriverUnavailable)
            return new McpToolException(McpErrorCode.DriverUnavailable, message);
        if (result.ExitCode == ExitCodes.Cancelled)
            return new McpToolException(McpErrorCode.Cancelled, message);
        if (result.ExitCode is ExitCodes.UsageOrConfiguration or ExitCodes.SuiteOrWorkspace)
            return new McpToolException(McpErrorCode.Validation, message);
        return new McpToolException(McpErrorCode.Internal, message);
    }

    private string? ToRelative(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return path;
        try
        {
            var full = Path.GetFullPath(path);
            var root = Path.GetFullPath(_workspace.RootPath);
            if (full.StartsWith(root, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                return Path.GetRelativePath(root, full).Replace('\\', '/');
        }
        catch
        {
            // fall through
        }

        return path;
    }

    private static int ClampOptional(int? value, int @default, int min, int max) =>
        Math.Clamp(value ?? @default, min, max);

    private static string? Truncate(string? value, int max)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= max)
            return value;
        return value[..Math.Max(0, max - 1)] + "…";
    }
}
