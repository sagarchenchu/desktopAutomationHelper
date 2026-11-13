using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Guarded AI semantic planner: retrieve → model proposal → validate/compile.
/// The model never calls the driver. Execution requires explicit authorization.
/// </summary>
public sealed class PlanningService
{
    private readonly IWorkspaceManager _workspace;
    private readonly PlanningOptions _options;
    private readonly AgentOptions _agentOptions;
    private readonly IPlanningModelClient _modelClient;
    private readonly PlanningRetrievalService _retrieval;
    private readonly PlanningProposalValidator _validator = new();
    private readonly PlanningProposalCompiler _compiler = new();
    private readonly PlanningArtifactWriter _artifacts;
    private readonly ObjectRepositoryReader _objectRepositoryReader;
    private readonly PlanObjectReferenceExpander _expander;
    private readonly ILogger<PlanningService> _logger;

    public PlanningService(
        IWorkspaceManager workspace,
        IPlanningModelClient modelClient,
        PlanningRetrievalService retrieval,
        PlanningArtifactWriter artifacts,
        ObjectRepositoryReader objectRepositoryReader,
        PlanObjectReferenceExpander expander,
        IOptions<AgentOptions> options,
        ILogger<PlanningService> logger)
    {
        _workspace = workspace;
        _modelClient = modelClient;
        _retrieval = retrieval;
        _artifacts = artifacts;
        _objectRepositoryReader = objectRepositoryReader;
        _expander = expander;
        _options = options.Value.Planning;
        _agentOptions = options.Value;
        _logger = logger;
    }

    public async Task<PlanningResult> ProposeAsync(PlanningRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        try
        {
            _workspace.EnsureInitialized();
            AgentOptionsValidator.Validate(_agentOptions, OptionsValidationScope.Workspace);

            var planningId = string.IsNullOrWhiteSpace(request.PlanningId)
                ? PlanningArtifactWriter.NewPlanningId()
                : PlanningArtifactWriter.SanitizePlanningId(request.PlanningId!);

            var goalErrors = PlanningPromptGuard.InspectGoal(request.Goal).ToList();
            var untrustedErrors = PlanningPromptGuard.InspectUntrusted(request.UntrustedContext).ToList();
            if (goalErrors.Count > 0 || untrustedErrors.Count > 0)
            {
                var denied = WriteDeniedEarly(
                    planningId,
                    request,
                    goalErrors.Concat(untrustedErrors).ToList(),
                    PlanningApprovalMode.Denied);
                return denied;
            }

            RetrievalManifest retrieval;
            try
            {
                retrieval = _retrieval.Retrieve(request);
            }
            catch (Exception ex) when (ex is InvalidOperationException or Knowledge.KnowledgePathException or WorkspaceException)
            {
                return WriteDeniedEarly(
                    planningId,
                    request,
                    [ex.Message],
                    PlanningApprovalMode.Denied);
            }

            var untrusted = TruncateUntrusted(request.UntrustedContext);
            var retrievalContextJson = BuildRetrievalContext(retrieval);

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(request.CancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(Math.Clamp(_options.RequestTimeoutSeconds, 1, 300)));

            PlanningModelResponse modelResponse;
            try
            {
                modelResponse = await _modelClient.ProposeAsync(
                    new PlanningModelRequest
                    {
                        CanonicalGoal = retrieval.CanonicalGoal,
                        SystemPolicy = PlanningPromptGuard.SystemPolicy,
                        RetrievalContextJson = retrievalContextJson,
                        UntrustedContext = untrusted,
                        MaxOutputBytes = _options.MaxOutputBytes,
                        MaxCompletionTokens = _options.MaxCompletionTokens,
                        CancellationToken = timeoutCts.Token
                    },
                    timeoutCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!request.CancellationToken.IsCancellationRequested)
            {
                return Fail(planningId, 7, "planning model request timed out or was cancelled by bound timeout.", null, null);
            }

            if (modelResponse.Truncated)
            {
                return PersistOutcome(
                    planningId,
                    request,
                    retrieval,
                    proposal: null,
                    candidate: null,
                    effectiveMode: PlanningApprovalMode.Denied,
                    validation: new PlanningValidationReport
                    {
                        Errors = ["model output truncated; rejected."],
                        EffectiveApprovalMode = PlanningRiskPolicy.ToWire(PlanningApprovalMode.Denied)
                    },
                    error: "model output truncated; rejected.",
                    exitCode: 5);
            }

            var parse = _validator.ParseAndValidateSchema(modelResponse.RawJson, _options.MaxOutputBytes);
            if (!parse.IsValid || parse.Proposal is null)
            {
                return PersistOutcome(
                    planningId,
                    request,
                    retrieval,
                    proposal: parse.Proposal,
                    candidate: null,
                    effectiveMode: PlanningApprovalMode.Denied,
                    validation: new PlanningValidationReport
                    {
                        SchemaValid = false,
                        Errors = parse.Errors,
                        EffectiveApprovalMode = PlanningRiskPolicy.ToWire(PlanningApprovalMode.Denied),
                        RetrievalManifestHash = PlanningArtifactWriter.HashObject(retrieval)
                    },
                    error: "proposal schema validation failed.",
                    exitCode: 5);
            }

            var proposal = parse.Proposal;
            // Align goal to canonical form (model may echo slightly different whitespace).
            proposal.Goal = retrieval.CanonicalGoal;

            OperationsCatalogDto catalog;
            if (!string.IsNullOrWhiteSpace(request.CatalogPath))
            {
                var catalogJson = File.ReadAllText(_workspace.ResolveSafePath(request.CatalogPath));
                catalog = JsonSerializer.Deserialize<OperationsCatalogDto>(catalogJson, PlanningJson.Compact)
                    ?? PlanningRetrievalService.MinimalSafeCatalog();
            }
            else
            {
                catalog = PlanningRetrievalService.MinimalSafeCatalog();
            }

            ObjectRepositorySnapshot? objectRepository = null;
            string? objectRepoRelative = null;
            var repoPath = string.IsNullOrWhiteSpace(request.ObjectRepositoryPath)
                ? "object-repository/repository.json"
                : request.ObjectRepositoryPath;
            var repoFull = _workspace.ResolveSafePath(repoPath);
            if (File.Exists(repoFull))
            {
                var repoResult = _objectRepositoryReader.Read(repoPath);
                if (!repoResult.IsValid || repoResult.Snapshot is null)
                    throw new InvalidOperationException("Object repository invalid: " + string.Join("; ", repoResult.Errors));
                objectRepository = repoResult.Snapshot;
                objectRepoRelative = repoPath.Replace('\\', '/');
            }

            var refErrors = _validator.ValidateApprovedRefs(proposal, retrieval, objectRepository, catalog);
            var riskMode = PlanningRiskPolicy.Decide(proposal, request.RequestedApprovalMode, out var riskReasons);
            var validation = new PlanningValidationReport
            {
                SchemaValid = true,
                RefsApproved = refErrors.Count == 0,
                Risk = proposal.Risk,
                ProposalHash = parse.ProposalHash,
                RetrievalManifestHash = PlanningArtifactWriter.HashObject(retrieval),
                Errors = [],
                Warnings = riskReasons
            };

            if (refErrors.Count > 0)
            {
                validation.Errors.AddRange(refErrors);
                validation.RiskPolicyPassed = false;
                validation.EffectiveApprovalMode = PlanningRiskPolicy.ToWire(PlanningApprovalMode.Denied);
                return PersistOutcome(
                    planningId,
                    request,
                    retrieval,
                    proposal,
                    candidate: null,
                    effectiveMode: PlanningApprovalMode.Denied,
                    validation,
                    error: "proposal references failed approval checks.",
                    exitCode: 5);
            }

            validation.RiskPolicyPassed = riskMode is not PlanningApprovalMode.Denied;
            validation.EffectiveApprovalMode = PlanningRiskPolicy.ToWire(riskMode);

            if (riskMode == PlanningApprovalMode.Denied)
            {
                validation.Errors.AddRange(riskReasons);
                return PersistOutcome(
                    planningId,
                    request,
                    retrieval,
                    proposal,
                    candidate: null,
                    effectiveMode: PlanningApprovalMode.Denied,
                    validation,
                    error: "proposal denied by risk policy.",
                    exitCode: 5);
            }

            // Compile only when not denied. approval-required and proposal-only still produce a candidate.
            var compile = _compiler.Compile(
                proposal,
                planningId,
                catalog,
                objectRepoRelative,
                objectRepository is null ? null : _expander,
                objectRepository);

            validation.PlanCompiled = compile.IsSuccess;
            validation.PlanValidatorPassed = compile.IsSuccess;
            validation.CatalogPreflightPassed = compile.IsSuccess;
            validation.CandidatePlanHash = compile.PlanHash;
            if (!compile.IsSuccess)
            {
                validation.Errors.AddRange(compile.Errors);
                return PersistOutcome(
                    planningId,
                    request,
                    retrieval,
                    proposal,
                    candidate: null,
                    effectiveMode: riskMode == PlanningApprovalMode.AutoRunKnownSafe
                        ? PlanningApprovalMode.ApprovalRequired
                        : riskMode,
                    validation,
                    error: "proposal failed Phase 2 plan validation/preflight.",
                    exitCode: 5);
            }

            // Auto-run cannot bypass validators — only execute when compile succeeded AND authorized.
            var executed = false;
            string? runId = null;
            if (request.ExecuteIfAuthorized
                && riskMode == PlanningApprovalMode.AutoRunKnownSafe
                && compile.Plan is not null)
            {
                // Execution is intentionally not performed inside PlanningService to keep the
                // model path free of driver calls. Callers (CLI) may run the candidate plan
                // through the Phase 2 runner after this result. Flag for CLI orchestration.
                executed = false;
                validation.Warnings.Add(
                    "auto-run-known-safe authorized; candidate plan ready for explicit Phase 2 run.");
            }
            else if (request.ExecuteIfAuthorized && riskMode != PlanningApprovalMode.AutoRunKnownSafe)
            {
                validation.Warnings.Add(
                    $"execute requested but effective mode is '{PlanningRiskPolicy.ToWire(riskMode)}'; not executing.");
            }

            var success = PersistOutcome(
                planningId,
                request,
                retrieval,
                proposal,
                compile.Plan,
                riskMode,
                validation,
                error: null,
                exitCode: 0);
            return new PlanningResult
            {
                Success = success.Success,
                ExitCode = success.ExitCode,
                PlanningId = planningId,
                ArtifactDirectory = success.ArtifactDirectory,
                EffectiveApprovalMode = riskMode,
                Proposal = proposal,
                CandidatePlan = compile.Plan,
                Retrieval = retrieval,
                Validation = validation,
                RequestSummary = success.RequestSummary,
                Executed = executed,
                RunId = runId
            };
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or AgentConfigurationException)
        {
            _logger.LogError(ex, "Planning failed");
            return new PlanningResult
            {
                Success = false,
                ExitCode = 5,
                Error = SecretRedactor.Redact(ex.Message),
                EffectiveApprovalMode = PlanningApprovalMode.Denied
            };
        }
    }

    private PlanningResult WriteDeniedEarly(
        string planningId,
        PlanningRequest request,
        List<string> errors,
        PlanningApprovalMode mode)
    {
        var retrieval = new RetrievalManifest
        {
            CanonicalGoal = PlanningPromptGuard.CanonicalizeGoal(request.Goal),
            GoalHash = PlanningArtifactWriter.HashObject(PlanningPromptGuard.CanonicalizeGoal(request.Goal)),
            Exclusions = ["early-deny"]
        };
        var validation = new PlanningValidationReport
        {
            SchemaValid = false,
            RefsApproved = false,
            RiskPolicyPassed = false,
            Errors = errors,
            EffectiveApprovalMode = PlanningRiskPolicy.ToWire(mode)
        };
        return PersistOutcome(
            planningId,
            request,
            retrieval,
            proposal: null,
            candidate: null,
            effectiveMode: mode,
            validation,
            error: string.Join("; ", errors),
            exitCode: 5);
    }

    private PlanningResult PersistOutcome(
        string planningId,
        PlanningRequest request,
        RetrievalManifest retrieval,
        PlanProposalDocument? proposal,
        Plans.PlanManifest? candidate,
        PlanningApprovalMode effectiveMode,
        PlanningValidationReport validation,
        string? error,
        int exitCode)
    {
        var summary = new PlanningRequestSummary
        {
            PlanningId = planningId,
            PlannerVersion = PlanningVersions.PlannerVersion,
            CanonicalGoal = retrieval.CanonicalGoal,
            GoalHash = retrieval.GoalHash,
            Provider = _modelClient.ProviderName,
            ModelId = _options.ModelId,
            RequestedApprovalMode = PlanningRiskPolicy.ToWire(request.RequestedApprovalMode),
            EffectiveApprovalMode = PlanningRiskPolicy.ToWire(effectiveMode),
            ExecuteIfAuthorized = request.ExecuteIfAuthorized,
            DryRun = request.DryRun,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            UntrustedContextPresent = !string.IsNullOrWhiteSpace(request.UntrustedContext),
            UntrustedContextChars = request.UntrustedContext?.Length ?? 0
        };

        validation.EffectiveApprovalMode = PlanningRiskPolicy.ToWire(effectiveMode);
        var dir = _artifacts.CreatePlanningDirectory(planningId);
        _artifacts.WriteAll(dir, summary, retrieval, proposal, validation, candidate);

        return new PlanningResult
        {
            Success = exitCode == 0 && candidate is not null,
            ExitCode = exitCode,
            PlanningId = planningId,
            ArtifactDirectory = dir,
            EffectiveApprovalMode = effectiveMode,
            Proposal = proposal,
            CandidatePlan = candidate,
            Retrieval = retrieval,
            Validation = validation,
            RequestSummary = summary,
            Error = error is null ? null : SecretRedactor.Redact(error)
        };
    }

    private static PlanningResult Fail(
        string planningId,
        int exitCode,
        string error,
        RetrievalManifest? retrieval,
        PlanProposalDocument? proposal) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            PlanningId = planningId,
            Error = SecretRedactor.Redact(error),
            EffectiveApprovalMode = PlanningApprovalMode.Denied,
            Retrieval = retrieval,
            Proposal = proposal
        };

    private string BuildRetrievalContext(RetrievalManifest retrieval)
    {
        var json = JsonSerializer.Serialize(retrieval, PlanningJson.Compact);
        var maxChars = Math.Max(512, _options.MaxPromptTokens * 4);
        if (json.Length > maxChars)
            json = json[..maxChars];
        return json;
    }

    private string? TruncateUntrusted(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        var max = Math.Clamp(_options.MaxUntrustedContextChars, 0, 100_000);
        var redacted = SecretRedactor.Redact(text);
        return redacted.Length <= max ? redacted : redacted[..max];
    }
}
