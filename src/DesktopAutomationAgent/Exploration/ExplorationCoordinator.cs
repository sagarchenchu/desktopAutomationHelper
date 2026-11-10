using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Exploration;

public interface IExplorationCoordinator
{
    Task<ExplorationResult> ExploreAsync(
        ExplorationRequest request,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// Policy-bounded observe/navigate explorer. No AI. Does not live in AgentCli,
/// DriverUiClient, or DeterministicPlanRunner.
/// </summary>
public sealed class ExplorationCoordinator : IExplorationCoordinator
{
    private readonly AgentOptions _options;
    private readonly ExplorationPolicyReader _policyReader;
    private readonly ISemanticObservationService _observation;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverUiClient _uiClient;
    private readonly StateFingerprinter _fingerprinter;
    private readonly CapabilityResolver _capabilityResolver;
    private readonly UiStabilizationService _stabilization;
    private readonly WindowTransitionDetector _windowDetector;
    private readonly ExplorationArtifactWriter _artifactWriter;
    private readonly ILogger<ExplorationCoordinator> _logger;

    public ExplorationCoordinator(
        IOptions<AgentOptions> options,
        ExplorationPolicyReader policyReader,
        ISemanticObservationService observation,
        IDriverConnectionResolver connectionResolver,
        IDriverUiClient uiClient,
        StateFingerprinter fingerprinter,
        CapabilityResolver capabilityResolver,
        UiStabilizationService stabilization,
        WindowTransitionDetector windowDetector,
        ExplorationArtifactWriter artifactWriter,
        ILogger<ExplorationCoordinator> logger)
    {
        _options = options.Value;
        _policyReader = policyReader;
        _observation = observation;
        _connectionResolver = connectionResolver;
        _uiClient = uiClient;
        _fingerprinter = fingerprinter;
        _capabilityResolver = capabilityResolver;
        _stabilization = stabilization;
        _windowDetector = windowDetector;
        _artifactWriter = artifactWriter;
        _logger = logger;
    }

    public async Task<ExplorationResult> ExploreAsync(
        ExplorationRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Workspace);
        }
        catch (AgentConfigurationException ex)
        {
            return Fail(ex.Message, ExitCodes.UsageOrConfiguration);
        }

        var policyResult = _policyReader.Read(request.PolicyPath);
        if (!policyResult.IsValid || policyResult.Policy is null)
        {
            return Fail(
                string.Join("; ", policyResult.Errors),
                ExitCodes.UsageOrConfiguration);
        }

        var policy = policyResult.Policy;
        var explorationId = DeterministicPlanRunner.GenerateRunId();
        var startedAt = DateTimeOffset.UtcNow;
        var deadline = startedAt.AddSeconds(policy.Budgets.TotalDurationSeconds);
        var observationOptions = (request.ObservationOptions ?? new ObservationOptions()) with
        {
            PersistArtifact = false,
            ActionableOnly = request.ObservationOptions?.ActionableOnly ?? false,
            IncludeFocusState = request.ObservationOptions?.IncludeFocusState ?? true,
            IncludeRuntimeId = request.ObservationOptions?.IncludeRuntimeId ?? true
        };

        string? runDirectory = null;
        try
        {
            runDirectory = _artifactWriter.CreateRunDirectory(explorationId);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or WorkspaceException)
        {
            return Fail(SecretRedactor.Redact(ex.Message), ExitCodes.SuiteOrWorkspace);
        }

        var recorder = new TransitionRecorder();
        var frontier = new FrontierQueue();
        var warnings = new List<string>(policyResult.Warnings);
        var deniedSamples = new List<string>();
        string? stopReason = null;
        var status = "completed";
        var exitCode = ExitCodes.Success;
        var actionsAttempted = 0;
        SemanticSnapshot? initialSnapshot = null;

        try
        {
            var initial = await _observation.ObserveAsync(observationOptions, cancellationToken).ConfigureAwait(false);
            if (!initial.Success || initial.Snapshot is null)
            {
                stopReason = initial.Error ?? "Initial observation failed.";
                status = "failed";
                exitCode = initial.ExitCode == 0 ? ExitCodes.ExecutionFailure : initial.ExitCode;
                return Finish();
            }

            initialSnapshot = initial.Snapshot;
            if (!WindowTransitionDetector.IsSnapshotInScope(initialSnapshot, policy))
            {
                stopReason = "Initial window/process/session is outside policy scope.";
                status = "failed";
                exitCode = ExitCodes.ExecutionFailure;
                return Finish();
            }

            if (!string.IsNullOrWhiteSpace(policy.Application.ProcessName)
                && !string.Equals(
                    initialSnapshot.Application.ProcessName,
                    policy.Application.ProcessName,
                    StringComparison.OrdinalIgnoreCase))
            {
                stopReason =
                    $"Observed process '{initialSnapshot.Application.ProcessName}' does not match policy processName '{policy.Application.ProcessName}'.";
                status = "failed";
                exitCode = ExitCodes.ExecutionFailure;
                return Finish();
            }

            var fingerprint = _fingerprinter.Fingerprint(initialSnapshot);
            var stateId = recorder.EnsureState(
                fingerprint,
                initialSnapshot.SnapshotId,
                initialSnapshot.Window?.Title,
                initialSnapshot.Window?.ClassName,
                initialSnapshot.Window?.IsModal,
                depth: 0,
                initialSnapshot.Nodes.Count,
                initialSnapshot.CapturedAtUtc);

            var candidates = _capabilityResolver.Resolve(initialSnapshot, fingerprint, policy, depth: 0);
            deniedSamples.AddRange(_capabilityResolver.DeniedActionSamples);

            if (request.DryRun || string.Equals(policy.Mode, "observe", StringComparison.OrdinalIgnoreCase))
            {
                stopReason = request.DryRun ? "dry-run" : "observe-mode";
                status = "completed";
                var summary = BuildSummary(
                    explorationId,
                    policy,
                    request.DryRun,
                    startedAt,
                    DateTimeOffset.UtcNow,
                    status,
                    ExitCodes.Success,
                    recorder,
                    actionsAttempted,
                    stopReason,
                    deniedSamples,
                    warnings,
                    runDirectory);

                _artifactWriter.WriteAll(
                    runDirectory!,
                    BuildManifest(explorationId, policy, policyResult, request, startedAt),
                    recorder.States,
                    recorder.Transitions,
                    summary);

                return new ExplorationResult
                {
                    Success = true,
                    ExitCode = ExitCodes.Success,
                    ExplorationId = explorationId,
                    ArtifactDirectory = runDirectory,
                    Summary = summary,
                    DryRunCandidates = candidates.ToList(),
                    InitialSnapshot = initialSnapshot
                };
            }

            frontier.EnqueueMany(candidates);
            _ = stateId;

            DriverConnection connection;
            try
            {
                connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (DriverConnectionException ex)
            {
                stopReason = ex.Message;
                status = "failed";
                exitCode = ExitCodes.DriverUnavailable;
                return Finish();
            }

            while (frontier.TryDequeue(out var action) && action is not null)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    stopReason = "cancelled";
                    status = "cancelled";
                    exitCode = ExitCodes.Cancelled;
                    break;
                }

                if (DateTimeOffset.UtcNow >= deadline)
                {
                    stopReason = "total-duration-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                if (actionsAttempted >= policy.Budgets.MaxActions)
                {
                    stopReason = "max-actions-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                if (recorder.States.Count > policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                if (action.Depth > policy.Budgets.MaxDepth)
                {
                    continue;
                }

                actionsAttempted++;
                var actionStarted = DateTimeOffset.UtcNow;
                var fromStateId = recorder.EnsureState(
                    action.StateFingerprint,
                    snapshotId: null,
                    windowTitle: null,
                    windowClassName: null,
                    isModal: null,
                    depth: action.Depth,
                    nodeCount: 0,
                    capturedAtUtc: actionStarted);

                WindowInventory beforeWindows;
                try
                {
                    beforeWindows = await _windowDetector.ListWindowsAsync(connection, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
                {
                    if (ex is OperationCanceledException && cancellationToken.IsCancellationRequested)
                    {
                        stopReason = "cancelled";
                        status = "cancelled";
                        exitCode = ExitCodes.Cancelled;
                        break;
                    }

                    recorder.Record(
                        fromStateId,
                        null,
                        action.StateFingerprint,
                        null,
                        action,
                        ExplorationOutcome.Failure,
                        SecretRedactor.Redact(ex.Message),
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                // Re-observe before acting so ephemeral refs are fresh for this snapshot.
                SemanticSnapshot beforeSnapshot;
                string beforeFingerprint;
                try
                {
                    var beforeObserve = await _observation.ObserveAsync(observationOptions, cancellationToken)
                        .ConfigureAwait(false);
                    if (!beforeObserve.Success || beforeObserve.Snapshot is null)
                    {
                        recorder.Record(
                            fromStateId,
                            null,
                            action.StateFingerprint,
                            null,
                            action,
                            ExplorationOutcome.Failure,
                            beforeObserve.Error ?? "Pre-action observation failed.",
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        continue;
                    }

                    beforeSnapshot = beforeObserve.Snapshot;
                    beforeFingerprint = _fingerprinter.Fingerprint(beforeSnapshot);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    stopReason = "cancelled";
                    status = "cancelled";
                    exitCode = ExitCodes.Cancelled;
                    break;
                }

                if (!string.Equals(beforeFingerprint, action.StateFingerprint, StringComparison.Ordinal))
                {
                    // State drifted; re-seed frontier from current state and skip stale action.
                    var reseeds = _capabilityResolver.Resolve(
                        beforeSnapshot,
                        beforeFingerprint,
                        policy,
                        action.Depth);
                    frontier.EnqueueMany(reseeds);
                    recorder.Record(
                        fromStateId,
                        recorder.EnsureState(
                            beforeFingerprint,
                            beforeSnapshot.SnapshotId,
                            beforeSnapshot.Window?.Title,
                            beforeSnapshot.Window?.ClassName,
                            beforeSnapshot.Window?.IsModal,
                            action.Depth,
                            beforeSnapshot.Nodes.Count,
                            beforeSnapshot.CapturedAtUtc),
                        action.StateFingerprint,
                        beforeFingerprint,
                        action,
                        ExplorationOutcome.StaleRef,
                        "State fingerprint changed before execution; action skipped.",
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                var refIndex = new EphemeralRefIndex(beforeSnapshot);
                if (!refIndex.TryResolve(action.SnapshotId, action.NodeRef, out var liveNode)
                    || liveNode is null)
                {
                    // Same fingerprint but refs rotated — resolve by element key.
                    liveNode = beforeSnapshot.Nodes.FirstOrDefault(n =>
                        string.Equals(
                            CapabilityResolver.BuildElementKey(n),
                            action.ElementKey,
                            StringComparison.Ordinal));
                    if (liveNode is null)
                    {
                        recorder.Record(
                            fromStateId,
                            null,
                            beforeFingerprint,
                            null,
                            action,
                            ExplorationOutcome.StaleRef,
                            $"Unable to resolve element '{action.ElementKey}' in current snapshot.",
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        continue;
                    }
                }

                PlanStep step;
                try
                {
                    step = BuildStep(action, liveNode, policy.Budgets.PerActionTimeoutSeconds);
                }
                catch (InvalidOperationException ex)
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.Failure,
                        ex.Message,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                try
                {
                    using var actionCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    actionCts.CancelAfter(TimeSpan.FromSeconds(policy.Budgets.PerActionTimeoutSeconds));
                    await _uiClient.ExecuteStepAsync(connection, step, actionCts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    stopReason = "cancelled";
                    status = "cancelled";
                    exitCode = ExitCodes.Cancelled;
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.Cancelled,
                        "Action cancelled.",
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    break;
                }
                catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.Failure,
                        SecretRedactor.Redact(ex.Message),
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                var stabilization = await _stabilization.WaitAsync(
                        observationOptions,
                        policy.Budgets,
                        snapshot => WindowTransitionDetector.IsSnapshotInScope(snapshot, policy),
                        cancellationToken)
                    .ConfigureAwait(false);

                if (stabilization.Detail == "Scope violation detected during stabilization."
                    && stabilization.Snapshot is not null)
                {
                    var scopeStateId = recorder.EnsureState(
                        stabilization.LastFingerprint ?? _fingerprinter.Fingerprint(stabilization.Snapshot),
                        stabilization.Snapshot.SnapshotId,
                        stabilization.Snapshot.Window?.Title,
                        stabilization.Snapshot.Window?.ClassName,
                        stabilization.Snapshot.Window?.IsModal,
                        action.Depth + 1,
                        stabilization.Snapshot.Nodes.Count,
                        stabilization.Snapshot.CapturedAtUtc);
                    recorder.Record(
                        fromStateId,
                        scopeStateId,
                        beforeFingerprint,
                        stabilization.LastFingerprint,
                        action,
                        ExplorationOutcome.ScopeViolation,
                        stabilization.Detail,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    stopReason = "scope-violation";
                    status = "failed";
                    exitCode = ExitCodes.ExecutionFailure;
                    break;
                }

                if (stabilization.TimedOut)
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        stabilization.LastFingerprint,
                        action,
                        ExplorationOutcome.StabilizationTimeout,
                        stabilization.Detail,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                if (!stabilization.Stabilized || stabilization.Snapshot is null)
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        stabilization.LastFingerprint,
                        action,
                        ExplorationOutcome.Failure,
                        stabilization.Detail ?? "Stabilization failed.",
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                var afterSnapshot = stabilization.Snapshot;
                var afterFingerprint = stabilization.LastFingerprint
                    ?? _fingerprinter.Fingerprint(afterSnapshot);

                WindowInventory afterWindows;
                try
                {
                    afterWindows = await _windowDetector.ListWindowsAsync(connection, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        afterFingerprint,
                        action,
                        ExplorationOutcome.Failure,
                        SecretRedactor.Redact(ex.Message),
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                var transition = _windowDetector.Compare(beforeWindows, afterWindows, afterSnapshot, policy);
                if (transition.ScopeViolation)
                {
                    var scopeState = recorder.EnsureState(
                        afterFingerprint,
                        afterSnapshot.SnapshotId,
                        afterSnapshot.Window?.Title,
                        afterSnapshot.Window?.ClassName,
                        afterSnapshot.Window?.IsModal,
                        action.Depth + 1,
                        afterSnapshot.Nodes.Count,
                        afterSnapshot.CapturedAtUtc);
                    recorder.Record(
                        fromStateId,
                        scopeState,
                        beforeFingerprint,
                        afterFingerprint,
                        action,
                        ExplorationOutcome.ScopeViolation,
                        transition.Detail,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    stopReason = "scope-violation";
                    status = "failed";
                    exitCode = ExitCodes.ExecutionFailure;
                    break;
                }

                var outcome = ExplorationOutcome.Success;
                if (string.Equals(beforeFingerprint, afterFingerprint, StringComparison.Ordinal))
                    outcome = ExplorationOutcome.NoOp;
                else if (transition.ModalDialog)
                    outcome = ExplorationOutcome.ModalDialog;
                else if (transition.NewWindow)
                    outcome = ExplorationOutcome.NewWindow;
                else if (transition.ClosedWindow)
                    outcome = ExplorationOutcome.ClosedWindow;

                var toDepth = Math.Min(action.Depth + 1, policy.Budgets.MaxDepth);
                var toStateId = recorder.EnsureState(
                    afterFingerprint,
                    afterSnapshot.SnapshotId,
                    afterSnapshot.Window?.Title,
                    afterSnapshot.Window?.ClassName,
                    afterSnapshot.Window?.IsModal,
                    toDepth,
                    afterSnapshot.Nodes.Count,
                    afterSnapshot.CapturedAtUtc);

                recorder.Record(
                    fromStateId,
                    toStateId,
                    beforeFingerprint,
                    afterFingerprint,
                    action,
                    outcome,
                    transition.Detail,
                    actionStarted,
                    DateTimeOffset.UtcNow);

                if (recorder.States.Count >= policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                if (toDepth <= policy.Budgets.MaxDepth
                    && !string.Equals(beforeFingerprint, afterFingerprint, StringComparison.Ordinal))
                {
                    var next = _capabilityResolver.Resolve(afterSnapshot, afterFingerprint, policy, toDepth);
                    deniedSamples.AddRange(_capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
                    frontier.EnqueueMany(next);
                }
            }

            stopReason ??= frontier.Count == 0 ? "frontier-exhausted" : stopReason;
            if (status == "completed" && stopReason is null)
                stopReason = "frontier-exhausted";

            return Finish();
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            stopReason = "cancelled";
            status = "cancelled";
            exitCode = ExitCodes.Cancelled;
            return Finish();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Exploration {ExplorationId} failed unexpectedly.", explorationId);
            stopReason = SecretRedactor.Redact(ex.Message);
            status = "failed";
            exitCode = ExitCodes.ExecutionFailure;
            return Finish();
        }

        ExplorationResult Finish()
        {
            var completedAt = DateTimeOffset.UtcNow;
            var summary = BuildSummary(
                explorationId,
                policy,
                request.DryRun,
                startedAt,
                completedAt,
                status,
                exitCode,
                recorder,
                actionsAttempted,
                stopReason,
                deniedSamples,
                warnings,
                runDirectory);

            try
            {
                _artifactWriter.WriteAll(
                    runDirectory!,
                    BuildManifest(explorationId, policy, policyResult, request, startedAt),
                    recorder.States,
                    recorder.Transitions,
                    summary);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                warnings.Add($"Artifact write failed: {SecretRedactor.Redact(ex.Message)}");
                summary.Warnings = warnings;
            }

            return new ExplorationResult
            {
                Success = exitCode is ExitCodes.Success or ExitCodes.Cancelled,
                ExitCode = exitCode,
                Error = exitCode == ExitCodes.Success ? null : stopReason,
                ExplorationId = explorationId,
                ArtifactDirectory = runDirectory,
                Summary = summary,
                InitialSnapshot = initialSnapshot
            };
        }
    }

    private static PlanStep BuildStep(
        ExplorationActionCandidate action,
        SemanticNode node,
        int timeoutSeconds)
    {
        if (string.Equals(action.DriverOperation, "doubleclick", StringComparison.OrdinalIgnoreCase)
            || string.Equals(action.DriverOperation, "doubleclickuia", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Double-click is unsupported in Phase 4B.");
        }

        var locator = new ObjectLocator
        {
            AutomationId = string.IsNullOrWhiteSpace(node.AutomationId) ? null : node.AutomationId,
            Name = string.IsNullOrWhiteSpace(node.Name) ? null : node.Name,
            ControlType = string.IsNullOrWhiteSpace(node.ControlType) ? null : node.ControlType,
            ClassName = string.IsNullOrWhiteSpace(node.ClassName) ? null : node.ClassName
        };

        if (locator.AutomationId is null && locator.Name is null && locator.ClassName is null)
            throw new InvalidOperationException($"Element '{action.ElementKey}' has no stable locator fields.");

        var args = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
        {
            ["locator"] = ObjectLocatorSerializer.ToJsonElement(locator),
            ["timeoutMs"] = JsonSerializer.SerializeToElement(Math.Clamp(timeoutSeconds * 1000, 500, 120_000))
        };

        if (action.DriverArguments is not null)
        {
            foreach (var pair in action.DriverArguments)
                args[pair.Key] = JsonSerializer.SerializeToElement(pair.Value);
        }

        return new PlanStep
        {
            Id = $"explore-{action.Kind}-{node.Ordinal}",
            Operation = action.DriverOperation,
            Arguments = args
        };
    }

    private static ExplorationManifest BuildManifest(
        string explorationId,
        ExplorationPolicyDocument policy,
        ExplorationPolicyValidationResult policyResult,
        ExplorationRequest request,
        DateTimeOffset startedAt) =>
        new()
        {
            ExplorationId = explorationId,
            PolicyId = policy.PolicyId,
            Mode = policy.Mode,
            DryRun = request.DryRun,
            StartedAtUtc = startedAt,
            PolicyPath = policyResult.FilePath,
            PolicySha256 = policyResult.Sha256
        };

    private static ExplorationSummary BuildSummary(
        string explorationId,
        ExplorationPolicyDocument policy,
        bool dryRun,
        DateTimeOffset startedAt,
        DateTimeOffset completedAt,
        string status,
        int exitCode,
        TransitionRecorder recorder,
        int actionsAttempted,
        string? stopReason,
        List<string> deniedSamples,
        List<string> warnings,
        string? runDirectory) =>
        new()
        {
            ExplorationId = explorationId,
            Status = status,
            ExitCode = exitCode,
            Mode = policy.Mode,
            DryRun = dryRun,
            StartedAtUtc = startedAt,
            CompletedAtUtc = completedAt,
            StatesVisited = recorder.States.Count,
            ActionsAttempted = actionsAttempted,
            TransitionsRecorded = recorder.Transitions.Count,
            MaxDepthReached = recorder.MaxDepthReached,
            StopReason = stopReason,
            DeniedActionSamples = deniedSamples.Distinct(StringComparer.Ordinal).Take(32).ToList(),
            Warnings = warnings,
            ArtifactDirectory = runDirectory
        };

    private static ExplorationResult Fail(string message, int exitCode) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = SecretRedactor.Redact(message)
        };
}
