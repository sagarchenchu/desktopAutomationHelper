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
    private readonly IProcessIdentityProbe _identityProbe;
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
        ILogger<ExplorationCoordinator> logger,
        IProcessIdentityProbe? identityProbe = null)
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
        _identityProbe = identityProbe ?? new SystemProcessIdentityProbe();
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

        // Crawl-scoped identity: pin root PID/session for this ExploreAsync only.
        var identityValidator = new ProcessIdentityValidator(_identityProbe);

        var recorder = new TransitionRecorder();
        var frontier = new FrontierQueue();
        var warnings = new List<string>(policyResult.Warnings);
        var deniedSamples = new List<string>();
        string? stopReason = null;
        var status = "completed";
        var exitCode = ExitCodes.Success;
        var actionsAttempted = 0;
        SemanticSnapshot? initialSnapshot = null;

        using var deadlineCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        LinkDeadline(deadlineCts, deadline);

        try
        {
            var initial = await _observation.ObserveAsync(observationOptions, deadlineCts.Token).ConfigureAwait(false);
            if (!initial.Success || initial.Snapshot is null)
            {
                stopReason = initial.Error ?? "Initial observation failed.";
                status = "failed";
                exitCode = initial.ExitCode == 0 ? ExitCodes.ExecutionFailure : initial.ExitCode;
                return Finish();
            }

            initialSnapshot = initial.Snapshot;
            if (!TryValidateScope(initialSnapshot, policy, identityValidator, out var initialScopeError))
            {
                stopReason = initialScopeError;
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

            if (policy.Budgets.MaxDepth > 0)
                frontier.EnqueueMany(candidates);
            _ = stateId;

            DriverConnection connection;
            try
            {
                connection = await _connectionResolver.ResolveAsync(deadlineCts.Token).ConfigureAwait(false);
            }
            catch (DriverConnectionException ex)
            {
                stopReason = ex.Message;
                status = "failed";
                exitCode = ExitCodes.DriverUnavailable;
                return Finish();
            }

            var windowStack = new ExplorationWindowStack();

            while (true)
            {
                if (!frontier.TryDequeue(out var action) || action is null)
                {
                    if (policy.AllowCloseChildWindowsForBacktrack
                        && windowStack.Count > 0
                        && recorder.States.Count < policy.Budgets.MaxStates
                        && actionsAttempted < policy.Budgets.MaxActions
                        && DateTimeOffset.UtcNow < deadline
                        && !cancellationToken.IsCancellationRequested
                        && !deadlineCts.IsCancellationRequested)
                    {
                        var backtrack = await TryBacktrackAsync(
                                connection,
                                windowStack,
                                policy,
                                observationOptions,
                                frontier,
                                recorder,
                                deniedSamples,
                                identityValidator,
                                deadline,
                                actionsAttempted,
                                deadlineCts.Token)
                            .ConfigureAwait(false);
                        actionsAttempted = backtrack.ActionsAttempted;
                        if (backtrack.Success)
                            continue;
                    }

                    // Prefer budget/cancel over mislabeling an empty frontier.
                    if (cancellationToken.IsCancellationRequested)
                    {
                        stopReason = "cancelled";
                        status = "cancelled";
                        exitCode = ExitCodes.Cancelled;
                    }
                    else if (deadlineCts.IsCancellationRequested || DateTimeOffset.UtcNow >= deadline)
                    {
                        stopReason = "total-duration-budget";
                        status = "budget-exhausted";
                        exitCode = ExitCodes.Success;
                    }
                    else if (actionsAttempted >= policy.Budgets.MaxActions)
                    {
                        stopReason = "max-actions-budget";
                        status = "budget-exhausted";
                        exitCode = ExitCodes.Success;
                    }
                    else if (recorder.States.Count >= policy.Budgets.MaxStates)
                    {
                        stopReason = "max-states-budget";
                        status = "budget-exhausted";
                        exitCode = ExitCodes.Success;
                    }
                    else
                    {
                        stopReason = "frontier-exhausted";
                        status = "completed";
                        exitCode = ExitCodes.Success;
                    }

                    break;
                }

                if (cancellationToken.IsCancellationRequested || deadlineCts.IsCancellationRequested)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        stopReason = "cancelled";
                        status = "cancelled";
                        exitCode = ExitCodes.Cancelled;
                    }
                    else
                    {
                        stopReason = "total-duration-budget";
                        status = "budget-exhausted";
                        exitCode = ExitCodes.Success;
                    }

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

                if (recorder.States.Count >= policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                // Depth hard stop: never execute at or beyond maxDepth (target would exceed).
                if (action.Depth >= policy.Budgets.MaxDepth)
                    continue;

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

                if (recorder.States.Count > policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                WindowInventory beforeWindows;
                try
                {
                    beforeWindows = await _windowDetector.ListWindowsAsync(connection, deadlineCts.Token)
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                    break;
                }
                catch (UiExecutionException ex)
                {
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
                    var beforeObserve = await _observation.ObserveAsync(observationOptions, deadlineCts.Token)
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
                catch (OperationCanceledException)
                {
                    (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                    break;
                }

                // Blocker: scope after every observation (including pre-action) before resolve/execute.
                if (!TryValidateScope(beforeSnapshot, policy, identityValidator, out var beforeScopeError))
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.ScopeViolation,
                        beforeScopeError,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    stopReason = "scope-violation";
                    status = "failed";
                    exitCode = ExitCodes.ExecutionFailure;
                    break;
                }

                // Window-aware frontier: switch/revalidate the candidate's target HWND+PID
                // before execution so sibling-window queues cannot act on the wrong surface.
                if (action.TargetHwnd is long targetHwnd && targetHwnd > 0
                    && action.TargetProcessId is int targetPid && targetPid > 0
                    && beforeSnapshot.Window?.Hwnd != targetHwnd)
                {
                    if (actionsAttempted >= policy.Budgets.MaxActions
                        || DateTimeOffset.UtcNow >= deadline
                        || deadlineCts.IsCancellationRequested)
                    {
                        (stopReason, status, exitCode) = BudgetOrCancel(
                            cancellationToken,
                            actionsAttempted,
                            policy.Budgets.MaxActions,
                            deadline);
                        break;
                    }

                    var retarget = await TrySwitchToHwndAsync(
                            connection,
                            targetHwnd,
                            targetPid,
                            identityValidator,
                            policy.Budgets.PerActionTimeoutSeconds,
                            deadlineCts.Token)
                        .ConfigureAwait(false);
                    actionsAttempted++;
                    if (!retarget.Success)
                    {
                        recorder.Record(
                            fromStateId,
                            null,
                            beforeFingerprint,
                            null,
                            action,
                            ExplorationOutcome.StaleRef,
                            retarget.Detail ?? "Target HWND switch failed before execution.",
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        continue;
                    }

                    try
                    {
                        var retargetObserve = await _observation
                            .ObserveAsync(observationOptions, deadlineCts.Token)
                            .ConfigureAwait(false);
                        if (!retargetObserve.Success || retargetObserve.Snapshot is null)
                        {
                            recorder.Record(
                                fromStateId,
                                null,
                                beforeFingerprint,
                                null,
                                action,
                                ExplorationOutcome.Failure,
                                retargetObserve.Error ?? "Post-retarget observation failed.",
                                actionStarted,
                                DateTimeOffset.UtcNow);
                            continue;
                        }

                        beforeSnapshot = retargetObserve.Snapshot;
                        beforeFingerprint = _fingerprinter.Fingerprint(beforeSnapshot);
                    }
                    catch (OperationCanceledException)
                    {
                        (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                        break;
                    }

                    if (!TryValidateScope(
                            beforeSnapshot,
                            policy,
                            identityValidator,
                            out var retargetScopeError))
                    {
                        recorder.Record(
                            fromStateId,
                            null,
                            beforeFingerprint,
                            null,
                            action,
                            ExplorationOutcome.ScopeViolation,
                            retargetScopeError,
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        stopReason = "scope-violation";
                        status = "failed";
                        exitCode = ExitCodes.ExecutionFailure;
                        break;
                    }

                    var livePid = beforeSnapshot.Window?.ProcessId
                                  ?? beforeSnapshot.Application.ProcessId;
                    if (beforeSnapshot.Window?.Hwnd != targetHwnd || livePid != targetPid)
                    {
                        recorder.Record(
                            fromStateId,
                            null,
                            beforeFingerprint,
                            null,
                            action,
                            ExplorationOutcome.StaleRef,
                            $"Post-retarget HWND/PID 0x{beforeSnapshot.Window?.Hwnd:X}/{livePid} " +
                            $"does not match target 0x{targetHwnd:X}/{targetPid}.",
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        continue;
                    }
                }

                if (!string.Equals(beforeFingerprint, action.StateFingerprint, StringComparison.Ordinal))
                {
                    // State drifted; re-seed frontier from current in-scope state and skip stale action.
                    if (action.Depth < policy.Budgets.MaxDepth
                        && recorder.States.Count < policy.Budgets.MaxStates)
                    {
                        var reseeds = _capabilityResolver.Resolve(
                            beforeSnapshot,
                            beforeFingerprint,
                            policy,
                            action.Depth);
                        deniedSamples.AddRange(
                            _capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
                        if (action.Depth < policy.Budgets.MaxDepth)
                            frontier.EnqueueMany(reseeds);
                    }

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

                if (!TryResolveLiveNode(beforeSnapshot, action, out var liveNode, out var resolveDetail))
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.StaleRef,
                        resolveDetail,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                if (!_capabilityResolver.IsStillAuthorized(liveNode!, action, policy, out var reauthDeny))
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.StaleRef,
                        $"Live node re-authorization failed: {reauthDeny ?? "deny-by-default"}.",
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    continue;
                }

                PlanStep step;
                try
                {
                    step = BuildStep(action, liveNode!, policy.Budgets.PerActionTimeoutSeconds);
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
                    using var actionCts = CancellationTokenSource.CreateLinkedTokenSource(deadlineCts.Token);
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
                catch (OperationCanceledException) when (deadlineCts.IsCancellationRequested
                                                         && !cancellationToken.IsCancellationRequested)
                {
                    stopReason = "total-duration-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        ExplorationOutcome.BudgetExhausted,
                        "Total duration budget reached during action.",
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

                StabilizationResult stabilization;
                try
                {
                    stabilization = await _stabilization.WaitAsync(
                            observationOptions,
                            policy.Budgets,
                            snapshot =>
                            {
                                if (!WindowTransitionDetector.IsSnapshotInScope(snapshot, policy))
                                    return false;
                                var identity = identityValidator.ValidateSnapshot(snapshot, policy);
                                return identity.IsValid;
                            },
                            deadlineCts.Token)
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        null,
                        action,
                        cancellationToken.IsCancellationRequested
                            ? ExplorationOutcome.Cancelled
                            : ExplorationOutcome.BudgetExhausted,
                        stopReason,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    break;
                }

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

                // Scope after post-action / stabilization observation.
                if (!TryValidateScope(afterSnapshot, policy, identityValidator, out var afterScopeError))
                {
                    recorder.Record(
                        fromStateId,
                        null,
                        beforeFingerprint,
                        afterFingerprint,
                        action,
                        ExplorationOutcome.ScopeViolation,
                        afterScopeError,
                        actionStarted,
                        DateTimeOffset.UtcNow);
                    stopReason = "scope-violation";
                    status = "failed";
                    exitCode = ExitCodes.ExecutionFailure;
                    break;
                }

                WindowInventory afterWindows;
                try
                {
                    afterWindows = await _windowDetector.ListWindowsAsync(connection, deadlineCts.Token)
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                    break;
                }
                catch (UiExecutionException ex)
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

                var transition = _windowDetector.Compare(
                    beforeWindows,
                    afterWindows,
                    afterSnapshot,
                    policy,
                    identityValidator);
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

                // NewWindow / ModalDialog are independent of fingerprint NoOp — a child can
                // appear while the parent fingerprint is unchanged. Never assume auto-active.
                var outcome = ExplorationOutcome.Success;
                if (transition.ModalDialog || afterSnapshot.Window?.IsModal == true)
                    outcome = ExplorationOutcome.ModalDialog;
                else if (transition.NewWindow)
                    outcome = ExplorationOutcome.NewWindow;
                else if (transition.ClosedWindow)
                    outcome = ExplorationOutcome.ClosedWindow;
                else if (string.Equals(beforeFingerprint, afterFingerprint, StringComparison.Ordinal))
                    outcome = ExplorationOutcome.NoOp;

                var nextDepth = action.Depth + 1;
                var boundChildSurfaces = new List<(SemanticSnapshot Snapshot, string Fingerprint)>();
                var didSwitchToChild = false;
                string? newWindowBindFailure = null;

                if (outcome is ExplorationOutcome.ModalDialog or ExplorationOutcome.NewWindow)
                {
                    foreach (var added in transition.AddedWindows)
                    {
                        if (added.Hwnd is not long childHwnd || childHwnd <= 0
                            || added.ProcessId is not int childPid || childPid <= 0)
                        {
                            newWindowBindFailure =
                                $"New window '{added.Title}' is missing hwnd/processId; fail-closed.";
                            break;
                        }

                        SemanticSnapshot? childSnapshot = null;
                        string? bindDetail = null;
                        const int maxBindAttempts = 3;
                        for (var attempt = 1; attempt <= maxBindAttempts; attempt++)
                        {
                            if (actionsAttempted >= policy.Budgets.MaxActions
                                || DateTimeOffset.UtcNow >= deadline
                                || deadlineCts.IsCancellationRequested)
                            {
                                // Prefer budget/cancel over new-window-bind-failed.
                                (stopReason, status, exitCode) = BudgetOrCancel(
                                    cancellationToken,
                                    actionsAttempted,
                                    policy.Budgets.MaxActions,
                                    deadline);
                                bindDetail = stopReason;
                                break;
                            }

                            var switched = await TrySwitchToHwndAsync(
                                    connection,
                                    childHwnd,
                                    childPid,
                                    identityValidator,
                                    policy.Budgets.PerActionTimeoutSeconds,
                                    deadlineCts.Token)
                                .ConfigureAwait(false);
                            actionsAttempted++;
                            if (!switched.Success)
                            {
                                bindDetail = switched.Detail ?? "switchwindow failed.";
                                continue;
                            }

                            try
                            {
                                var childObserve = await _observation
                                    .ObserveAsync(observationOptions, deadlineCts.Token)
                                    .ConfigureAwait(false);
                                if (!childObserve.Success || childObserve.Snapshot is null)
                                {
                                    bindDetail = childObserve.Error ?? "Post-switch observation failed.";
                                    continue;
                                }

                                childSnapshot = childObserve.Snapshot;
                            }
                            catch (OperationCanceledException)
                            {
                                (stopReason, status, exitCode) = CancelOrDeadline(cancellationToken);
                                break;
                            }

                            if (!TryValidateScope(
                                    childSnapshot,
                                    policy,
                                    identityValidator,
                                    out var childScopeError))
                            {
                                bindDetail = childScopeError;
                                childSnapshot = null;
                                continue;
                            }

                            // Exact-HWND verification — only then mark switch successful.
                            if (childSnapshot.Window?.Hwnd != childHwnd)
                            {
                                bindDetail =
                                    $"Post-switch snapshot HWND 0x{childSnapshot.Window?.Hwnd:X} " +
                                    $"does not equal requested 0x{childHwnd:X}.";
                                childSnapshot = null;
                                continue;
                            }

                            bindDetail = null;
                            break;
                        }

                        if (stopReason is not null)
                            break;

                        if (childSnapshot is null || childSnapshot.Window?.Hwnd != childHwnd)
                        {
                            newWindowBindFailure =
                                $"Failed to bind new window HWND 0x{childHwnd:X}: {bindDetail ?? "unknown"}.";
                            break;
                        }

                        didSwitchToChild = true;
                        var childFp = _fingerprinter.Fingerprint(childSnapshot);
                        boundChildSurfaces.Add((childSnapshot, childFp));

                        // Keep the latest bound surface as the recorded after-state.
                        afterSnapshot = childSnapshot;
                        afterFingerprint = childFp;

                        if (childSnapshot.Window?.IsModal == true)
                            outcome = ExplorationOutcome.ModalDialog;

                        if (policy.AllowCloseChildWindowsForBacktrack)
                        {
                            windowStack.PushChild(new ExplorationWindowFrame
                            {
                                ChildHwnd = childHwnd,
                                ChildProcessId = childPid,
                                ChildTitle = added.Title ?? childSnapshot.Window?.Title,
                                ParentHwnd = beforeSnapshot.Window?.Hwnd,
                                ParentProcessId = beforeSnapshot.Application.ProcessId,
                                ParentTitle = beforeSnapshot.Window?.Title,
                                IsModal = outcome == ExplorationOutcome.ModalDialog
                                          || childSnapshot.Window?.IsModal == true,
                                Depth = nextDepth
                            });
                        }
                    }

                    if (stopReason is not null)
                        break;

                    if (newWindowBindFailure is not null)
                    {
                        // If binding stopped because the budget/cancel already applied,
                        // BudgetOrCancel set stopReason above — never rewrite as bind-failed.
                        recorder.Record(
                            fromStateId,
                            null,
                            beforeFingerprint,
                            afterFingerprint,
                            action,
                            ExplorationOutcome.ScopeViolation,
                            newWindowBindFailure,
                            actionStarted,
                            DateTimeOffset.UtcNow);
                        stopReason = "new-window-bind-failed";
                        status = "failed";
                        exitCode = ExitCodes.ExecutionFailure;
                        break;
                    }
                }

                var toStateId = recorder.EnsureState(
                    afterFingerprint,
                    afterSnapshot.SnapshotId,
                    afterSnapshot.Window?.Title,
                    afterSnapshot.Window?.ClassName,
                    afterSnapshot.Window?.IsModal,
                    Math.Min(nextDepth, policy.Budgets.MaxDepth),
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

                if (recorder.States.Count > policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                    break;
                }

                // Enqueue every successfully bound added window separately (not only the last).
                if (nextDepth < policy.Budgets.MaxDepth
                    && recorder.States.Count < policy.Budgets.MaxStates)
                {
                    if (boundChildSurfaces.Count > 0)
                    {
                        foreach (var bound in boundChildSurfaces)
                        {
                            if (recorder.States.Count >= policy.Budgets.MaxStates)
                                break;

                            var next = _capabilityResolver.Resolve(
                                bound.Snapshot,
                                bound.Fingerprint,
                                policy,
                                nextDepth);
                            deniedSamples.AddRange(
                                _capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
                            frontier.EnqueueMany(next);
                        }
                    }
                    else if (didSwitchToChild
                             || outcome is ExplorationOutcome.ModalDialog or ExplorationOutcome.NewWindow
                             || !string.Equals(beforeFingerprint, afterFingerprint, StringComparison.Ordinal))
                    {
                        var next = _capabilityResolver.Resolve(
                            afterSnapshot,
                            afterFingerprint,
                            policy,
                            nextDepth);
                        deniedSamples.AddRange(
                            _capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
                        frontier.EnqueueMany(next);
                    }
                }
            }

            // Intentionally no post-budget mutating cleanup: when the loop stops for budget/
            // cancel/scope, do not close remaining stacked windows.

            // If the main loop broke for an action-path reason, stopReason is already set.
            // Only default to frontier-exhausted when nothing else explains the stop.
            if (stopReason is null)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    stopReason = "cancelled";
                    status = "cancelled";
                    exitCode = ExitCodes.Cancelled;
                }
                else if (deadlineCts.IsCancellationRequested || DateTimeOffset.UtcNow >= deadline)
                {
                    stopReason = "total-duration-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                }
                else if (actionsAttempted >= policy.Budgets.MaxActions)
                {
                    stopReason = "max-actions-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                }
                else if (recorder.States.Count >= policy.Budgets.MaxStates)
                {
                    stopReason = "max-states-budget";
                    status = "budget-exhausted";
                    exitCode = ExitCodes.Success;
                }
                else if (frontier.Count == 0)
                {
                    stopReason = "frontier-exhausted";
                }
            }

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
        catch (OperationCanceledException) when (deadlineCts.IsCancellationRequested)
        {
            stopReason = "total-duration-budget";
            status = "budget-exhausted";
            exitCode = ExitCodes.Success;
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

    /// <summary>
    /// Close the top stacked child by exact HWND + verified PID using HWND-targeted
    /// WindowPattern-only closewindow (never ActiveWindow / first top-level / Close button).
    /// Verifies post-switch snapshot HWND. Pop only after inventory confirms the child is
    /// gone. Counts every mutation — including failed/timed-out closes — against maxActions
    /// with per-action timeout. Preserves the frame's original depth.
    /// </summary>
    private async Task<(bool Success, int ActionsAttempted)> TryBacktrackAsync(
        DriverConnection connection,
        ExplorationWindowStack windowStack,
        ExplorationPolicyDocument policy,
        ObservationOptions observationOptions,
        FrontierQueue frontier,
        TransitionRecorder recorder,
        List<string> deniedSamples,
        ProcessIdentityValidator identityValidator,
        DateTimeOffset deadline,
        int actionsAttempted,
        CancellationToken cancellationToken)
    {
        if (!windowStack.TryPeek(out var frame) || frame is null)
            return (false, actionsAttempted);

        if (frame.ChildHwnd <= 0 || frame.ChildProcessId <= 0)
            return (false, actionsAttempted);

        if (actionsAttempted >= policy.Budgets.MaxActions
            || DateTimeOffset.UtcNow >= deadline
            || cancellationToken.IsCancellationRequested)
        {
            return (false, actionsAttempted);
        }

        var startedAt = DateTimeOffset.UtcNow;
        var preservedDepth = frame.Depth;
        var backtrackAction = new ExplorationActionCandidate
        {
            ActionKey = $"backtrack:close:0x{frame.ChildHwnd:X}",
            Kind = ExplorationActionKind.Invoke,
            Category = "backtrack",
            DriverOperation = "closewindow",
            ElementKey = $"Window|hwnd:0x{frame.ChildHwnd:X}|pid:{frame.ChildProcessId}",
            Name = frame.ChildTitle,
            ControlType = "Window",
            Depth = preservedDepth,
            StateFingerprint = string.Empty
        };

        string fromFingerprint;
        string fromStateId;
        try
        {
            var beforeObserve = await _observation
                .ObserveAsync(observationOptions, cancellationToken)
                .ConfigureAwait(false);
            if (beforeObserve.Success && beforeObserve.Snapshot is not null)
            {
                fromFingerprint = _fingerprinter.Fingerprint(beforeObserve.Snapshot);
                fromStateId = recorder.EnsureState(
                    fromFingerprint,
                    beforeObserve.Snapshot.SnapshotId,
                    beforeObserve.Snapshot.Window?.Title,
                    beforeObserve.Snapshot.Window?.ClassName,
                    beforeObserve.Snapshot.Window?.IsModal,
                    preservedDepth,
                    beforeObserve.Snapshot.Nodes.Count,
                    beforeObserve.Snapshot.CapturedAtUtc);
            }
            else
            {
                fromFingerprint = $"backtrack-pending:0x{frame.ChildHwnd:X}";
                fromStateId = recorder.EnsureState(
                    fromFingerprint,
                    null,
                    frame.ChildTitle,
                    null,
                    frame.IsModal,
                    preservedDepth,
                    0,
                    startedAt);
            }
        }
        catch (OperationCanceledException)
        {
            return (false, actionsAttempted);
        }

        // 1) Switch to child by exact HWND.
        var switchChild = await TrySwitchToHwndAsync(
                connection,
                frame.ChildHwnd,
                frame.ChildProcessId,
                identityValidator,
                policy.Budgets.PerActionTimeoutSeconds,
                cancellationToken)
            .ConfigureAwait(false);
        actionsAttempted++;
        if (!switchChild.Success)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.Failure,
                switchChild.Detail ?? "Backtrack switch to child HWND failed.",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        if (actionsAttempted >= policy.Budgets.MaxActions || DateTimeOffset.UtcNow >= deadline)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.BudgetExhausted,
                "Budget reached after backtrack switch; close skipped (no mutating cleanup).",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        // 2) Scope-validate after switch, before close.
        SemanticSnapshot? switchedSnapshot;
        string? scopeAfterSwitch = null;
        try
        {
            var switchedObserve = await _observation
                .ObserveAsync(observationOptions, cancellationToken)
                .ConfigureAwait(false);
            switchedSnapshot = switchedObserve.Success ? switchedObserve.Snapshot : null;
        }
        catch (OperationCanceledException)
        {
            return (false, actionsAttempted);
        }

        if (switchedSnapshot is null
            || !TryValidateScope(switchedSnapshot, policy, identityValidator, out scopeAfterSwitch))
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.ScopeViolation,
                scopeAfterSwitch ?? "Scope validation failed after backtrack switch.",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        if (switchedSnapshot.Window?.Hwnd != frame.ChildHwnd)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.Failure,
                $"Post-switch snapshot HWND 0x{switchedSnapshot.Window?.Hwnd:X} does not equal requested 0x{frame.ChildHwnd:X}.",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        // 3) HWND+PID WindowPattern-only close (driver must not use ActiveWindow / Close button).
        actionsAttempted++; // Count the close attempt even when it fails or times out.
        try
        {
            using var closeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            closeCts.CancelAfter(TimeSpan.FromSeconds(policy.Budgets.PerActionTimeoutSeconds));
            var closeStep = new PlanStep
            {
                Id = "explore-backtrack-close-windowpattern",
                Operation = "closewindow",
                CaptureResponse = true,
                Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                {
                    ["hwnd"] = JsonSerializer.SerializeToElement(frame.ChildHwnd),
                    ["processId"] = JsonSerializer.SerializeToElement(frame.ChildProcessId)
                }
            };
            var closeResponse = await _uiClient
                .ExecuteStepAsync(connection, closeStep, closeCts.Token)
                .ConfigureAwait(false);

            if (!TryReadClosedHwnd(closeResponse.Value, out var closedHwnd)
                || closedHwnd != frame.ChildHwnd)
            {
                recorder.Record(
                    fromStateId,
                    null,
                    fromFingerprint,
                    null,
                    backtrackAction,
                    ExplorationOutcome.Failure,
                    $"HWND-targeted close did not confirm closed HWND 0x{frame.ChildHwnd:X}.",
                    startedAt,
                    DateTimeOffset.UtcNow);
                return (false, actionsAttempted);
            }
        }
        catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ex is OperationCanceledException
                    ? ExplorationOutcome.BudgetExhausted
                    : ExplorationOutcome.Failure,
                SecretRedactor.Redact(ex.Message),
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        // 4) Confirm child HWND gone from inventory before popping.
        WindowInventory inventory;
        try
        {
            inventory = await _windowDetector.ListWindowsAsync(connection, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.Failure,
                SecretRedactor.Redact(ex.Message),
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        var childStillPresent = inventory.Windows.Any(w => w.Hwnd == frame.ChildHwnd);
        if (childStillPresent)
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.Failure,
                $"Child HWND 0x{frame.ChildHwnd:X} still present after WindowPattern close.",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        if (!windowStack.TryPop(out _))
            return (false, actionsAttempted);

        // 5) Switch back to parent HWND when known (counts as mutation).
        if (frame.ParentHwnd is long parentHwnd && parentHwnd > 0
            && actionsAttempted < policy.Budgets.MaxActions
            && DateTimeOffset.UtcNow < deadline)
        {
            var parentPid = frame.ParentProcessId
                            ?? identityValidator.PinnedRootProcessId
                            ?? 0;
            if (parentPid > 0)
            {
                var switchParent = await TrySwitchToHwndAsync(
                        connection,
                        parentHwnd,
                        parentPid,
                        identityValidator,
                        policy.Budgets.PerActionTimeoutSeconds,
                        cancellationToken)
                    .ConfigureAwait(false);
                actionsAttempted++;
                if (!switchParent.Success)
                {
                    _logger.LogWarning(
                        "Backtrack parent switch to HWND 0x{ParentHwnd:X} failed: {Detail}",
                        parentHwnd,
                        switchParent.Detail);
                }
            }
        }

        SemanticSnapshot? afterSnapshot;
        string? afterScopeError = null;
        try
        {
            var afterObserve = await _observation
                .ObserveAsync(observationOptions, cancellationToken)
                .ConfigureAwait(false);
            afterSnapshot = afterObserve.Success ? afterObserve.Snapshot : null;
        }
        catch (OperationCanceledException)
        {
            return (false, actionsAttempted);
        }

        if (afterSnapshot is null
            || !TryValidateScope(afterSnapshot, policy, identityValidator, out afterScopeError))
        {
            recorder.Record(
                fromStateId,
                null,
                fromFingerprint,
                null,
                backtrackAction,
                ExplorationOutcome.ScopeViolation,
                afterScopeError ?? "Scope validation failed after backtrack close.",
                startedAt,
                DateTimeOffset.UtcNow);
            return (false, actionsAttempted);
        }

        var afterFingerprint = _fingerprinter.Fingerprint(afterSnapshot);
        // Preserve the original stacked depth (not depth 0).
        var enqueueDepth = Math.Min(preservedDepth, Math.Max(0, policy.Budgets.MaxDepth - 1));
        var toStateId = recorder.EnsureState(
            afterFingerprint,
            afterSnapshot.SnapshotId,
            afterSnapshot.Window?.Title,
            afterSnapshot.Window?.ClassName,
            afterSnapshot.Window?.IsModal,
            enqueueDepth,
            afterSnapshot.Nodes.Count,
            afterSnapshot.CapturedAtUtc);

        recorder.Record(
            fromStateId,
            toStateId,
            fromFingerprint,
            afterFingerprint,
            backtrackAction,
            ExplorationOutcome.ClosedWindow,
            $"backtrack: closed HWND 0x{frame.ChildHwnd:X} (pid {frame.ChildProcessId}) via WindowPattern",
            startedAt,
            DateTimeOffset.UtcNow);

        if (enqueueDepth < policy.Budgets.MaxDepth
            && recorder.States.Count < policy.Budgets.MaxStates)
        {
            var next = _capabilityResolver.Resolve(
                afterSnapshot,
                afterFingerprint,
                policy,
                enqueueDepth);
            deniedSamples.AddRange(
                _capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
            frontier.EnqueueMany(next);
        }

        return (true, actionsAttempted);
    }

    private async Task<(bool Success, string? Detail)> TrySwitchToHwndAsync(
        DriverConnection connection,
        long hwnd,
        int expectedProcessId,
        ProcessIdentityValidator identityValidator,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));
            var step = new PlanStep
            {
                Id = $"explore-switch-hwnd-0x{hwnd:X}",
                Operation = "switchwindow",
                Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                {
                    ["hwnd"] = JsonSerializer.SerializeToElement(hwnd),
                    ["processId"] = JsonSerializer.SerializeToElement(expectedProcessId)
                }
            };
            await _uiClient.ExecuteStepAsync(connection, step, cts.Token).ConfigureAwait(false);

            var inventory = await _windowDetector.ListWindowsAsync(connection, cts.Token)
                .ConfigureAwait(false);
            var pidCheck = identityValidator.ValidateHwndProcess(hwnd, expectedProcessId, inventory);
            if (!pidCheck.IsValid)
                return (false, pidCheck.Detail);

            return (true, null);
        }
        catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
        {
            return (false, SecretRedactor.Redact(ex.Message));
        }
    }

    private static bool TryReadClosedHwnd(JsonElement? value, out long hwnd)
    {
        hwnd = 0;
        if (value is null
            || value.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
        {
            return false;
        }

        var root = value.Value;
        if (root.ValueKind == JsonValueKind.Object
            && root.TryGetProperty("value", out var nested)
            && nested.ValueKind == JsonValueKind.Object)
        {
            root = nested;
        }

        if (root.ValueKind != JsonValueKind.Object)
            return false;

        if (root.TryGetProperty("hwnd", out var hwndProp)
            || root.TryGetProperty("Hwnd", out hwndProp))
        {
            if (hwndProp.ValueKind == JsonValueKind.Number && hwndProp.TryGetInt64(out hwnd))
                return hwnd > 0;
            if (hwndProp.ValueKind == JsonValueKind.String
                && long.TryParse(hwndProp.GetString(), out hwnd))
            {
                return hwnd > 0;
            }
        }

        return false;
    }

    private bool TryValidateScope(
        SemanticSnapshot snapshot,
        ExplorationPolicyDocument policy,
        ProcessIdentityValidator identityValidator,
        out string error)
    {
        if (!WindowTransitionDetector.IsSnapshotInScope(snapshot, policy))
        {
            error = "Window/process/session is outside policy scope.";
            return false;
        }

        var identity = identityValidator.ValidateSnapshot(snapshot, policy);
        if (!identity.IsValid)
        {
            error = identity.Detail ?? "Process identity validation failed.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    /// <summary>
    /// Resolve a live node without structural/positional fallback. NodeRef is used only when
    /// action.SnapshotId matches the live snapshot; otherwise require an unambiguous
    /// stable-identity match. Caller re-authorizes.
    /// </summary>
    internal static bool TryResolveLiveNode(
        SemanticSnapshot snapshot,
        ExplorationActionCandidate action,
        out SemanticNode? liveNode,
        out string detail)
    {
        liveNode = null;
        detail = string.Empty;

        if (!string.IsNullOrWhiteSpace(action.NodeRef)
            && string.Equals(action.SnapshotId, snapshot.SnapshotId, StringComparison.Ordinal)
            && snapshot.Nodes.FirstOrDefault(n =>
                string.Equals(n.Ref, action.NodeRef, StringComparison.Ordinal)) is { } byRef)
        {
            liveNode = byRef;
            return true;
        }

        var expectedKey = CapabilityResolver.BuildStableIdentityKey(new SemanticNode
        {
            ControlType = action.ControlType ?? string.Empty,
            AutomationId = action.AutomationId,
            Name = action.Name,
            ClassName = action.ClassName
        });

        var matches = snapshot.Nodes
            .Where(n => string.Equals(
                CapabilityResolver.BuildStableIdentityKey(n),
                expectedKey,
                StringComparison.Ordinal))
            .ToList();

        if (matches.Count == 1)
        {
            liveNode = matches[0];
            return true;
        }

        if (matches.Count == 0)
        {
            detail =
                $"Unable to resolve element '{action.ElementKey}' in current snapshot (no unambiguous identity match).";
            return false;
        }

        detail =
            $"Ambiguous identity match for '{action.ElementKey}' ({matches.Count} nodes); fail-closed.";
        return false;
    }

    private static void LinkDeadline(CancellationTokenSource deadlineCts, DateTimeOffset deadline)
    {
        var remaining = deadline - DateTimeOffset.UtcNow;
        if (remaining <= TimeSpan.Zero)
        {
            deadlineCts.Cancel();
            return;
        }

        // CancelAfter rejects > int.MaxValue ms; clamp to a safe upper bound.
        var ms = remaining.TotalMilliseconds;
        if (ms > int.MaxValue)
            ms = int.MaxValue;
        deadlineCts.CancelAfter(TimeSpan.FromMilliseconds(ms));
    }

    private static (string stopReason, string status, int exitCode) CancelOrDeadline(
        CancellationToken cancellationToken) =>
        cancellationToken.IsCancellationRequested
            ? ("cancelled", "cancelled", ExitCodes.Cancelled)
            : ("total-duration-budget", "budget-exhausted", ExitCodes.Success);

    private static (string stopReason, string status, int exitCode) BudgetOrCancel(
        CancellationToken cancellationToken,
        int actionsAttempted,
        int maxActions,
        DateTimeOffset deadline)
    {
        if (cancellationToken.IsCancellationRequested)
            return ("cancelled", "cancelled", ExitCodes.Cancelled);
        if (actionsAttempted >= maxActions)
            return ("max-actions-budget", "budget-exhausted", ExitCodes.Success);
        if (DateTimeOffset.UtcNow >= deadline)
            return ("total-duration-budget", "budget-exhausted", ExitCodes.Success);
        // Linked deadline CTS fired without wall-clock race — treat as duration budget.
        return ("total-duration-budget", "budget-exhausted", ExitCodes.Success);
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
