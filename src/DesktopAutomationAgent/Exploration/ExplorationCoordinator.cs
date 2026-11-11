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
    private readonly ProcessIdentityValidator _identityValidator;
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
        ProcessIdentityValidator? identityValidator = null)
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
        _identityValidator = identityValidator ?? new ProcessIdentityValidator();
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
            if (!TryValidateScope(initialSnapshot, policy, out var initialScopeError))
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
            var backtrackDepth = 0;

            while (true)
            {
                if (!frontier.TryDequeue(out var action) || action is null)
                {
                    if (policy.AllowCloseChildWindowsForBacktrack
                        && windowStack.Count > 0
                        && recorder.States.Count < policy.Budgets.MaxStates
                        && DateTimeOffset.UtcNow < deadline
                        && !cancellationToken.IsCancellationRequested
                        && !deadlineCts.IsCancellationRequested)
                    {
                        var backtracked = await TryBacktrackAsync(
                                connection,
                                windowStack,
                                policy,
                                observationOptions,
                                frontier,
                                recorder,
                                deniedSamples,
                                backtrackDepth,
                                deadlineCts.Token)
                            .ConfigureAwait(false);
                        if (backtracked)
                        {
                            backtrackDepth = Math.Max(0, backtrackDepth);
                            continue;
                        }
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
                if (!TryValidateScope(beforeSnapshot, policy, out var beforeScopeError))
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
                                var identity = _identityValidator.ValidateSnapshot(snapshot, policy);
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
                if (!TryValidateScope(afterSnapshot, policy, out var afterScopeError))
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

                if (policy.AllowCloseChildWindowsForBacktrack
                    && (outcome is ExplorationOutcome.ModalDialog or ExplorationOutcome.NewWindow))
                {
                    foreach (var added in transition.AddedWindows)
                    {
                        if (string.IsNullOrWhiteSpace(added.Title))
                            continue;

                        windowStack.PushChild(new ExplorationWindowFrame
                        {
                            ChildTitle = added.Title,
                            ChildHwnd = added.Hwnd,
                            ParentTitle = beforeSnapshot.Window?.Title,
                            ParentHwnd = beforeSnapshot.Window?.Hwnd,
                            IsModal = outcome == ExplorationOutcome.ModalDialog
                        });
                    }
                }

                var nextDepth = action.Depth + 1;
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

                // Enqueue only when next depth can still execute (strictly below maxDepth).
                if (nextDepth < policy.Budgets.MaxDepth
                    && recorder.States.Count < policy.Budgets.MaxStates
                    && !string.Equals(beforeFingerprint, afterFingerprint, StringComparison.Ordinal))
                {
                    var next = _capabilityResolver.Resolve(afterSnapshot, afterFingerprint, policy, nextDepth);
                    deniedSamples.AddRange(_capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
                    frontier.EnqueueMany(next);
                }
            }

            // Drain in-scope child windows opened during the crawl when backtrack is allowed.
            if (policy.AllowCloseChildWindowsForBacktrack
                && windowStack.Count > 0
                && !cancellationToken.IsCancellationRequested
                && !deadlineCts.IsCancellationRequested)
            {
                while (windowStack.Count > 0
                       && DateTimeOffset.UtcNow < deadline
                       && !deadlineCts.IsCancellationRequested)
                {
                    var closed = await TryBacktrackAsync(
                            connection,
                            windowStack,
                            policy,
                            observationOptions,
                            frontier,
                            recorder,
                            deniedSamples,
                            backtrackDepth,
                            deadlineCts.Token)
                        .ConfigureAwait(false);
                    if (!closed)
                        break;
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

    private async Task<bool> TryBacktrackAsync(
        DriverConnection connection,
        ExplorationWindowStack windowStack,
        ExplorationPolicyDocument policy,
        ObservationOptions observationOptions,
        FrontierQueue frontier,
        TransitionRecorder recorder,
        List<string> deniedSamples,
        int depth,
        CancellationToken cancellationToken)
    {
        if (!windowStack.TryPop(out var frame) || frame is null)
            return false;

        if (string.IsNullOrWhiteSpace(frame.ChildTitle))
            return false;

        try
        {
            // Prefer switch + dismiss over closewindow: FlaUI closewindow often fails with
            // "No process is associated with this object" for owned/top-level children.
            var switchChild = new PlanStep
            {
                Id = "explore-backtrack-switch-child",
                Operation = "switchwindow",
                Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                {
                    ["value"] = JsonSerializer.SerializeToElement(frame.ChildTitle)
                }
            };
            await _uiClient.ExecuteStepAsync(connection, switchChild, cancellationToken).ConfigureAwait(false);

            var closeButtonName =
                frame.IsModal
                || (!string.IsNullOrWhiteSpace(frame.ChildTitle)
                    && frame.ChildTitle.Contains("Dialog", StringComparison.OrdinalIgnoreCase))
                    ? "Close"
                    : "Close Second";
            var dismiss = new PlanStep
            {
                Id = "explore-backtrack-dismiss",
                Operation = "clickuia",
                Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                {
                    ["locator"] = ObjectLocatorSerializer.ToJsonElement(new ObjectLocator
                    {
                        Name = closeButtonName,
                        ControlType = "Button"
                    }),
                    ["timeoutMs"] = JsonSerializer.SerializeToElement(5_000)
                }
            };
            try
            {
                await _uiClient.ExecuteStepAsync(connection, dismiss, cancellationToken).ConfigureAwait(false);
            }
            catch (UiExecutionException)
            {
                // Fallback: Window-pattern close by title.
                var closeStep = new PlanStep
                {
                    Id = "explore-backtrack-close",
                    Operation = "closewindow",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["value"] = JsonSerializer.SerializeToElement(frame.ChildTitle)
                    }
                };
                await _uiClient.ExecuteStepAsync(connection, closeStep, cancellationToken).ConfigureAwait(false);
            }

            if (!string.IsNullOrWhiteSpace(frame.ParentTitle))
            {
                var switchParent = new PlanStep
                {
                    Id = "explore-backtrack-switch-parent",
                    Operation = "switchwindow",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["value"] = JsonSerializer.SerializeToElement(frame.ParentTitle)
                    }
                };
                await _uiClient.ExecuteStepAsync(connection, switchParent, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (Exception ex) when (ex is UiExecutionException or OperationCanceledException)
        {
            _logger.LogWarning(
                ex,
                "Backtrack close/switch failed for child '{ChildTitle}'.",
                frame.ChildTitle);
            return false;
        }

        var observed = await _observation.ObserveAsync(observationOptions, cancellationToken).ConfigureAwait(false);
        if (!observed.Success || observed.Snapshot is null)
            return false;

        if (!TryValidateScope(observed.Snapshot, policy, out _))
            return false;

        var fingerprint = _fingerprinter.Fingerprint(observed.Snapshot);
        var backtrackDepth = Math.Min(depth, Math.Max(0, policy.Budgets.MaxDepth - 1));
        recorder.EnsureState(
            fingerprint,
            observed.Snapshot.SnapshotId,
            observed.Snapshot.Window?.Title,
            observed.Snapshot.Window?.ClassName,
            observed.Snapshot.Window?.IsModal,
            backtrackDepth,
            observed.Snapshot.Nodes.Count,
            observed.Snapshot.CapturedAtUtc);

        if (backtrackDepth < policy.Budgets.MaxDepth
            && recorder.States.Count < policy.Budgets.MaxStates)
        {
            var next = _capabilityResolver.Resolve(observed.Snapshot, fingerprint, policy, backtrackDepth);
            deniedSamples.AddRange(_capabilityResolver.DeniedActionSamples.Where(s => !deniedSamples.Contains(s)));
            frontier.EnqueueMany(next);
        }

        // Close/switch succeeded even when no new frontier actions (duplicates already executed).
        return true;
    }

    private bool TryValidateScope(
        SemanticSnapshot snapshot,
        ExplorationPolicyDocument policy,
        out string error)
    {
        if (!WindowTransitionDetector.IsSnapshotInScope(snapshot, policy))
        {
            error = "Window/process/session is outside policy scope.";
            return false;
        }

        var identity = _identityValidator.ValidateSnapshot(snapshot, policy);
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
