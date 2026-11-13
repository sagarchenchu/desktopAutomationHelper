using System.Diagnostics;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Applies the bounded recovery ladder. Never mutates the original plan hash or approved repos.
/// Budget is consumed per probe/retry/model call. Timed-out mutations are never replayed.
/// </summary>
public sealed class RecoveryLadderService
{
    private readonly RecoveryOptions _options;
    private readonly RecoveryFailureClassifier _classifier;
    private readonly ApprovedAlternateLocatorResolver _alternateResolver;
    private readonly RecoveryArtifactWriter _artifactWriter;
    private readonly CandidateRepairWriter _candidateWriter;
    private readonly IRecoveryModelClient _modelClient;
    private readonly IDriverUiClient _uiClient;
    private readonly ISemanticObservationService? _observation;
    private readonly ObjectReferenceResolver _objectResolver;
    private readonly ILogger<RecoveryLadderService> _logger;

    public RecoveryLadderService(
        IOptions<AgentOptions> options,
        RecoveryFailureClassifier classifier,
        ApprovedAlternateLocatorResolver alternateResolver,
        RecoveryArtifactWriter artifactWriter,
        CandidateRepairWriter candidateWriter,
        IRecoveryModelClient modelClient,
        IDriverUiClient uiClient,
        ObjectReferenceResolver objectResolver,
        ILogger<RecoveryLadderService> logger,
        ISemanticObservationService? observation = null)
    {
        _options = options.Value.Recovery;
        _classifier = classifier;
        _alternateResolver = alternateResolver;
        _artifactWriter = artifactWriter;
        _candidateWriter = candidateWriter;
        _modelClient = modelClient;
        _uiClient = uiClient;
        _objectResolver = objectResolver;
        _observation = observation;
        _logger = logger;
    }

    public RecoveryOptions Options => _options;

    public async Task<RecoveryOutcome> TryRecoverAsync(
        DriverConnection connection,
        PlanStep originalStep,
        StepRunResult failedResult,
        RecoveryContext context,
        RecoveryBudgetTracker budget,
        Func<PlanStep, CancellationToken, Task<StepRunResult>> executeStepAsync,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(connection);
        ArgumentNullException.ThrowIfNull(originalStep);
        ArgumentNullException.ThrowIfNull(failedResult);
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(budget);
        ArgumentNullException.ThrowIfNull(executeStepAsync);

        var decision = _classifier.Classify(failedResult, context);
        if (!decision.MayAttemptRecovery || !RecoveryClassificationRules.IsRecoverable(decision.Classification))
        {
            return new RecoveryOutcome
            {
                Recovered = false,
                Classification = decision.Classification,
                TierUsed = RecoveryTier.None,
                Detail = decision.Evidence
            };
        }

        var attempts = new List<RecoveryAttemptRecord>();
        string? beforePath = null;
        string? recoveryDir = null;
        try
        {
        recoveryDir = _artifactWriter.EnsureRecoveryDirectory(context.RunDirectory);
        beforePath = _artifactWriter.WriteSnapshot(
            recoveryDir,
            TruncateId($"before-{originalStep.Id}-{Guid.NewGuid():N}"),
            new
            {
                stepId = originalStep.Id,
                classification = RecoveryClassificationRules.DisplayName(decision.Classification),
                evidence = decision.Evidence,
                bindingKey = decision.BindingKey,
                error = failedResult.Error,
                originalPlanSha256 = context.OriginalPlanSha256
            });

        // Already-exhausted duration on ladder entry (or mid-ladder) must surface BudgetExhausted,
        // not fall through to the original classification.
        if (budget.RemainingDuration() <= TimeSpan.Zero)
        {
            var exhausted = Record(
                budget.TotalAttempts,
                originalStep.Id,
                RecoveryClassification.BudgetExhausted,
                RecoveryTier.FailAndCleanup,
                "budget-exhausted",
                "max-additional-duration",
                beforePath,
                null,
                DateTimeOffset.UtcNow,
                TimeSpan.Zero);
            attempts.Add(exhausted);
            if (!string.IsNullOrWhiteSpace(recoveryDir))
                _artifactWriter.WriteAttempt(recoveryDir, exhausted);
            return BudgetExhausted(decision, attempts, "max-additional-duration");
        }

        foreach (var tier in LadderFor(decision.Classification))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (budget.RemainingDuration() <= TimeSpan.Zero)
            {
                var exhausted = Record(
                    budget.TotalAttempts,
                    originalStep.Id,
                    RecoveryClassification.BudgetExhausted,
                    tier,
                    "budget-exhausted",
                    "max-additional-duration",
                    beforePath,
                    null,
                    DateTimeOffset.UtcNow,
                    TimeSpan.Zero);
                attempts.Add(exhausted);
                if (!string.IsNullOrWhiteSpace(recoveryDir))
                    _artifactWriter.WriteAttempt(recoveryDir, exhausted);
                return BudgetExhausted(decision, attempts, "max-additional-duration");
            }

            try
            {
            switch (tier)
            {
                case RecoveryTier.FreshObservationAndBoundedWait:
                {
                    if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var token, out var deny))
                    {
                        return BudgetExhausted(decision, attempts, deny);
                    }

                    try
                    {
                        var started = DateTimeOffset.UtcNow;
                        var sw = Stopwatch.StartNew();
                        string? observationPath = null;
                        if (_observation is not null)
                        {
                            try
                            {
                                var observed = await _observation.ObserveAsync(
                                    new ObservationOptions
                                    {
                                        View = "control",
                                        Root = "activeWindow",
                                        MaxDepth = 4,
                                        MaxChildren = 50,
                                        MaxNodes = 200,
                                        IncludeBounds = false,
                                        ActionableOnly = true
                                    },
                                    token).ConfigureAwait(false);
                                observationPath = _artifactWriter.WriteSnapshot(
                                    recoveryDir,
                                    $"observe-{originalStep.Id}-{budget.TotalAttempts:00}",
                                    new
                                    {
                                        observed.Success,
                                        observed.Error,
                                        observed.NodeCount,
                                        observed.SnapshotHash
                                    });
                            }
                            catch (Exception ex) when (ex is not OperationCanceledException)
                            {
                                _logger.LogDebug(ex, "Fresh observation during recovery failed; continuing with wait.");
                            }
                        }

                        var waitMs = Math.Min(
                            _options.BoundedWaitMilliseconds,
                            (int)Math.Max(0, budget.RemainingDuration().TotalMilliseconds));
                        if (waitMs > 0)
                            await Task.Delay(waitMs, token).ConfigureAwait(false);

                        if (!IsReplayAllowed(originalStep, context, failedResult))
                        {
                            var blocked = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                "skipped", "Replay blocked for mutation/uncertain transport failure.",
                                beforePath, observationPath, started, sw.Elapsed);
                            attempts.Add(blocked);
                            _artifactWriter.WriteAttempt(recoveryDir, blocked);
                            break;
                        }

                        var retry = await executeStepAsync(originalStep, token).ConfigureAwait(false);
                        if (IsRecoveryDeadlineFailure(retry, token, cancellationToken))
                        {
                            var budgetRecord = Record(
                                budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                                "budget-exhausted", "max-additional-duration", beforePath, null,
                                DateTimeOffset.UtcNow, TimeSpan.Zero);
                            attempts.Add(budgetRecord);
                            if (!string.IsNullOrWhiteSpace(recoveryDir))
                                _artifactWriter.WriteAttempt(recoveryDir, budgetRecord);
                            return BudgetExhausted(decision, attempts, "max-additional-duration");
                        }
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-wait-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new { retry.Status, retry.Success, retry.Error, observationPath });
                        var record = Record(
                            budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success ? "Fresh observation + bounded wait then re-execute succeeded." : retry.Error,
                            beforePath, afterPath, started, sw.Elapsed);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);
                        if (retry.Success)
                            return Success(decision.Classification, tier, attempts, record.Detail, retry);
                    }
                    finally
                    {
                        budget.EndUnit();
                    }

                    break;
                }

                case RecoveryTier.ReResolveApprovedLocator:
                {
                    if (string.IsNullOrWhiteSpace(decision.ObjectReference)
                        || string.IsNullOrWhiteSpace(decision.BindingKey)
                        || context.RepositorySnapshot is null)
                    {
                        break;
                    }

                    if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var token, out var deny))
                    {
                        return BudgetExhausted(decision, attempts, deny);
                    }

                    try
                    {
                        var started = DateTimeOffset.UtcNow;
                        var sw = Stopwatch.StartNew();
                        var resolution = _objectResolver.Resolve(context.RepositorySnapshot, decision.ObjectReference);
                        if (!resolution.IsResolved || resolution.Locator is null)
                        {
                            var miss = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                "failed", "Re-resolve of approved object reference failed.",
                                beforePath, null, started, sw.Elapsed);
                            attempts.Add(miss);
                            _artifactWriter.WriteAttempt(recoveryDir, miss);
                            break;
                        }

                        var rewritten = _alternateResolver.CloneStepWithLocator(
                            originalStep, resolution.Locator, decision.BindingKey);
                        if (!IsReplayAllowed(rewritten, context, failedResult))
                        {
                            var blocked = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                "skipped", "Replay blocked for mutation/uncertain transport failure.",
                                beforePath, null, started, sw.Elapsed);
                            attempts.Add(blocked);
                            _artifactWriter.WriteAttempt(recoveryDir, blocked);
                            break;
                        }

                        var retry = await executeStepAsync(rewritten, token).ConfigureAwait(false);
                        if (IsRecoveryDeadlineFailure(retry, token, cancellationToken))
                        {
                            var budgetRecord = Record(
                                budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                                "budget-exhausted", "max-additional-duration", beforePath, null,
                                DateTimeOffset.UtcNow, TimeSpan.Zero);
                            attempts.Add(budgetRecord);
                            if (!string.IsNullOrWhiteSpace(recoveryDir))
                                _artifactWriter.WriteAttempt(recoveryDir, budgetRecord);
                            return BudgetExhausted(decision, attempts, "max-additional-duration");
                        }
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-reresolve-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new
                            {
                                objectRef = decision.ObjectReference,
                                bindingKey = decision.BindingKey,
                                retry.Status,
                                retry.Success,
                                retry.Error
                            });
                        var record = Record(
                            budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success ? "Re-resolved approved locator succeeded." : retry.Error,
                            beforePath, afterPath, started, sw.Elapsed);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);
                        if (retry.Success)
                        {
                            return new RecoveryOutcome
                            {
                                Recovered = true,
                                Classification = decision.Classification,
                                TierUsed = tier,
                                Detail = record.Detail,
                                RewrittenStep = rewritten,
                                RecoveredStepResult = retry,
                                Attempts = attempts
                            };
                        }
                    }
                    finally
                    {
                        budget.EndUnit();
                    }

                    break;
                }

                case RecoveryTier.ApprovedAlternateLocator:
                {
                    var alternates = _alternateResolver.GetApprovedAlternates(
                        decision.ObjectReference, context.RepositorySnapshot);
                    if (alternates.Count == 0)
                        break;

                    var bindingKey = decision.BindingKey ?? "locator";
                    ObjectLocator? unique = null;
                    string? uniquenessDetail = null;

                    foreach (var alternate in alternates)
                    {
                        if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var probeToken, out var deny))
                    {
                        return BudgetExhausted(decision, attempts, deny);
                    }

                        try
                        {
                            var probe = await ProbeUniqueAsync(connection, alternate, probeToken).ConfigureAwait(false);
                            if (probeToken.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
                            {
                                return BudgetExhausted(decision, attempts, "max-additional-duration");
                            }

                            if (probe == UniqueProbeResult.Unique)
                            {
                                if (unique is not null)
                                {
                                    unique = null;
                                    uniquenessDetail = "Multiple approved alternates resolved; treating as ambiguous.";
                                    break;
                                }

                                unique = alternate;
                            }
                            else if (probe == UniqueProbeResult.Ambiguous)
                            {
                                unique = null;
                                uniquenessDetail = "Approved alternate resolved ambiguously.";
                                break;
                            }
                        }
                        finally
                        {
                            budget.EndUnit();
                        }

                        if (uniquenessDetail is not null)
                            break;
                    }

                    if (unique is null)
                    {
                        var started = DateTimeOffset.UtcNow;
                        var record = Record(
                            budget.TotalAttempts, originalStep.Id,
                            uniquenessDetail is not null
                                ? RecoveryClassification.AmbiguousLocator
                                : decision.Classification,
                            tier, "failed",
                            uniquenessDetail ?? "No approved alternate resolved uniquely.",
                            beforePath, null, started, TimeSpan.Zero);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);
                        if (uniquenessDetail is not null)
                        {
                            return new RecoveryOutcome
                            {
                                Recovered = false,
                                Classification = RecoveryClassification.AmbiguousLocator,
                                TierUsed = tier,
                                Detail = uniquenessDetail,
                                Attempts = attempts
                            };
                        }

                        break;
                    }

                    if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var execToken, out var execDeny))
                    {
                        return BudgetExhausted(decision, attempts, execDeny);
                    }

                    try
                    {
                        var started = DateTimeOffset.UtcNow;
                        var sw = Stopwatch.StartNew();
                        var rewritten = _alternateResolver.CloneStepWithLocator(originalStep, unique, bindingKey);
                        if (!IsReplayAllowed(rewritten, context, failedResult))
                        {
                            var blocked = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                "skipped", "Alternate replay blocked for mutation/uncertain protocol failure.",
                                beforePath, null, started, sw.Elapsed);
                            attempts.Add(blocked);
                            _artifactWriter.WriteAttempt(recoveryDir, blocked);
                            break;
                        }

                        var retry = await executeStepAsync(rewritten, execToken).ConfigureAwait(false);
                        if (IsRecoveryDeadlineFailure(retry, execToken, cancellationToken))
                        {
                            var budgetRecord = Record(
                                budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                                "budget-exhausted", "max-additional-duration", beforePath, null,
                                DateTimeOffset.UtcNow, TimeSpan.Zero);
                            attempts.Add(budgetRecord);
                            if (!string.IsNullOrWhiteSpace(recoveryDir))
                                _artifactWriter.WriteAttempt(recoveryDir, budgetRecord);
                            return BudgetExhausted(decision, attempts, "max-additional-duration");
                        }
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-alternate-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new
                            {
                                bindingKey,
                                alternate = ObjectLocatorSerializer.ToJsonElement(unique),
                                retry.Status,
                                retry.Success,
                                retry.Error
                            });
                        var altRecord = Record(
                            budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success ? "Approved alternate locator recovered the step." : retry.Error,
                            beforePath, afterPath, started, sw.Elapsed);
                        attempts.Add(altRecord);
                        _artifactWriter.WriteAttempt(recoveryDir, altRecord);
                        if (retry.Success)
                        {
                            return new RecoveryOutcome
                            {
                                Recovered = true,
                                Classification = decision.Classification,
                                TierUsed = tier,
                                Detail = altRecord.Detail,
                                RewrittenStep = rewritten,
                                RecoveredStepResult = retry,
                                Attempts = attempts
                            };
                        }
                    }
                    finally
                    {
                        budget.EndUnit();
                    }

                    break;
                }

                case RecoveryTier.ApprovedEquivalentTransition:
                {
                    if (context.EquivalentTransitions is null
                        || !context.EquivalentTransitions.Values.Any(t =>
                            string.Equals(t.FromAction, originalStep.Operation, StringComparison.OrdinalIgnoreCase)))
                    {
                        break;
                    }

                    if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var token, out var deny))
                    {
                        return BudgetExhausted(decision, attempts, deny);
                    }

                    try
                    {
                        var started = DateTimeOffset.UtcNow;
                        var sw = Stopwatch.StartNew();
                        var transition = context.EquivalentTransitions.Values.First(t =>
                            string.Equals(t.FromAction, originalStep.Operation, StringComparison.OrdinalIgnoreCase));
                        var rewritten = _alternateResolver.CloneStepWithEquivalentTransition(originalStep, transition);
                        if (rewritten is null)
                            break;

                        var failureMessage = Combine(failedResult.Error, failedResult.DriverReason);
                        if (!RecoveryOperationSafety.IsEquivalentTransitionAllowed(
                                originalStep.Operation,
                                rewritten.Operation,
                                context.Catalog,
                                failedResult.Classification,
                                failureMessage,
                                _options.AllowMutationReplay,
                                failedResult.ExecutionState,
                                failedResult.FailureCode))
                        {
                            var blocked = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                "skipped",
                                "Equivalent transition blocked: cannot upgrade to mutation without AllowMutationReplay / non-execution proof.",
                                beforePath, null, started, sw.Elapsed);
                            attempts.Add(blocked);
                            _artifactWriter.WriteAttempt(recoveryDir, blocked);
                            break;
                        }

                        var retry = await executeStepAsync(rewritten, token).ConfigureAwait(false);
                        if (IsRecoveryDeadlineFailure(retry, token, cancellationToken))
                        {
                            var budgetRecord = Record(
                                budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                                "budget-exhausted", "max-additional-duration", beforePath, null,
                                DateTimeOffset.UtcNow, TimeSpan.Zero);
                            attempts.Add(budgetRecord);
                            if (!string.IsNullOrWhiteSpace(recoveryDir))
                                _artifactWriter.WriteAttempt(recoveryDir, budgetRecord);
                            return BudgetExhausted(decision, attempts, "max-additional-duration");
                        }
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-transition-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new { transition.TransitionId, transition.EquivalentAction, retry.Status, retry.Success, retry.Error });
                        var record = Record(
                            budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success
                                ? $"Approved equivalent transition '{transition.TransitionId}' recovered the step."
                                : retry.Error,
                            beforePath, afterPath, started, sw.Elapsed);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);
                        if (retry.Success)
                        {
                            return new RecoveryOutcome
                            {
                                Recovered = true,
                                Classification = decision.Classification,
                                TierUsed = tier,
                                Detail = record.Detail,
                                RewrittenStep = rewritten,
                                RecoveredStepResult = retry,
                                Attempts = attempts
                            };
                        }
                    }
                    finally
                    {
                        budget.EndUnit();
                    }

                    break;
                }

                case RecoveryTier.CandidateAiRepairProposal:
                {
                    // Explicitly approved candidate: apply after unique verification — do not fall through.
                    if (decision.Classification == RecoveryClassification.ExplicitlyApprovedCandidateLocator
                        && !string.IsNullOrWhiteSpace(decision.ObjectReference)
                        && context.ApprovedCandidateLocators is not null
                        && context.ApprovedCandidateLocators.TryGetValue(decision.ObjectReference, out var approvedLocator))
                    {
                        if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var probeToken, out var deny))
                    {
                        return BudgetExhausted(decision, attempts, deny);
                    }

                        UniqueProbeResult probe;
                        try
                        {
                            probe = await ProbeUniqueAsync(connection, approvedLocator, probeToken).ConfigureAwait(false);
                            if (probeToken.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
                            {
                                return BudgetExhausted(decision, attempts, "max-additional-duration");
                            }
                        }
                        finally
                        {
                            budget.EndUnit();
                        }

                        if (probe != UniqueProbeResult.Unique)
                        {
                            var started = DateTimeOffset.UtcNow;
                            var fail = Record(
                                budget.TotalAttempts, originalStep.Id,
                                probe == UniqueProbeResult.Ambiguous
                                    ? RecoveryClassification.AmbiguousLocator
                                    : decision.Classification,
                                tier, "failed",
                                probe == UniqueProbeResult.Ambiguous
                                    ? "Approved candidate locator resolved ambiguously."
                                    : "Approved candidate locator did not resolve uniquely.",
                                beforePath, null, started, TimeSpan.Zero);
                            attempts.Add(fail);
                            _artifactWriter.WriteAttempt(recoveryDir, fail);
                            return new RecoveryOutcome
                            {
                                Recovered = false,
                                Classification = fail.Classification,
                                TierUsed = tier,
                                Detail = fail.Detail,
                                Attempts = attempts
                            };
                        }

                        if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var execToken, out var execDeny))
                    {
                        return BudgetExhausted(decision, attempts, execDeny);
                    }

                        try
                        {
                            var started = DateTimeOffset.UtcNow;
                            var sw = Stopwatch.StartNew();
                            var bindingKey = decision.BindingKey ?? "locator";
                            var rewritten = _alternateResolver.CloneStepWithLocator(originalStep, approvedLocator, bindingKey);
                            if (!IsReplayAllowed(rewritten, context, failedResult))
                            {
                                var blocked = Record(
                                    budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                    "skipped", "Approved-candidate replay blocked for mutation/uncertain protocol failure.",
                                    beforePath, null, started, sw.Elapsed);
                                attempts.Add(blocked);
                                _artifactWriter.WriteAttempt(recoveryDir, blocked);
                                return new RecoveryOutcome
                                {
                                    Recovered = false,
                                    Classification = decision.Classification,
                                    TierUsed = tier,
                                    Detail = blocked.Detail,
                                    Attempts = attempts
                                };
                            }

                            var retry = await executeStepAsync(rewritten, execToken).ConfigureAwait(false);
                        if (IsRecoveryDeadlineFailure(retry, execToken, cancellationToken))
                        {
                            var budgetRecord = Record(
                                budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                                "budget-exhausted", "max-additional-duration", beforePath, null,
                                DateTimeOffset.UtcNow, TimeSpan.Zero);
                            attempts.Add(budgetRecord);
                            if (!string.IsNullOrWhiteSpace(recoveryDir))
                                _artifactWriter.WriteAttempt(recoveryDir, budgetRecord);
                            return BudgetExhausted(decision, attempts, "max-additional-duration");
                        }
                            var afterPath = _artifactWriter.WriteSnapshot(
                                recoveryDir,
                                $"after-approved-candidate-{originalStep.Id}-{budget.TotalAttempts:00}",
                                new { bindingKey, retry.Status, retry.Success, retry.Error });
                            var record = Record(
                                budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                                retry.Success ? "recovered" : "failed",
                                retry.Success
                                    ? "Explicitly approved candidate locator recovered the step."
                                    : retry.Error,
                                beforePath, afterPath, started, sw.Elapsed);
                            attempts.Add(record);
                            _artifactWriter.WriteAttempt(recoveryDir, record);
                            if (retry.Success)
                            {
                                return new RecoveryOutcome
                                {
                                    Recovered = true,
                                    Classification = decision.Classification,
                                    TierUsed = tier,
                                    Detail = record.Detail,
                                    RewrittenStep = rewritten,
                                    RecoveredStepResult = retry,
                                    Attempts = attempts
                                };
                            }

                            return new RecoveryOutcome
                            {
                                Recovered = false,
                                Classification = decision.Classification,
                                TierUsed = tier,
                                Detail = record.Detail,
                                Attempts = attempts
                            };
                        }
                        finally
                        {
                            budget.EndUnit();
                        }
                    }

                    if (!TryConsume(budget, originalStep.Id, cancellationToken, attempts, decision, tier, beforePath,
                            _artifactWriter, recoveryDir, out var modelToken, out var modelDeny))
                    {
                        return BudgetExhausted(decision, attempts, modelDeny);
                    }

                    try
                    {
                        var started = DateTimeOffset.UtcNow;
                        var sw = Stopwatch.StartNew();
                        var proposal = await _modelClient.ProposeRepairAsync(
                            new RecoveryRepairRequest
                            {
                                StepId = originalStep.Id,
                                Operation = originalStep.Operation,
                                FailureEvidence = decision.Evidence ?? failedResult.Error ?? "recovery",
                                Classification = decision.Classification,
                                ObjectReference = decision.ObjectReference,
                                ApplicationFingerprint = context.ApplicationFingerprint,
                                ObservationSummary = "bounded-recovery"
                            },
                            modelToken).ConfigureAwait(false);

                        string? candidatePath = null;
                        if (proposal is not null)
                        {
                            proposal.AutoPromoted = false;
                            proposal.RequiresHumanApproval = true;
                            candidatePath = _candidateWriter.Write(proposal);
                        }

                        var record = Record(
                            budget.TotalAttempts, originalStep.Id, decision.Classification, tier,
                            "stop-for-approval",
                            proposal is null
                                ? "No candidate repair proposed."
                                : "Candidate AI repair written; human approval required (never auto-promoted).",
                            beforePath, null, started, sw.Elapsed, candidatePath);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);
                        return new RecoveryOutcome
                        {
                            Recovered = false,
                            StopForApproval = true,
                            Classification = decision.Classification,
                            TierUsed = tier,
                            Detail = record.Detail,
                            CandidateRepairPath = candidatePath,
                            Attempts = attempts
                        };
                    }
                    finally
                    {
                        budget.EndUnit();
                    }
                }
            }
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                // Linked recovery-deadline CTS cancelled mid-unit — map to budget exhausted
                // so the runner always receives a normal outcome (summary + cleanup).
                if (budget.InRecovery)
                    budget.EndUnit();

                var started = DateTimeOffset.UtcNow;
                var record = Record(
                    budget.TotalAttempts, originalStep.Id, RecoveryClassification.BudgetExhausted, tier,
                    "budget-exhausted", "max-additional-duration", beforePath, null, started, TimeSpan.Zero);
                attempts.Add(record);
                _artifactWriter.WriteAttempt(recoveryDir, record);
                return BudgetExhausted(decision, attempts, "max-additional-duration");
            }
        }

        var failRecord = Record(
            budget.TotalAttempts, originalStep.Id, decision.Classification,
            RecoveryTier.FailAndCleanup, "failed",
            "Recovery ladder exhausted; executing original cleanup.",
            beforePath, null, DateTimeOffset.UtcNow, TimeSpan.Zero);
        attempts.Add(failRecord);
        if (!string.IsNullOrWhiteSpace(recoveryDir))
            _artifactWriter.WriteAttempt(recoveryDir, failRecord);
        return new RecoveryOutcome
        {
            Recovered = false,
            Classification = decision.Classification,
            TierUsed = RecoveryTier.FailAndCleanup,
            Detail = failRecord.Detail,
            Attempts = attempts
        };
        }
        catch (Exception ex) when (ex is IOException or RecoveryPathException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            // Artifact/path failures must not escape TryRecoverAsync — runner must still cleanup.
            if (budget.InRecovery)
                budget.EndUnit();

            var started = DateTimeOffset.UtcNow;
            var record = Record(
                budget.TotalAttempts,
                originalStep.Id,
                RecoveryClassification.NotRecoverable,
                RecoveryTier.FailAndCleanup,
                "artifact-failure",
                $"Recovery artifact failure: {ex.Message}",
                beforePath,
                null,
                started,
                TimeSpan.Zero);
            attempts.Add(record);
            try
            {
                if (!string.IsNullOrWhiteSpace(recoveryDir))
                    _artifactWriter.WriteAttempt(recoveryDir, record);
            }
            catch
            {
                // best-effort
            }

            return new RecoveryOutcome
            {
                Recovered = false,
                Classification = RecoveryClassification.NotRecoverable,
                TierUsed = RecoveryTier.FailAndCleanup,
                Detail = record.Detail,
                Attempts = attempts
            };
        }
    }

    private bool IsReplayAllowed(PlanStep step, RecoveryContext context, StepRunResult failedResult) =>
        RecoveryOperationSafety.IsSafeToReplay(
            step.Operation,
            context.Catalog,
            failedResult.Classification,
            Combine(failedResult.Error, failedResult.DriverReason),
            _options.AllowMutationReplay,
            failedResult.ExecutionState,
            failedResult.FailureCode);

    private static bool IsRecoveryDeadlineFailure(
        StepRunResult result,
        CancellationToken recoveryToken,
        CancellationToken parentToken) =>
        recoveryToken.IsCancellationRequested
        && !parentToken.IsCancellationRequested
        && result.Classification is UiFailureClassification.Cancelled
            or UiFailureClassification.ExecutionTimeout;

    private static string Combine(string? a, string? b)
    {
        if (string.IsNullOrWhiteSpace(a)) return b ?? string.Empty;
        if (string.IsNullOrWhiteSpace(b)) return a;
        return a + " " + b;
    }

    private static bool TryConsume(
        RecoveryBudgetTracker budget,
        string stepId,
        CancellationToken parent,
        List<RecoveryAttemptRecord> attempts,
        RecoveryDecision decision,
        RecoveryTier tier,
        string? beforePath,
        RecoveryArtifactWriter artifactWriter,
        string recoveryDir,
        out CancellationToken token,
        out string deny)
    {
        if (budget.TryConsumeUnit(stepId, parent, out token, out deny))
            return true;

        var started = DateTimeOffset.UtcNow;
        var record = Record(
            budget.TotalAttempts, stepId, RecoveryClassification.BudgetExhausted, tier,
            "budget-exhausted", deny, beforePath, null, started, TimeSpan.Zero);
        attempts.Add(record);
        artifactWriter.WriteAttempt(recoveryDir, record);
        return false;
    }

    private static RecoveryOutcome BudgetExhausted(
        RecoveryDecision decision,
        List<RecoveryAttemptRecord> attempts,
        string deny) =>
        new()
        {
            Recovered = false,
            Classification = RecoveryClassification.BudgetExhausted,
            TierUsed = RecoveryTier.FailAndCleanup,
            Detail = deny,
            Attempts = attempts
        };

    private async Task<UniqueProbeResult> ProbeUniqueAsync(
        DriverConnection connection,
        ObjectLocator locator,
        CancellationToken cancellationToken)
    {
        var probeStep = new PlanStep
        {
            Id = "recovery-finduia-probe",
            Operation = "finduia",
            Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
            {
                ["locator"] = ObjectLocatorSerializer.ToJsonElement(locator)
            }
        };

        try
        {
            var response = await _uiClient.ExecuteStepAsync(connection, probeStep, cancellationToken)
                .ConfigureAwait(false);
            if (!response.Success)
            {
                var detail = response.Error ?? response.Reason ?? string.Empty;
                if (detail.Contains("ambiguous", StringComparison.OrdinalIgnoreCase)
                    || detail.Contains("multiple", StringComparison.OrdinalIgnoreCase))
                {
                    return UniqueProbeResult.Ambiguous;
                }

                return UniqueProbeResult.Missing;
            }

            if (response.Value is JsonElement value
                && value.ValueKind == JsonValueKind.Object
                && value.TryGetProperty("matchCount", out var countEl)
                && countEl.TryGetInt32(out var count))
            {
                if (count > 1) return UniqueProbeResult.Ambiguous;
                if (count == 1) return UniqueProbeResult.Unique;
                return UniqueProbeResult.Missing;
            }

            return UniqueProbeResult.Unique;
        }
        catch (UiExecutionException ex)
        {
            var detail = ex.Message + " " + (ex.Response?.Reason ?? string.Empty);
            if (detail.Contains("ambiguous", StringComparison.OrdinalIgnoreCase)
                || detail.Contains("multiple", StringComparison.OrdinalIgnoreCase))
            {
                return UniqueProbeResult.Ambiguous;
            }

            _logger.LogDebug(ex, "Alternate locator probe did not resolve uniquely.");
            return UniqueProbeResult.Missing;
        }
    }

    private static IEnumerable<RecoveryTier> LadderFor(RecoveryClassification classification) =>
        classification switch
        {
            // Classification already selected the recovery path — do not burn per-unit budget
            // on unrelated earlier tiers before the matching one.
            RecoveryClassification.ApprovedAlternateLocator =>
            [
                RecoveryTier.ApprovedAlternateLocator,
                RecoveryTier.CandidateAiRepairProposal
            ],
            RecoveryClassification.ApprovedEquivalentTransition =>
            [
                RecoveryTier.ApprovedEquivalentTransition,
                RecoveryTier.CandidateAiRepairProposal
            ],
            RecoveryClassification.ExplicitlyApprovedCandidateLocator =>
            [
                // Apply the approved candidate (handled inside CandidateAiRepairProposal case).
                RecoveryTier.CandidateAiRepairProposal
            ],
            RecoveryClassification.ExpectedChildWindowNotActive =>
            [
                RecoveryTier.FreshObservationAndBoundedWait,
                RecoveryTier.ReResolveApprovedLocator,
                RecoveryTier.CandidateAiRepairProposal
            ],
            _ =>
            [
                RecoveryTier.FreshObservationAndBoundedWait,
                RecoveryTier.ReResolveApprovedLocator,
                RecoveryTier.ApprovedAlternateLocator,
                RecoveryTier.ApprovedEquivalentTransition,
                RecoveryTier.CandidateAiRepairProposal
            ]
        };

    private static RecoveryOutcome Success(
        RecoveryClassification classification,
        RecoveryTier tier,
        List<RecoveryAttemptRecord> attempts,
        string? detail,
        StepRunResult recoveredStep) =>
        new()
        {
            Recovered = true,
            Classification = classification,
            TierUsed = tier,
            Detail = detail,
            RecoveredStepResult = recoveredStep,
            Attempts = attempts
        };

    private static RecoveryAttemptRecord Record(
        int attemptNumber,
        string stepId,
        RecoveryClassification classification,
        RecoveryTier tier,
        string outcome,
        string? detail,
        string? before,
        string? after,
        DateTimeOffset started,
        TimeSpan duration,
        string? candidatePath = null) =>
        new()
        {
            AttemptNumber = attemptNumber,
            StepId = stepId,
            Classification = classification,
            Tier = tier,
            Outcome = outcome,
            Detail = detail,
            SnapshotBeforePath = before,
            SnapshotAfterPath = after,
            CandidateRepairPath = candidatePath,
            StartedAtUtc = started,
            FinishedAtUtc = DateTimeOffset.UtcNow,
            DurationMilliseconds = duration.TotalMilliseconds
        };

    private static string TruncateId(string value, int max = 48) =>
        value.Length <= max ? value : value[..max];

    private enum UniqueProbeResult
    {
        Missing,
        Unique,
        Ambiguous
    }
}
