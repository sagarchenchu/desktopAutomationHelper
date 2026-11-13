using System.Diagnostics;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Applies the bounded recovery ladder. Never mutates the original plan hash or approved repos.
/// Candidate AI repairs stop for approval and never auto-promote.
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
    private readonly ILogger<RecoveryLadderService> _logger;

    public RecoveryLadderService(
        IOptions<AgentOptions> options,
        RecoveryFailureClassifier classifier,
        ApprovedAlternateLocatorResolver alternateResolver,
        RecoveryArtifactWriter artifactWriter,
        CandidateRepairWriter candidateWriter,
        IRecoveryModelClient modelClient,
        IDriverUiClient uiClient,
        ILogger<RecoveryLadderService> logger)
    {
        _options = options.Value.Recovery;
        _classifier = classifier;
        _alternateResolver = alternateResolver;
        _artifactWriter = artifactWriter;
        _candidateWriter = candidateWriter;
        _modelClient = modelClient;
        _uiClient = uiClient;
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

        if (!budget.TryBeginAttempt(originalStep.Id, out var denyReason))
        {
            return new RecoveryOutcome
            {
                Recovered = false,
                Classification = RecoveryClassification.BudgetExhausted,
                TierUsed = RecoveryTier.FailAndCleanup,
                Detail = denyReason
            };
        }

        var attempts = new List<RecoveryAttemptRecord>();
        var recoveryDir = _artifactWriter.EnsureRecoveryDirectory(context.RunDirectory);

        try
        {
            var beforePath = _artifactWriter.WriteSnapshot(
                recoveryDir,
                $"before-{originalStep.Id}-{budget.TotalAttempts:00}",
                new
                {
                    stepId = originalStep.Id,
                    classification = RecoveryClassificationRules.DisplayName(decision.Classification),
                    evidence = decision.Evidence,
                    error = failedResult.Error,
                    driverReason = failedResult.DriverReason,
                    originalPlanSha256 = context.OriginalPlanSha256
                });

            // Tier order depends on classification, but never skips budget or approval gates.
            foreach (var tier in LadderFor(decision.Classification))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (budget.RemainingDuration() <= TimeSpan.Zero)
                    break;

                var started = DateTimeOffset.UtcNow;
                var sw = Stopwatch.StartNew();

                switch (tier)
                {
                    case RecoveryTier.FreshObservationAndBoundedWait:
                    {
                        var waitMs = Math.Min(
                            _options.BoundedWaitMilliseconds,
                            (int)Math.Max(0, budget.RemainingDuration().TotalMilliseconds));
                        if (waitMs > 0)
                            await Task.Delay(waitMs, cancellationToken).ConfigureAwait(false);

                        var retry = await executeStepAsync(originalStep, cancellationToken).ConfigureAwait(false);
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-wait-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new { retry.Status, retry.Success, retry.Error, retry.DriverReason });

                        var record = Record(
                            budget.TotalAttempts,
                            originalStep.Id,
                            decision.Classification,
                            tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success ? "Bounded wait then re-execute succeeded." : retry.Error,
                            beforePath,
                            afterPath,
                            started,
                            sw.Elapsed);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);

                        if (retry.Success)
                        {
                            return Success(decision.Classification, tier, attempts, record.Detail, retry);
                        }

                        break;
                    }

                    case RecoveryTier.ReResolveApprovedLocator:
                    {
                        // Re-resolve uses the same approved locator already bound on the step.
                        var retry = await executeStepAsync(originalStep, cancellationToken).ConfigureAwait(false);
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-reresolve-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new { retry.Status, retry.Success, retry.Error });

                        var record = Record(
                            budget.TotalAttempts,
                            originalStep.Id,
                            decision.Classification,
                            tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success ? "Re-resolved approved locator succeeded." : retry.Error,
                            beforePath,
                            afterPath,
                            started,
                            sw.Elapsed);
                        attempts.Add(record);
                        _artifactWriter.WriteAttempt(recoveryDir, record);

                        if (retry.Success)
                            return Success(decision.Classification, tier, attempts, record.Detail, retry);
                        break;
                    }

                    case RecoveryTier.ApprovedAlternateLocator:
                    {
                        var alternates = _alternateResolver.GetApprovedAlternates(
                            decision.ObjectReference,
                            context.RepositorySnapshot);
                        if (alternates.Count == 0)
                        {
                            attempts.Add(Record(
                                budget.TotalAttempts,
                                originalStep.Id,
                                decision.Classification,
                                tier,
                                "skipped",
                                "No approved alternate locators.",
                                beforePath,
                                null,
                                started,
                                sw.Elapsed));
                            break;
                        }

                        ObjectLocator? unique = null;
                        string? uniquenessDetail = null;
                        foreach (var alternate in alternates)
                        {
                            var probe = await ProbeUniqueAsync(connection, alternate, cancellationToken)
                                .ConfigureAwait(false);
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

                        if (unique is null)
                        {
                            var record = Record(
                                budget.TotalAttempts,
                                originalStep.Id,
                                uniquenessDetail is not null
                                    ? RecoveryClassification.AmbiguousLocator
                                    : decision.Classification,
                                tier,
                                "failed",
                                uniquenessDetail ?? "No approved alternate resolved uniquely.",
                                beforePath,
                                null,
                                started,
                                sw.Elapsed);
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

                        var rewritten = _alternateResolver.CloneStepWithLocator(originalStep, unique);
                        var retry = await executeStepAsync(rewritten, cancellationToken).ConfigureAwait(false);
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-alternate-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new
                            {
                                alternate = ObjectLocatorSerializer.ToJsonElement(unique),
                                retry.Status,
                                retry.Success,
                                retry.Error
                            });

                        var altRecord = Record(
                            budget.TotalAttempts,
                            originalStep.Id,
                            decision.Classification,
                            tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success
                                ? "Approved alternate locator recovered the step."
                                : retry.Error,
                            beforePath,
                            afterPath,
                            started,
                            sw.Elapsed);
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

                        break;
                    }

                    case RecoveryTier.ApprovedEquivalentTransition:
                    {
                        if (context.EquivalentTransitions is null
                            || !context.EquivalentTransitions.Values.Any(t =>
                                string.Equals(t.FromAction, originalStep.Operation, StringComparison.OrdinalIgnoreCase)))
                        {
                            attempts.Add(Record(
                                budget.TotalAttempts,
                                originalStep.Id,
                                decision.Classification,
                                tier,
                                "skipped",
                                "No approved equivalent transition.",
                                beforePath,
                                null,
                                started,
                                sw.Elapsed));
                            break;
                        }

                        var transition = context.EquivalentTransitions.Values.First(t =>
                            string.Equals(t.FromAction, originalStep.Operation, StringComparison.OrdinalIgnoreCase));
                        var rewritten = _alternateResolver.CloneStepWithEquivalentTransition(originalStep, transition);
                        if (rewritten is null)
                            break;

                        var retry = await executeStepAsync(rewritten, cancellationToken).ConfigureAwait(false);
                        var afterPath = _artifactWriter.WriteSnapshot(
                            recoveryDir,
                            $"after-transition-{originalStep.Id}-{budget.TotalAttempts:00}",
                            new
                            {
                                transition.TransitionId,
                                transition.EquivalentAction,
                                retry.Status,
                                retry.Success,
                                retry.Error
                            });

                        var record = Record(
                            budget.TotalAttempts,
                            originalStep.Id,
                            decision.Classification,
                            tier,
                            retry.Success ? "recovered" : "failed",
                            retry.Success
                                ? $"Approved equivalent transition '{transition.TransitionId}' recovered the step."
                                : retry.Error,
                            beforePath,
                            afterPath,
                            started,
                            sw.Elapsed);
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

                        break;
                    }

                    case RecoveryTier.CandidateAiRepairProposal:
                    {
                        // Explicitly approved candidate locators may be applied; otherwise propose and stop.
                        if (decision.Classification == RecoveryClassification.ExplicitlyApprovedCandidateLocator
                            && !string.IsNullOrWhiteSpace(decision.ObjectReference)
                            && context.ApprovedCandidateLocatorIds is not null
                            && context.ApprovedCandidateLocatorIds.Contains(decision.ObjectReference)
                            && context.RepositorySnapshot is not null)
                        {
                            // Already-approved candidate path: still require unique verification via primary page element
                            // alternates if present; otherwise stop for approval rather than inventing a locator.
                        }

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
                            cancellationToken).ConfigureAwait(false);

                        string? candidatePath = null;
                        if (proposal is not null)
                        {
                            proposal.AutoPromoted = false;
                            proposal.RequiresHumanApproval = true;
                            candidatePath = _candidateWriter.Write(proposal);
                        }

                        var record = Record(
                            budget.TotalAttempts,
                            originalStep.Id,
                            decision.Classification,
                            tier,
                            "stop-for-approval",
                            proposal is null
                                ? "No candidate repair proposed."
                                : "Candidate AI repair written; human approval required (never auto-promoted).",
                            beforePath,
                            null,
                            started,
                            sw.Elapsed,
                            candidatePath);
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
                }
            }

            var failRecord = Record(
                budget.TotalAttempts,
                originalStep.Id,
                decision.Classification,
                RecoveryTier.FailAndCleanup,
                "failed",
                "Recovery ladder exhausted; executing original cleanup.",
                beforePath,
                null,
                DateTimeOffset.UtcNow,
                TimeSpan.Zero);
            attempts.Add(failRecord);
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
        finally
        {
            budget.EndAttempt();
        }
    }

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
                var detail = (response.Error ?? response.Reason ?? string.Empty);
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
                if (count > 1)
                    return UniqueProbeResult.Ambiguous;
                if (count == 1)
                    return UniqueProbeResult.Unique;
                return UniqueProbeResult.Missing;
            }

            // Successful finduia without explicit matchCount is treated as unique.
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
            RecoveryClassification.ApprovedAlternateLocator =>
            [
                RecoveryTier.FreshObservationAndBoundedWait,
                RecoveryTier.ReResolveApprovedLocator,
                RecoveryTier.ApprovedAlternateLocator,
                RecoveryTier.CandidateAiRepairProposal
            ],
            RecoveryClassification.ApprovedEquivalentTransition =>
            [
                RecoveryTier.FreshObservationAndBoundedWait,
                RecoveryTier.ApprovedEquivalentTransition,
                RecoveryTier.CandidateAiRepairProposal
            ],
            RecoveryClassification.ExplicitlyApprovedCandidateLocator =>
            [
                RecoveryTier.FreshObservationAndBoundedWait,
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

    private enum UniqueProbeResult
    {
        Missing,
        Unique,
        Ambiguous
    }
}
