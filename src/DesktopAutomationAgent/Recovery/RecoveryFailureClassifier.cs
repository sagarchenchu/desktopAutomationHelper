using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Maps a failed step into an explicit recovery classification. Fail-closed for unknowns.
/// Mutation recovery requires typed driver non-execution proof — never free-text alone.
/// </summary>
public sealed class RecoveryFailureClassifier
{
    private readonly RecoveryOptions _options;

    public RecoveryFailureClassifier(Microsoft.Extensions.Options.IOptions<AgentOptions> options)
    {
        _options = options.Value.Recovery;
    }

    /// <summary>Test helper — not registered with DI (avoids ambiguous constructors).</summary>
    public static RecoveryFailureClassifier ForTests(RecoveryOptions? options = null) =>
        new(Microsoft.Extensions.Options.Options.Create(new AgentOptions
        {
            Recovery = options ?? new RecoveryOptions()
        }));

    public RecoveryDecision Classify(
        StepRunResult failedStep,
        RecoveryContext context)
    {
        ArgumentNullException.ThrowIfNull(failedStep);
        ArgumentNullException.ThrowIfNull(context);

        if (failedStep.Classification == UiFailureClassification.AssertionFailure
            || failedStep.Assertions.Any(a => !a.Passed))
        {
            return Deny(RecoveryClassification.AssertionFailure, "Assertion/business-result failure.");
        }

        if (failedStep.Classification == UiFailureClassification.Authentication)
        {
            return Deny(RecoveryClassification.AuthenticationFailure, "Authentication/authorization failure.");
        }

        if (failedStep.Classification is UiFailureClassification.PlanValidation
            or UiFailureClassification.Catalog
            or UiFailureClassification.ArtifactFailure)
        {
            return Deny(RecoveryClassification.CompileSchemaFailure, failedStep.Error ?? "Compile/schema failure.");
        }

        var message = Combine(failedStep.Error, failedStep.DriverReason);

        if (ContainsAny(message, "ambiguous", "multiple matches", "more than one"))
            return Deny(RecoveryClassification.AmbiguousLocator, message);

        if (ContainsAny(message, "destructive", "denied destructive", "prohibited action"))
            return Deny(RecoveryClassification.DestructiveActionDenial, message);

        if (ContainsAny(message, "build mismatch", "fingerprint mismatch", "stale build", "application build"))
            return Deny(RecoveryClassification.ApplicationBuildMismatch, message);

        if (!string.IsNullOrWhiteSpace(context.ExpectedApplicationFingerprint))
        {
            if (string.IsNullOrWhiteSpace(context.ApplicationFingerprint))
            {
                return Deny(
                    RecoveryClassification.ApplicationBuildMismatch,
                    "Expected application fingerprint is configured but live target process identity could not be resolved.");
            }

            if (!string.Equals(
                    context.ExpectedApplicationFingerprint,
                    context.ApplicationFingerprint,
                    StringComparison.OrdinalIgnoreCase))
            {
                return Deny(
                    RecoveryClassification.ApplicationBuildMismatch,
                    "Live application fingerprint does not match approved knowledge.");
            }
        }

        if (ContainsAny(message, "unexpected external", "unexpected window", "foreign process", "outside approved"))
            return Deny(RecoveryClassification.UnexpectedExternalWindow, message);

        if (ContainsAny(message, "unknown operation", "operation is not supported", "not in catalog"))
            return Deny(RecoveryClassification.UnknownOperation, message);

        if (!TryResolveBinding(failedStep.Id, context, out var bindingKey, out var objectRef, out var bindingError))
        {
            return Deny(RecoveryClassification.AmbiguousLocator, bindingError ?? "Ambiguous object bindings.");
        }

        var safeToReplay = RecoveryOperationSafety.IsSafeToReplay(
            failedStep, context.Catalog, _options.AllowMutationReplay);
        var isRead = RecoveryOperationSafety.IsReadOrQuery(failedStep.Operation, context.Catalog);

        // Timed-out / uncertain protocol-or-transport mutations must never enter the ladder.
        if (!safeToReplay
            && RecoveryOperationSafety.IsUncertainExecutionOutcome(failedStep.Classification, message))
        {
            return Deny(
                RecoveryClassification.NotRecoverable,
                "Timed-out or uncertain mutation/invoke/input must not be replayed.");
        }

        if (HasApprovedAlternate(objectRef, context))
        {
            if (!safeToReplay)
            {
                return Deny(
                    RecoveryClassification.NotRecoverable,
                    "Mutation recovery requires typed driver proof (executionState: notStarted or failureCode: locatorNotFoundBeforeInvoke).");
            }

            return Allow(
                RecoveryClassification.ApprovedAlternateLocator,
                "Approved primary locator missing; approved alternate may resolve uniquely.",
                objectRef,
                bindingKey);
        }

        if (HasApprovedCandidate(objectRef, context))
        {
            if (!safeToReplay)
            {
                return Deny(
                    RecoveryClassification.NotRecoverable,
                    "Approved-candidate recovery requires typed driver non-execution proof.");
            }

            return Allow(
                RecoveryClassification.ExplicitlyApprovedCandidateLocator,
                "Explicitly approved candidate locator proposal is available.",
                objectRef,
                bindingKey);
        }

        if (HasEquivalentTransition(failedStep, context, out var transitionEvidence))
        {
            if (!safeToReplay && !isRead)
            {
                return Deny(
                    RecoveryClassification.NotRecoverable,
                    "Equivalent-transition recovery requires typed driver non-execution proof.");
            }

            return Allow(
                RecoveryClassification.ApprovedEquivalentTransition,
                transitionEvidence,
                objectRef,
                bindingKey);
        }

        // Timing / child-window text may classify reads; mutations need typed proof (safeToReplay).
        if (ContainsAny(message, "child window", "dialog not", "window not active", "not yet active", "modal not")
            || string.Equals(
                failedStep.FailureCode,
                RecoveryOperationSafety.FailureCodeLocatorNotFoundBeforeInvoke,
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                failedStep.ExecutionState,
                RecoveryOperationSafety.ExecutionStateNotStarted,
                StringComparison.OrdinalIgnoreCase))
        {
            if (!safeToReplay && !isRead)
            {
                return Deny(
                    RecoveryClassification.NotRecoverable,
                    "Mutation recovery requires typed driver non-execution proof.");
            }

            if (ContainsAny(message, "child window", "dialog not", "window not active", "not yet active", "modal not"))
            {
                return Allow(
                    RecoveryClassification.ExpectedChildWindowNotActive,
                    "Expected child window/dialog not yet active.",
                    objectRef,
                    bindingKey);
            }

            return Allow(
                RecoveryClassification.StateNotStable,
                "State not yet stable / bounded timing drift (typed non-execution proof).",
                objectRef,
                bindingKey);
        }

        if (isRead
            && ContainsAny(
                message,
                "element not found",
                "not found",
                "not stable",
                "stabiliz",
                "timing"))
        {
            // Free-text timing signals remain allowed for read/query retries only.
            return Allow(
                RecoveryClassification.StateNotStable,
                "State not yet stable / bounded timing drift.",
                objectRef,
                bindingKey);
        }

        if (!safeToReplay && !isRead)
        {
            return Deny(
                RecoveryClassification.NotRecoverable,
                "Mutation recovery requires typed driver proof (executionState: notStarted or failureCode: locatorNotFoundBeforeInvoke); free-text errors do not authorize replay.");
        }

        return Deny(RecoveryClassification.NotRecoverable, message.Length == 0 ? "Unclassified failure." : message);
    }

    private static RecoveryDecision Allow(
        RecoveryClassification classification,
        string evidence,
        string? objectReference,
        string? bindingKey) =>
        new()
        {
            Classification = classification,
            MayAttemptRecovery = true,
            Evidence = evidence,
            ObjectReference = objectReference,
            BindingKey = bindingKey
        };

    private static RecoveryDecision Deny(RecoveryClassification classification, string evidence) =>
        new()
        {
            Classification = classification,
            MayAttemptRecovery = false,
            Evidence = evidence
        };

    private static bool TryResolveBinding(
        string stepId,
        RecoveryContext context,
        out string? bindingKey,
        out string? objectRef,
        out string? error)
    {
        bindingKey = null;
        objectRef = null;
        error = null;

        if (!context.StepObjectBindings.TryGetValue(stepId, out var bindings) || bindings.Count == 0)
            return true;

        if (bindings.Count > 1)
        {
            error =
                $"Step '{stepId}' has multiple object bindings ({string.Join(", ", bindings.Keys)}); " +
                "ambiguous recovery target — stop without rewriting locators.";
            return false;
        }

        var pair = bindings.First();
        bindingKey = pair.Key;
        objectRef = pair.Value;
        return true;
    }

    private static bool HasApprovedAlternate(string? objectRef, RecoveryContext context)
    {
        if (string.IsNullOrWhiteSpace(objectRef) || context.RepositorySnapshot is null)
            return false;

        var parts = objectRef.Split('.', 2);
        if (parts.Length != 2)
            return false;

        if (!context.RepositorySnapshot.Pages.TryGetValue(parts[0], out var page))
            return false;

        if (page.Elements is null || !page.Elements.TryGetValue(parts[1], out var element))
            return false;

        return element.AlternateLocators is { Count: > 0 };
    }

    private static bool HasApprovedCandidate(string? objectRef, RecoveryContext context) =>
        !string.IsNullOrWhiteSpace(objectRef)
        && ((context.ApprovedCandidateLocatorIds is not null
             && context.ApprovedCandidateLocatorIds.Contains(objectRef))
            || (context.ApprovedCandidateLocators is not null
                && context.ApprovedCandidateLocators.ContainsKey(objectRef)));

    private static bool HasEquivalentTransition(
        StepRunResult failedStep,
        RecoveryContext context,
        out string evidence)
    {
        evidence = string.Empty;
        if (context.EquivalentTransitions is null || context.EquivalentTransitions.Count == 0)
            return false;

        foreach (var transition in context.EquivalentTransitions.Values)
        {
            if (string.Equals(transition.FromAction, failedStep.Operation, StringComparison.OrdinalIgnoreCase))
            {
                evidence =
                    $"Approved equivalent transition '{transition.TransitionId}' maps '{transition.FromAction}' → '{transition.EquivalentAction}'.";
                return true;
            }
        }

        return false;
    }

    private static string Combine(string? error, string? reason)
    {
        if (string.IsNullOrWhiteSpace(error))
            return reason ?? string.Empty;
        if (string.IsNullOrWhiteSpace(reason))
            return error;
        return error + " " + reason;
    }

    private static bool ContainsAny(string message, params string[] needles)
    {
        if (string.IsNullOrWhiteSpace(message))
            return false;

        foreach (var needle in needles)
        {
            if (message.Contains(needle, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }
}
