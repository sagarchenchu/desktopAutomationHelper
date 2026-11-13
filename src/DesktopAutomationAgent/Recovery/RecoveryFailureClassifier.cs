using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Maps a failed step into an explicit recovery classification. Fail-closed for unknowns.
/// </summary>
public sealed class RecoveryFailureClassifier
{
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

        if (!string.IsNullOrWhiteSpace(context.ExpectedApplicationFingerprint)
            && !string.IsNullOrWhiteSpace(context.ApplicationFingerprint)
            && !string.Equals(
                context.ExpectedApplicationFingerprint,
                context.ApplicationFingerprint,
                StringComparison.OrdinalIgnoreCase))
        {
            return Deny(
                RecoveryClassification.ApplicationBuildMismatch,
                "Live application fingerprint does not match approved knowledge.");
        }

        if (ContainsAny(message, "unexpected external", "unexpected window", "foreign process", "outside approved"))
            return Deny(RecoveryClassification.UnexpectedExternalWindow, message);

        if (ContainsAny(message, "unknown operation", "operation is not supported", "not in catalog"))
            return Deny(RecoveryClassification.UnknownOperation, message);

        var objectRef = ResolveObjectRef(failedStep.Id, context);

        if (HasApprovedAlternate(objectRef, context))
        {
            return Allow(
                RecoveryClassification.ApprovedAlternateLocator,
                "Approved primary locator missing; approved alternate may resolve uniquely.",
                objectRef);
        }

        if (HasApprovedCandidate(objectRef, context))
        {
            return Allow(
                RecoveryClassification.ExplicitlyApprovedCandidateLocator,
                "Explicitly approved candidate locator proposal is available.",
                objectRef);
        }

        if (HasEquivalentTransition(failedStep, context, out var transitionEvidence))
        {
            return Allow(
                RecoveryClassification.ApprovedEquivalentTransition,
                transitionEvidence,
                objectRef);
        }

        if (ContainsAny(message, "child window", "dialog not", "window not active", "not yet active", "modal not"))
        {
            return Allow(
                RecoveryClassification.ExpectedChildWindowNotActive,
                "Expected child window/dialog not yet active.",
                objectRef);
        }

        if (ContainsAny(
                message,
                "element not found",
                "not found",
                "timed out",
                "timeout",
                "not stable",
                "stabiliz",
                "timing"))
        {
            return Allow(
                RecoveryClassification.StateNotStable,
                "State not yet stable / bounded timing drift.",
                objectRef);
        }

        return Deny(RecoveryClassification.NotRecoverable, message.Length == 0 ? "Unclassified failure." : message);
    }

    private static RecoveryDecision Allow(
        RecoveryClassification classification,
        string evidence,
        string? objectReference) =>
        new()
        {
            Classification = classification,
            MayAttemptRecovery = true,
            Evidence = evidence,
            ObjectReference = objectReference
        };

    private static RecoveryDecision Deny(RecoveryClassification classification, string evidence) =>
        new()
        {
            Classification = classification,
            MayAttemptRecovery = false,
            Evidence = evidence
        };

    private static string? ResolveObjectRef(string stepId, RecoveryContext context)
    {
        if (!context.StepObjectBindings.TryGetValue(stepId, out var bindings) || bindings.Count == 0)
            return null;

        if (bindings.TryGetValue("locator", out var locatorRef))
            return locatorRef;

        return bindings.Values.FirstOrDefault();
    }

    private static bool HasApprovedAlternate(string? objectRef, RecoveryContext context)
    {
        if (string.IsNullOrWhiteSpace(objectRef) || context.RepositorySnapshot is null)
            return false;

        var pageId = objectRef.Split('.', 2)[0];
        var elementId = objectRef.Contains('.') ? objectRef.Split('.', 2)[1] : null;
        if (elementId is null)
            return false;

        if (!context.RepositorySnapshot.Pages.TryGetValue(pageId, out var page))
            return false;

        if (page.Elements is null || !page.Elements.TryGetValue(elementId, out var element))
            return false;

        return element.AlternateLocators is { Count: > 0 };
    }

    private static bool HasApprovedCandidate(string? objectRef, RecoveryContext context) =>
        !string.IsNullOrWhiteSpace(objectRef)
        && context.ApprovedCandidateLocatorIds is not null
        && context.ApprovedCandidateLocatorIds.Contains(objectRef);

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
