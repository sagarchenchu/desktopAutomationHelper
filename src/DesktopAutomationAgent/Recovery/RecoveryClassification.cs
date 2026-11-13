namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Explicit recovery classifications. Only the recoverable set may enter the ladder.
/// </summary>
public enum RecoveryClassification
{
    // Recoverable
    StateNotStable = 1,
    ExpectedChildWindowNotActive = 2,
    ApprovedAlternateLocator = 3,
    ApprovedEquivalentTransition = 4,
    ExplicitlyApprovedCandidateLocator = 5,

    // Not automatically recoverable
    AssertionFailure = 100,
    AuthenticationFailure = 101,
    DestructiveActionDenial = 102,
    AmbiguousLocator = 103,
    ApplicationBuildMismatch = 104,
    UnexpectedExternalWindow = 105,
    UnknownOperation = 106,
    CompileSchemaFailure = 107,
    BudgetExhausted = 108,
    NotRecoverable = 109
}

public static class RecoveryClassificationRules
{
    public static bool IsRecoverable(RecoveryClassification classification) =>
        classification is RecoveryClassification.StateNotStable
            or RecoveryClassification.ExpectedChildWindowNotActive
            or RecoveryClassification.ApprovedAlternateLocator
            or RecoveryClassification.ApprovedEquivalentTransition
            or RecoveryClassification.ExplicitlyApprovedCandidateLocator;

    public static string DisplayName(RecoveryClassification classification) =>
        classification switch
        {
            RecoveryClassification.StateNotStable => "state-not-stable",
            RecoveryClassification.ExpectedChildWindowNotActive => "expected-child-window-not-active",
            RecoveryClassification.ApprovedAlternateLocator => "approved-alternate-locator",
            RecoveryClassification.ApprovedEquivalentTransition => "approved-equivalent-transition",
            RecoveryClassification.ExplicitlyApprovedCandidateLocator => "explicitly-approved-candidate-locator",
            RecoveryClassification.AssertionFailure => "assertion-failure",
            RecoveryClassification.AuthenticationFailure => "authentication-failure",
            RecoveryClassification.DestructiveActionDenial => "destructive-action-denial",
            RecoveryClassification.AmbiguousLocator => "ambiguous-locator",
            RecoveryClassification.ApplicationBuildMismatch => "application-build-mismatch",
            RecoveryClassification.UnexpectedExternalWindow => "unexpected-external-window",
            RecoveryClassification.UnknownOperation => "unknown-operation",
            RecoveryClassification.CompileSchemaFailure => "compile-schema-failure",
            RecoveryClassification.BudgetExhausted => "budget-exhausted",
            _ => "not-recoverable"
        };
}
