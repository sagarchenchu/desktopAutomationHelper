namespace DesktopAutomationAgent.Driver;

public enum UiFailureClassification
{
    DriverUnavailable,
    Authentication,
    Catalog,
    PlanValidation,
    OperationFailure,
    AssertionFailure,
    ExecutionTimeout,
    Cancelled,
    ArtifactFailure
}
