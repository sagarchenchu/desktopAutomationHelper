namespace DesktopAutomationAgent.Driver;

public enum UiFailureClassification
{
    DriverUnavailable,
    Authentication,
    Catalog,
    OperationFailure,
    AssertionFailure,
    ExecutionTimeout,
    Cancelled
}
