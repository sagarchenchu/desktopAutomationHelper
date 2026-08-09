namespace DesktopAutomationAgent.Cli;

public static class ExitCodes
{
    public const int Success = 0;
    public const int UsageOrConfiguration = 2;
    public const int DriverUnavailable = 3;
    public const int AuthOrCatalog = 4;
    public const int SuiteOrWorkspace = 5;
    public const int ExecutionFailure = 6;
    public const int Cancelled = 7;
}
