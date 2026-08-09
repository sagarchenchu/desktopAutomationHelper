namespace DesktopAutomationAgent.Configuration;

public sealed class RunnerOptions
{
    public const string SectionName = "Runner";

    public int StepTransportTimeoutSeconds { get; set; } = 60;

    public int CleanupTimeoutSeconds { get; set; } = 15;

    public int MaxPlanBytes { get; set; } = 1_048_576;

    public int MaxResponseBytes { get; set; } = 10_485_760;

    public int RegexTimeoutMilliseconds { get; set; } = 500;
}
