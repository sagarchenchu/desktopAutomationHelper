namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Top-level agent configuration bound from appsettings, local overrides,
/// environment variables (<c>DA_AGENT__*</c>), and command-line arguments.
/// </summary>
public sealed class AgentOptions
{
    public DriverOptions Driver { get; set; } = new();

    public WorkspaceOptions Workspace { get; set; } = new();

    public SuiteOptions Suites { get; set; } = new();

    public RunnerOptions Runner { get; set; } = new();
}
