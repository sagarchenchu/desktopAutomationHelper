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

    public ObjectRepositoryOptions ObjectRepository { get; set; } = new();

    public KnowledgeOptions Knowledge { get; set; } = new();

    public JiraOptions Jira { get; set; } = new();

    public PlanningOptions Planning { get; set; } = new();

    public RecoveryOptions Recovery { get; set; } = new();
}
