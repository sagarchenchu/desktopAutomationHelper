namespace DesktopAutomationAgent.Configuration;

public sealed class KnowledgeOptions
{
    public int MaxFileBytes { get; set; } = 10_485_760;

    public int MaxExplorationRuns { get; set; } = 500;
}
