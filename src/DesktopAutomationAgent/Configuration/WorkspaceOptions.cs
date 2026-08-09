namespace DesktopAutomationAgent.Configuration;

public sealed class WorkspaceOptions
{
    public const string SectionName = "Workspace";

    /// <summary>Workspace root relative to the process working directory, or an absolute path.</summary>
    public string Root { get; set; } = "automation";
}
