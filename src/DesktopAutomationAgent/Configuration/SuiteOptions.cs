namespace DesktopAutomationAgent.Configuration;

public sealed class SuiteOptions
{
    public const string SectionName = "Suites";

    public string JiraKeyPattern { get; set; } = "^[A-Z][A-Z0-9_]*-[0-9]+$";
}
