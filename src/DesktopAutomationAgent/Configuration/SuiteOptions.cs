namespace DesktopAutomationAgent.Configuration;

public sealed class SuiteOptions
{
    public const string SectionName = "Suites";

    /// <summary>
    /// Optional project-specific additional restriction applied after
    /// <see cref="JiraKeyContract.CanonicalPattern"/>. Defaults to the canonical
    /// pattern. Configured patterns cannot broaden acceptance beyond the canonical
    /// contract because canonical validation always runs first.
    /// </summary>
    public string JiraKeyPattern { get; set; } = JiraKeyContract.CanonicalPattern;
}
