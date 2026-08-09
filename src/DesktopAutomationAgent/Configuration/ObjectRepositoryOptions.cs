namespace DesktopAutomationAgent.Configuration;

public sealed class ObjectRepositoryOptions
{
    public const string SectionName = "ObjectRepository";

    public int MaxFileBytes { get; set; } = 5_242_880;

    public int MaxPages { get; set; } = 500;

    public int MaxElementsPerPage { get; set; } = 5000;

    public int MaxTotalElements { get; set; } = 50_000;

    public int DiagnosticTimeoutMilliseconds { get; set; } = 15_000;
}
