namespace DesktopAutomationAgent.Suites;

public sealed class SuiteValidationResult
{
    public required string FilePath { get; init; }

    public required string SuiteName { get; init; }

    public bool SuiteEnabled { get; init; }

    public int TotalCount { get; init; }

    public int EnabledCount { get; init; }

    public int DisabledCount { get; init; }

    public int DuplicateCount { get; init; }

    public IReadOnlyList<string> EnabledJiraKeys { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public bool IsValid => Errors.Count == 0;
}

public sealed class KeyValidationResult
{
    public IReadOnlyList<string> ValidKeys { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public bool IsValid => Errors.Count == 0;
}
