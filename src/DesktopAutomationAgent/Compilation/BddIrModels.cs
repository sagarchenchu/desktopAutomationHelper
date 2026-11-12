namespace DesktopAutomationAgent.Compilation;

/// <summary>Immutable intermediate representation after BDD parse + And/But normalization.</summary>
public sealed class BddIrDocument
{
    public required string FeatureName { get; init; }

    public IReadOnlyList<BddIrScenario> Scenarios { get; init; } = Array.Empty<BddIrScenario>();

    public IReadOnlyList<string> Tags { get; init; } = Array.Empty<string>();
}

public sealed class BddIrScenario
{
    public required string Name { get; init; }

    public IReadOnlyList<string> Tags { get; init; } = Array.Empty<string>();

    public IReadOnlyList<BddIrStep> Steps { get; init; } = Array.Empty<BddIrStep>();
}

public sealed class BddIrStep
{
    /// <summary>Original keyword including And/But.</summary>
    public required string RawKeyword { get; init; }

    /// <summary>Normalized primary keyword: Given / When / Then.</summary>
    public required string Keyword { get; init; }

    public required string Text { get; init; }

    public int Line { get; init; }

    public int Ordinal { get; init; }
}
