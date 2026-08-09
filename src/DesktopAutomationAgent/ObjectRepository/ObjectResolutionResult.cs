namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectResolutionResult
{
    public bool IsResolved => Errors.Count == 0 && Locator is not null;

    public string Reference { get; init; } = string.Empty;

    public string? PageId { get; init; }

    public string? ElementId { get; init; }

    public string? PageName { get; init; }

    public string? ElementDescription { get; init; }

    public ObjectLocator? Locator { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}
