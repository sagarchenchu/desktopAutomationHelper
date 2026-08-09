namespace DesktopAutomationAgent.ObjectRepository;

public sealed class PlanObjectReferenceExpansionResult
{
    public bool Success => Errors.Count == 0;

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public string? RepositoryPath { get; init; }

    public string? RepositoryId { get; init; }

    public string? RepositorySha256 { get; init; }

    public IReadOnlyList<string> ResolvedObjectReferences { get; init; } = Array.Empty<string>();
}
