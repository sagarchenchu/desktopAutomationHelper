namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectRepositoryValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public ObjectRepositorySnapshot? Snapshot { get; init; }

    public string RepositoryPath { get; init; } = string.Empty;

    public string? ManifestSha256 { get; init; }

    public IReadOnlyDictionary<string, string>? FileHashes { get; init; }

    public string? AggregateSha256 { get; init; }
}
