namespace DesktopAutomationAgent.Plans;

public sealed class PlanValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string>? Warnings { get; init; }

    public string PlanPath { get; init; } = string.Empty;

    public string PlanId { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    public int StepCount { get; init; }

    public int OnFailureStepCount { get; init; }

    public int CleanupStepCount { get; init; }

    public int TotalStepCount { get; init; }

    public string? Sha256 { get; init; }

    public string? ObjectRepositoryPath { get; init; }

    public string? ObjectRepositoryId { get; init; }

    public string? ObjectRepositorySha256 { get; init; }

    public IReadOnlyList<string>? ResolvedObjectReferences { get; init; }

    public PlanManifest? Plan { get; init; }
}
