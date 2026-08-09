using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Execution;

public sealed class ExecutablePlan
{
    public required PlanManifest Manifest { get; init; }

    public required string PlanPath { get; init; }

    public required string Sha256 { get; init; }

    public required string RunDirectory { get; init; }

    public required string RunId { get; init; }
}
