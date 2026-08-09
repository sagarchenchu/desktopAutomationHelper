using System.Text.Json;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Execution;

public sealed class RunReport
{
    public required string RunId { get; init; }

    public required string Status { get; init; }

    public required int ExitCode { get; init; }

    public required string PlanPath { get; init; }

    public string? PlanId { get; init; }

    public string? PlanName { get; init; }

    public string? PlanSha256 { get; init; }

    public bool DryRun { get; init; }

    public string? DriverBaseUrl { get; init; }

    public int? CatalogSchemaVersion { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public DateTimeOffset? FinishedAtUtc { get; init; }

    public IReadOnlyList<StepRunResult> Steps { get; init; } = Array.Empty<StepRunResult>();

    public IReadOnlyList<StepRunResult> OnFailureSteps { get; init; } = Array.Empty<StepRunResult>();

    public RunFailure? Failure { get; init; }
}

public sealed class StepRunResult
{
    public required string Id { get; init; }

    public required string Operation { get; init; }

    public required string Phase { get; init; }

    public required bool Success { get; init; }

    public bool Sensitive { get; init; }

    public bool CaptureResponse { get; init; }

    public bool Skipped { get; init; }

    public string? SkipReason { get; init; }

    public Dictionary<string, JsonElement>? Arguments { get; init; }

    public JsonElement? ResponseValue { get; init; }

    public string? Error { get; init; }

    public string? ScreenshotPath { get; init; }

    public IReadOnlyList<AssertionRunResult> Assertions { get; init; } = Array.Empty<AssertionRunResult>();

    public TimeSpan Duration { get; init; }
}

public sealed class AssertionRunResult
{
    public required string Path { get; init; }

    public required string Operator { get; init; }

    public required bool Passed { get; init; }

    public string? Message { get; init; }

    public JsonElement? Expected { get; init; }

    public JsonElement? Actual { get; init; }
}

public sealed class RunFailure
{
    public required UiFailureClassification Classification { get; init; }

    public required string Message { get; init; }

    public string? StepId { get; init; }

    public string? ScreenshotPath { get; init; }
}
