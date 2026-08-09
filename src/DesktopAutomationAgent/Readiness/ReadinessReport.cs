namespace DesktopAutomationAgent.Readiness;

public sealed class ReadinessReport
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string GeneratedAtUtc { get; init; } = DateTimeOffset.UtcNow.ToString("o");

    public string? Username { get; init; }

    public string? DriverBaseUrl { get; init; }

    public string? DiscoveryMethod { get; init; }

    public string? DriverVersion { get; init; }

    public int? CatalogSchemaVersion { get; init; }

    public int? OperationCount { get; init; }

    public string WorkspaceRoot { get; init; } = string.Empty;

    public IReadOnlyList<ReadinessCheck> Checks { get; init; } = Array.Empty<ReadinessCheck>();

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();
}

public sealed class ReadinessCheck
{
    public required string Name { get; init; }

    public required string Status { get; init; }

    public string? Detail { get; init; }
}

public static class ReadinessCheckStatus
{
    public const string Passed = "passed";
    public const string Failed = "failed";
    public const string Skipped = "skipped";
}
