namespace DesktopAutomationAgent.ObjectRepository;

public enum ObjectVerificationStatus
{
    Passed,
    Missing,
    Ambiguous,
    Fragile,
    Failed
}

public sealed class ObjectVerificationItemResult
{
    public required string Reference { get; init; }

    public required ObjectVerificationStatus Status { get; init; }

    public int MatchCount { get; init; }

    public string? Error { get; init; }

    public string? Warning { get; init; }
}

public sealed class ObjectVerificationResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? Error { get; init; }

    public string? RepositoryPath { get; init; }

    public string? RepositoryId { get; init; }

    public string? RepositorySha256 { get; init; }

    public int Total { get; init; }

    public int Passed { get; init; }

    public int Missing { get; init; }

    public int Ambiguous { get; init; }

    public int Fragile { get; init; }

    public int Failed { get; init; }

    public double DurationMilliseconds { get; init; }

    public IReadOnlyList<ObjectVerificationItemResult> Items { get; init; } = Array.Empty<ObjectVerificationItemResult>();
}

public sealed class ObjectVerificationOptions
{
    public string? PageId { get; init; }

    public string? ObjectRef { get; init; }

    public string View { get; init; } = "control";

    public string Root { get; init; } = "activeWindow";

    public int? MaxDepth { get; init; }

    public bool? IncludeOffscreen { get; init; }
}
