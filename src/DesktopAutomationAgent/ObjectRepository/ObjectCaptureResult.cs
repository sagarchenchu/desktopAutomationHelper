namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectCaptureResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? Error { get; init; }

    public string? RepositoryPath { get; init; }

    public string? RepositoryId { get; init; }

    public string? RepositorySha256 { get; init; }

    public string? CaptureId { get; init; }

    public string? PageId { get; init; }

    public string? CaptureFilePath { get; init; }

    public string? CandidateFilePath { get; init; }

    public string? CaptureSha256 { get; init; }

    public string? CandidateSha256 { get; init; }

    public int NodeCount { get; init; }

    public int ElementCount { get; init; }

    public int UnresolvedCount { get; init; }
}

public sealed class ObjectCaptureOptions
{
    public string View { get; init; } = "control";

    public string Root { get; init; } = "activeWindow";

    public int? MaxDepth { get; init; }

    public int? MaxChildren { get; init; }

    public bool? IncludeOffscreen { get; init; }
}
