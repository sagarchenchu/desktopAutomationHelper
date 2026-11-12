using DesktopAutomationAgent.Compilation;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Provider-neutral read-only source of normalized Phase 4D test cases.
/// Vendor/plugin differences (plain fields, Xray, Zephyr, …) stay behind adapters.
/// </summary>
public interface IJiraTestCaseSource
{
    /// <summary>
    /// Fetches a single test case by canonical Jira key. Implementations must
    /// reject non-canonical keys before any network I/O.
    /// </summary>
    Task<JiraTestCaseFetchResult> FetchAsync(string jiraKey, CancellationToken cancellationToken = default);

    /// <summary>
    /// Optional paginated search of issue keys (for suite/filter expansion).
    /// Default implementations may throw <see cref="NotSupportedException"/>.
    /// </summary>
    Task<JiraKeySearchResult> SearchKeysAsync(
        string jql,
        CancellationToken cancellationToken = default);
}

public sealed class JiraTestCaseFetchResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? JiraKey { get; init; }

    public JiraTestCaseDocument? Document { get; init; }

    public JiraSourceMetadata? Metadata { get; init; }

    public bool NotModified { get; init; }

    public bool FromCache { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = [];

    public IReadOnlyList<string> Warnings { get; init; } = [];

    public JiraRateLimitInfo? RateLimit { get; init; }
}

public sealed class JiraKeySearchResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public IReadOnlyList<string> Keys { get; init; } = [];

    public int PageCount { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = [];

    public JiraRateLimitInfo? RateLimit { get; init; }
}

public sealed record JiraSourceMetadata
{
    public string JiraKey { get; init; } = string.Empty;

    public string MappingProfile { get; init; } = string.Empty;

    public string RevisionToken { get; init; } = string.Empty;

    public string? ETag { get; init; }

    public string? LastModified { get; init; }

    public string? UpdatedAt { get; init; }

    public string ContentSha256 { get; init; } = string.Empty;

    public string FetchedAtUtc { get; init; } = string.Empty;

    public string SourceKind { get; init; } = "jira-rest";

    public bool Redacted { get; init; } = true;

    /// <summary>Fields intentionally omitted from cache (never dump unrestricted issue JSON).</summary>
    public IReadOnlyList<string> RedactedFieldPaths { get; init; } = [];
}

public sealed class JiraRateLimitInfo
{
    public int? Limit { get; init; }

    public int? Remaining { get; init; }

    public int? RetryAfterSeconds { get; init; }
}
