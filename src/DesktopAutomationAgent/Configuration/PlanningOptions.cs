namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Guarded AI planning configuration. Credentials come only from env/secret stores —
/// never from plan/proposal JSON or UI/Jira text.
/// </summary>
public sealed class PlanningOptions
{
    /// <summary>
    /// Model client kind: <c>fake</c> (CI default) or a future provider-neutral adapter name.
    /// </summary>
    public string Provider { get; set; } = "fake";

    /// <summary>
    /// Optional provider model id (provider-neutral). Ignored by the fake client.
    /// </summary>
    public string? ModelId { get; set; }

    /// <summary>
    /// Default approval mode when the CLI omits <c>--approval-mode</c>.
    /// </summary>
    public string DefaultApprovalMode { get; set; } = "proposal-only";

    /// <summary>Bound model call duration (seconds).</summary>
    public int RequestTimeoutSeconds { get; set; } = 30;

    /// <summary>Soft max prompt tokens (chars/4 heuristic for budgeting).</summary>
    public int MaxPromptTokens { get; set; } = 4000;

    /// <summary>Soft max completion tokens.</summary>
    public int MaxCompletionTokens { get; set; } = 2000;

    /// <summary>Max raw model output bytes before reject.</summary>
    public int MaxOutputBytes { get; set; } = 65_536;

    /// <summary>Max retrieval hits across all corpora.</summary>
    public int MaxRetrievalHits { get; set; } = 40;

    /// <summary>Max characters of untrusted UI/Jira text attached to a request.</summary>
    public int MaxUntrustedContextChars { get; set; } = 4000;

    /// <summary>Optional API key env var name (never logged). Fake provider ignores.</summary>
    public string? ApiKeyEnvironmentVariable { get; set; }

    /// <summary>When true, embeddings may augment lexical retrieval (never source of truth).</summary>
    public bool AllowOptionalEmbeddings { get; set; }
}
