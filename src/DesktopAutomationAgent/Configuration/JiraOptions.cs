namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Jira read-only test-source configuration (Phase 4E).
/// Secrets must come from env (<c>DA_AGENT__Jira__*</c>), OS secret store, or
/// gitignored local config — never committed.
/// </summary>
public sealed class JiraOptions
{
    /// <summary>Absolute Jira base URL (https required unless local test server is explicitly allowed).</summary>
    public string? BaseUrl { get; set; }

    /// <summary>Auth mode: basic (email+apiToken), bearer, or pat.</summary>
    public string AuthMode { get; set; } = "basic";

    public string? Email { get; set; }

    /// <summary>Atlassian API token (basic auth password). Never log.</summary>
    public string? ApiToken { get; set; }

    /// <summary>OAuth / bearer access token. Never log.</summary>
    public string? BearerToken { get; set; }

    /// <summary>Personal access token (Authorization: Bearer). Never log.</summary>
    public string? PersonalAccessToken { get; set; }

    /// <summary>
    /// Optional Windows Credential Manager target name. When set, the stored
    /// password/secret is used for the active auth mode (apiToken / bearer / pat).
    /// </summary>
    public string? CredentialTarget { get; set; }

    /// <summary>
    /// Mapping profile name: <c>plain-fields</c> (default), <c>xray</c>, <c>zephyr</c>,
    /// or <c>normalized</c>. Profiles only supply default field paths; they do not
    /// hardcode a vendor HTTP model.
    /// </summary>
    public string MappingProfile { get; set; } = "plain-fields";

    public JiraFieldMappingOptions FieldMapping { get; set; } = new();

    public int RequestTimeoutSeconds { get; set; } = 30;

    public int MaxResponseBytes { get; set; } = 2_097_152;

    /// <summary>Bounded retries for idempotent GET only (0–3).</summary>
    public int MaxRetries { get; set; } = 2;

    public int PageSize { get; set; } = 50;

    public int MaxPages { get; set; } = 20;

    /// <summary>
    /// When true, http://127.0.0.1 and http://localhost are allowed (unit tests /
    /// explicit local mock servers only). Remote http is always rejected.
    /// </summary>
    public bool AllowLocalHttpTestServer { get; set; }

    /// <summary>Comma-separated Jira <c>fields</c> query values. Empty uses mapping-derived fields.</summary>
    public string? IssueFields { get; set; }
}

/// <summary>Configurable field/plugin mapping into the Phase 4D normalized model.</summary>
public sealed class JiraFieldMappingOptions
{
    public string SummaryPath { get; set; } = "fields.summary";

    public string? DescriptionPath { get; set; } = "fields.description";

    /// <summary>Path to Gherkin/feature text (custom field or Xray cucumber field).</summary>
    public string? FeaturePath { get; set; }

    /// <summary>Path to a JSON string/object already shaped as Phase 4D test-case.</summary>
    public string? NormalizedDocumentPath { get; set; }

    public string? RevisionPath { get; set; }

    public string? UpdatedAtPath { get; set; } = "fields.updated";

    public string? TagsPath { get; set; }

    public string KeyPath { get; set; } = "key";

    /// <summary>Optional ADF → plain text extraction when description is Atlassian Document Format.</summary>
    public bool ExtractAdfText { get; set; } = true;
}
