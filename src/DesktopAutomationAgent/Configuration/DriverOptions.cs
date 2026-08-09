namespace DesktopAutomationAgent.Configuration;

public sealed class DriverOptions
{
    public const string SectionName = "Driver";

    /// <summary>Explicit driver base URL (for example http://127.0.0.1:33201).</summary>
    public string? BaseUrl { get; set; }

    /// <summary>Bearer token for authenticated driver endpoints. Never log this value.</summary>
    public string? BearerToken { get; set; }

    /// <summary>Verify discovery endpoint used when BaseUrl/BearerToken are omitted.</summary>
    public string VerifyUrl { get; set; } = "http://localhost:9102/verify";

    public int RequestTimeoutSeconds { get; set; } = 20;

    public int ExpectedCatalogSchemaVersion { get; set; } = 2;

    /// <summary>
    /// When false (default), only loopback driver URLs are accepted.
    /// </summary>
    public bool AllowRemoteDriver { get; set; }
}
