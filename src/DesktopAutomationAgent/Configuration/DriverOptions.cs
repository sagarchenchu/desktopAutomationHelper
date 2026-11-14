namespace DesktopAutomationAgent.Configuration;

public sealed class DriverOptions
{
    public const string SectionName = "Driver";

    /// <summary>Explicit driver base URL (for example http://127.0.0.1:33201).</summary>
    public string? BaseUrl { get; set; }

    /// <summary>Bearer token for authenticated driver endpoints. Never log this value.</summary>
    public string? BearerToken { get; set; }

    /// <summary>
    /// Discovery mode when BaseUrl/BearerToken are omitted.
    /// <c>Secure</c> (default) uses the SID ACL-protected named pipe / descriptor.
    /// <c>LegacyVerify</c> uses unauthenticated HTTP <c>/verify</c> (unsafe on shared hosts).
    /// </summary>
    public string DiscoveryMode { get; set; } = "Secure";

    /// <summary>
    /// Required when <see cref="DiscoveryMode"/> is <c>LegacyVerify</c>.
    /// Documents acceptance of the shared-host token-leak risk of HTTP <c>/verify</c>.
    /// </summary>
    public bool AcknowledgeLegacyVerifyRisk { get; set; }

    /// <summary>
    /// Legacy verify discovery endpoint. Used only when
    /// <see cref="DiscoveryMode"/> is <c>LegacyVerify</c>.
    /// </summary>
    public string VerifyUrl { get; set; } = "http://localhost:9102/verify";

    public int RequestTimeoutSeconds { get; set; } = 20;

    public int ExpectedCatalogSchemaVersion { get; set; } = 2;

    /// <summary>
    /// When false (default), only loopback driver URLs are accepted.
    /// </summary>
    public bool AllowRemoteDriver { get; set; }

    public bool UsesSecureDiscovery =>
        string.Equals(DiscoveryMode, "Secure", StringComparison.OrdinalIgnoreCase)
        || string.IsNullOrWhiteSpace(DiscoveryMode);

    public bool UsesLegacyVerifyDiscovery =>
        string.Equals(DiscoveryMode, "LegacyVerify", StringComparison.OrdinalIgnoreCase);
}
