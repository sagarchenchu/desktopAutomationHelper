namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Driver bootstrap / discovery security settings.
/// Bound from configuration section <c>Bootstrap</c>.
/// </summary>
public sealed class BootstrapOptions
{
    public const string SectionName = "Bootstrap";

    /// <summary>
    /// <see cref="BootstrapDiscoveryMode.Secure"/> (default) or
    /// <see cref="BootstrapDiscoveryMode.LegacyVerify"/> (explicit opt-in).
    /// Accepts the enum name as a string in configuration.
    /// </summary>
    public BootstrapDiscoveryMode Mode { get; set; } = BootstrapDiscoveryMode.Secure;

    /// <summary>
    /// When true (and only meaningful with <see cref="BootstrapDiscoveryMode.LegacyVerify"/>),
    /// documents that the operator accepts the shared-host token-leak risk of HTTP <c>/verify</c>.
    /// Secure mode ignores this flag and never returns the token over HTTP.
    /// </summary>
    public bool AcknowledgeLegacyVerifyRisk { get; set; }

    public bool IsSecure => Mode == BootstrapDiscoveryMode.Secure;

    public bool AllowsLegacyVerifyHttp =>
        Mode == BootstrapDiscoveryMode.LegacyVerify && AcknowledgeLegacyVerifyRisk;
}
