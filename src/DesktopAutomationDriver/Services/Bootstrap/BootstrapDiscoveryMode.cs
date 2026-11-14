namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// How the driver exposes connection bootstrap material (port + bearer token).
/// </summary>
public enum BootstrapDiscoveryMode
{
    /// <summary>
    /// Default for shared-host / production. Token is delivered only via a
    /// per-user SID ACL-protected named pipe and connection descriptor.
    /// Unauthenticated HTTP <c>/verify</c> never returns the bearer token.
    /// </summary>
    Secure = 0,

    /// <summary>
    /// Legacy compatibility: unauthenticated <c>GET /verify</c> returns the
    /// bearer token. Unsafe on shared Citrix/RDS/VDI hosts — opt-in only.
    /// </summary>
    LegacyVerify = 1
}
