namespace DesktopAutomationDriver.Services;

/// <summary>
/// Holds per-user driver context: the deterministic user-specific port,
/// the one-time Bearer token generated at startup, and probe-port status.
/// </summary>
public interface IDriverContext
{
    /// <summary>Windows login name of the user who launched the driver.</summary>
    string Username { get; }

    /// <summary>
    /// Deterministic port derived from <see cref="Username"/>.
    /// Each Windows account always maps to the same port (range 30000–39999)
    /// so multiple users on a shared machine do not collide.
    /// </summary>
    int MainPort { get; }

    /// <summary>
    /// Fixed well-known probe port (9102).
    /// The driver attempts to bind this port for the /verify endpoint on
    /// startup; if the port is already held by another user's driver the
    /// attempt is silently skipped.
    /// </summary>
    int ProbePort { get; }

    /// <summary>
    /// Whether the driver successfully bound the probe port (9102).
    /// Set by <c>Program.cs</c> after the probe-server startup attempt.
    /// </summary>
    bool ProbePortActive { get; set; }

    /// <summary>
    /// Randomly generated opaque token that must be sent as
    /// <c>Authorization: Bearer &lt;token&gt;</c> with every request
    /// (except <c>GET /verify</c>).
    /// </summary>
    string BearerToken { get; }
}
