namespace DesktopAutomationDriver.Models.Response;

/// <summary>
/// Response payload for the unauthenticated <c>GET /verify</c> endpoint.
/// Contains everything a client needs to connect to this driver instance.
/// </summary>
public class VerifyResponse
{
    /// <summary>Always <c>true</c> when this response is returned.</summary>
    public bool Running { get; set; } = true;

    /// <summary>Windows login name of the user running this driver instance.</summary>
    public string Username { get; set; } = string.Empty;

    /// <summary>
    /// Port on which all automation endpoints are listening.
    /// Determined at startup from the Windows username.
    /// </summary>
    public int Port { get; set; }

    /// <summary>
    /// The fixed probe port (9102) if the driver successfully bound it;
    /// <c>null</c> if another driver instance already holds that port.
    /// </summary>
    public int? ProbePort { get; set; }

    /// <summary>
    /// The Bearer token required in the <c>Authorization</c> header for all
    /// automation endpoints (every route except <c>GET /verify</c>).
    /// </summary>
    public string Token { get; set; } = string.Empty;

    /// <summary>
    /// Ready-to-use value for the <c>Authorization</c> HTTP header,
    /// i.e. <c>"Bearer &lt;token&gt;"</c>.
    /// </summary>
    public string AuthorizationHeader { get; set; } = string.Empty;
}
