using System.Text.Json.Serialization;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Connection material delivered over the SID ACL-protected bootstrap channel.
/// Never log <see cref="Token"/> or <see cref="AuthorizationHeader"/>.
/// </summary>
public sealed class SecureBootstrapPayload
{
    [JsonPropertyName("schemaVersion")]
    public int SchemaVersion { get; set; } = SecureBootstrapConstants.SchemaVersion;

    [JsonPropertyName("userSid")]
    public string UserSid { get; set; } = string.Empty;

    [JsonPropertyName("username")]
    public string Username { get; set; } = string.Empty;

    [JsonPropertyName("port")]
    public int Port { get; set; }

    [JsonPropertyName("token")]
    public string Token { get; set; } = string.Empty;

    [JsonPropertyName("authorizationHeader")]
    public string AuthorizationHeader { get; set; } = string.Empty;

    [JsonPropertyName("pid")]
    public int Pid { get; set; }

    /// <summary>
    /// UTC start time of the owning driver process. Used with <see cref="Pid"/>
    /// to detect PID reuse after crashes.
    /// </summary>
    [JsonPropertyName("processStartTimeUtc")]
    public DateTimeOffset ProcessStartTimeUtc { get; set; }

    [JsonPropertyName("startedAtUtc")]
    public DateTimeOffset StartedAtUtc { get; set; }

    [JsonPropertyName("pipeName")]
    public string PipeName { get; set; } = string.Empty;
}
