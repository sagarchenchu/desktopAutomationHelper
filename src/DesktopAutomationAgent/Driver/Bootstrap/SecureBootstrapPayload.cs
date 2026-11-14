using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Driver.Bootstrap;

/// <summary>Payload read from the secure bootstrap channel. Never log Token.</summary>
public sealed class SecureBootstrapPayload
{
    [JsonPropertyName("schemaVersion")]
    public int SchemaVersion { get; set; }

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

    [JsonPropertyName("startedAtUtc")]
    public DateTimeOffset StartedAtUtc { get; set; }

    [JsonPropertyName("pipeName")]
    public string PipeName { get; set; } = string.Empty;
}
