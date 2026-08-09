using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Driver.Models;

/// <summary>WebDriver-style envelope used by /verify and /status.</summary>
public sealed class WebDriverEnvelope<T>
{
    public string? SessionId { get; set; }

    public int Status { get; set; }

    public T? Value { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

/// <summary>UiResponse envelope used by GET /ui/operations.</summary>
public sealed class UiEnvelope<T>
{
    public bool Success { get; set; }

    public T? Value { get; set; }

    public string? Error { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class VerifyValueDto
{
    public bool Running { get; set; }

    public string Username { get; set; } = string.Empty;

    public int Port { get; set; }

    public int? ProbePort { get; set; }

    public string Token { get; set; } = string.Empty;

    public string AuthorizationHeader { get; set; } = string.Empty;

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class StatusValueDto
{
    public bool Ready { get; set; }

    public string Message { get; set; } = string.Empty;

    public BuildInfoDto? Build { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class BuildInfoDto
{
    public string Version { get; set; } = string.Empty;

    public string Revision { get; set; } = string.Empty;

    public string Time { get; set; } = string.Empty;

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class OperationsCatalogDto
{
    public int SchemaVersion { get; set; }

    public string DriverVersion { get; set; } = string.Empty;

    public List<OperationDescriptorDto> Operations { get; set; } = [];

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class OperationDescriptorDto
{
    public string Name { get; set; } = string.Empty;

    public List<string> Aliases { get; set; } = [];

    public List<string> DeprecatedAliases { get; set; } = [];

    public string Category { get; set; } = string.Empty;

    public string OperationType { get; set; } = string.Empty;

    public bool RequiresSession { get; set; }

    public List<string> RequiredInputs { get; set; } = [];

    public List<List<string>> RequiredInputAlternatives { get; set; } = [];

    public bool Deprecated { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
