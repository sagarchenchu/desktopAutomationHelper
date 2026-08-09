using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectSource
{
    public string Kind { get; set; } = string.Empty;

    public string? Path { get; set; }

    public Dictionary<string, JsonElement>? Metadata { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
