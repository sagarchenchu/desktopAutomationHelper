using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectElementDefinition
{
    public string? Description { get; set; }

    public ObjectLocator? Locator { get; set; }

    public ObjectQuality? Quality { get; set; }

    public ObjectSource? Source { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
