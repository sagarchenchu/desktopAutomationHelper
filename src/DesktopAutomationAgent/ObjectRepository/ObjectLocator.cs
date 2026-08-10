using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectLocator
{
    public string? AutomationId { get; set; }

    public string? Name { get; set; }

    public string? ClassName { get; set; }

    public string? ControlType { get; set; }

    public string? MatchMode { get; set; }

    public int? FoundIndex { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
