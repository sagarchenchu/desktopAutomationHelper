using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanAssertion
{
    public string Path { get; set; } = string.Empty;

    public string Operator { get; set; } = string.Empty;

    public JsonElement? Expected { get; set; }

    public bool IgnoreCase { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
