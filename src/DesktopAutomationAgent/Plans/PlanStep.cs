using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanStep
{
    public string Id { get; set; } = string.Empty;

    public string Operation { get; set; } = string.Empty;

    public Dictionary<string, JsonElement>? Arguments { get; set; }

    public List<PlanAssertion>? Assertions { get; set; }

    public bool Sensitive { get; set; }

    public bool CaptureResponse { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
