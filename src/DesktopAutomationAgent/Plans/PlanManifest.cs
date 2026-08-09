using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanManifest
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    public int SchemaVersion { get; set; }

    public int CatalogSchemaVersion { get; set; }

    public string PlanId { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string? Description { get; set; }

    public List<string>? Tags { get; set; }

    public Dictionary<string, JsonElement>? Metadata { get; set; }

    /// <summary>
    /// Required. Null means the JSON omitted <c>steps</c>.
    /// </summary>
    public List<PlanStep>? Steps { get; set; }

    /// <summary>
    /// Optional failure cleanup steps. Null means omitted.
    /// </summary>
    public List<PlanStep>? OnFailureSteps { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
