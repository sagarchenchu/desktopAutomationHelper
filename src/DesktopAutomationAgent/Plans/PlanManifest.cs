using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanManifest
{
    public int SchemaVersion { get; set; }

    public int CatalogSchemaVersion { get; set; }

    public string PlanId { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Required plan property. Null means the JSON omitted <c>steps</c>
    /// (or set it to null).
    /// </summary>
    public List<PlanStep>? Steps { get; set; }

    /// <summary>
    /// Optional failure-handling steps. Null means the JSON omitted
    /// <c>onFailureSteps</c> (or set it to null).
    /// </summary>
    public List<PlanStep>? OnFailureSteps { get; set; }

    /// <summary>
    /// Optional cleanup steps. Null means the JSON omitted
    /// <c>cleanupSteps</c> (or set it to null).
    /// </summary>
    public List<PlanStep>? CleanupSteps { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
