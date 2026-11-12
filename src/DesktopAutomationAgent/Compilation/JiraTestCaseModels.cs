using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Compilation;

public sealed class JiraTestCaseDocument
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    public int SchemaVersion { get; set; }

    public string JiraKey { get; set; } = string.Empty;

    public string? Revision { get; set; }

    public DateTimeOffset? UpdatedAt { get; set; }

    public string Summary { get; set; } = string.Empty;

    public string? Description { get; set; }

    public string? Feature { get; set; }

    public StructuredBackground? Background { get; set; }

    public List<StructuredScenario>? Scenarios { get; set; }

    public List<string>? Tags { get; set; }

    public List<string>? TestDataRefs { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }

    /// <summary>Stable revision token used in fingerprints (revision preferred).</summary>
    public string ResolveRevisionToken()
    {
        if (!string.IsNullOrWhiteSpace(Revision))
            return Revision.Trim();
        if (UpdatedAt is { } updated)
            return updated.UtcDateTime.ToString("O");
        return string.Empty;
    }
}

public sealed class StructuredBackground
{
    public string? Name { get; set; }

    public List<StructuredStep> Steps { get; set; } = [];
}

public sealed class StructuredScenario
{
    public string Name { get; set; } = string.Empty;

    public List<string>? Tags { get; set; }

    public List<StructuredStep> Steps { get; set; } = [];
}

public sealed class StructuredStep
{
    public string Keyword { get; set; } = string.Empty;

    public string Text { get; set; } = string.Empty;
}
