using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Compilation;

public sealed class StepRegistryDocument
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    public int SchemaVersion { get; set; }

    public string RegistryId { get; set; } = string.Empty;

    public string? Name { get; set; }

    public string Version { get; set; } = string.Empty;

    public List<StepMapping> Mappings { get; set; } = [];

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class StepMapping
{
    public string MappingId { get; set; } = string.Empty;

    public StepPattern Pattern { get; set; } = new();

    public List<string> Keywords { get; set; } = [];

    public List<StepParameter> Parameters { get; set; } = [];

    public List<PlanStepTemplate> PlanSteps { get; set; } = [];

    public List<string> ApprovedObjectRefs { get; set; } = [];

    public List<ApprovedKnowledgeRef> ApprovedKnowledgeRefs { get; set; } = [];

    public string? ExpectedState { get; set; }

    public string? Postcondition { get; set; }

    public string Risk { get; set; } = "low";

    public string Source { get; set; } = string.Empty;

    public string Version { get; set; } = string.Empty;

    public StepMappingApproval Approval { get; set; } = new();

    [JsonIgnore]
    public Regex? CompiledRegex { get; set; }
}

public sealed class StepPattern
{
    public string Kind { get; set; } = "exact";

    public string Value { get; set; } = string.Empty;
}

public sealed class StepParameter
{
    public string Name { get; set; } = string.Empty;

    public string Type { get; set; } = "string";
}

public sealed class PlanStepTemplate
{
    public string IdTemplate { get; set; } = string.Empty;

    public string Operation { get; set; } = string.Empty;

    public Dictionary<string, JsonElement>? Arguments { get; set; }

    public List<Dictionary<string, JsonElement>>? Assertions { get; set; }

    public bool? Sensitive { get; set; }

    public bool? CaptureResponse { get; set; }
}

public sealed class ApprovedKnowledgeRef
{
    public string Kind { get; set; } = string.Empty;

    public string Id { get; set; } = string.Empty;

    public string Role { get; set; } = "navigation";
}

public sealed class StepMappingApproval
{
    public string ApprovedBy { get; set; } = string.Empty;

    public DateTimeOffset ApprovedAt { get; set; }

    public string? Notes { get; set; }
}
