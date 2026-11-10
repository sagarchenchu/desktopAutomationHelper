using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Compilation;

/// <summary>
/// Agent-side DTO matching <c>automation/schemas/bdd-action-map.schema.json</c>.
/// Duplicated intentionally — the agent must not reference the driver project.
/// </summary>
public sealed class BddActionMapDocument
{
    public int SchemaVersion { get; set; }

    public string RecordingId { get; set; } = string.Empty;

    public string JiraKey { get; set; } = string.Empty;

    public string SourceRecording { get; set; } = string.Empty;

    public DateTimeOffset CreatedAt { get; set; }

    public List<BddActionMapPageRef> Pages { get; set; } = [];

    public List<BddActionMapGroup> BddGroups { get; set; } = [];

    public List<string> UnmappedEventIds { get; set; } = [];

    public List<string>? Warnings { get; set; }
}

public sealed class BddActionMapPageRef
{
    public string PageId { get; set; } = string.Empty;

    public string WindowTitle { get; set; } = string.Empty;

    public string File { get; set; } = string.Empty;
}

public sealed class BddActionMapGroup
{
    public string GroupId { get; set; } = string.Empty;

    public string Statement { get; set; } = string.Empty;

    public List<BddActionMapActionRef> Actions { get; set; } = [];
}

public sealed class BddActionMapActionRef
{
    public string EventId { get; set; } = string.Empty;

    public int Sequence { get; set; }

    public string PageId { get; set; } = string.Empty;

    public string? ObjectRef { get; set; }

    public string? TargetObjectRef { get; set; }

    public string Operation { get; set; } = string.Empty;
}

/// <summary>Minimal recording export shape needed to join values/locators by eventId.</summary>
public sealed class RecordingDocument
{
    public string? RecordingId { get; set; }

    public List<RecordingAction> Actions { get; set; } = [];
}

public sealed class RecordingAction
{
    public string? EventId { get; set; }

    public int? Sequence { get; set; }

    public string? ActionType { get; set; }

    public string? Operation { get; set; }

    public string? Value { get; set; }

    public string? Description { get; set; }

    public string? ObjectRef { get; set; }

    public string? TargetObjectRef { get; set; }

    public RecordingElement? Element { get; set; }

    public RecordingElement? TargetElement { get; set; }

    public List<RecordingElement>? MenuPath { get; set; }

    public Dictionary<string, string>? Metadata { get; set; }
}

public sealed class RecordingElement
{
    public string? Name { get; set; }

    public string? AutomationId { get; set; }

    public string? ClassName { get; set; }

    public string? ControlType { get; set; }

    public string? SuggestedXPath { get; set; }
}
