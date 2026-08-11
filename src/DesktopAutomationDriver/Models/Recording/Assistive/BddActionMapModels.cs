using System.Text.Json.Serialization;

namespace DesktopAutomationDriver.Models.Recording.Assistive;

public sealed class BddActionMapDocument
{
    public int SchemaVersion { get; set; } = 1;

    public string RecordingId { get; set; } = string.Empty;

    public string JiraKey { get; set; } = string.Empty;

    public string SourceRecording { get; set; } = string.Empty;

    public DateTimeOffset CreatedAt { get; set; }

    public List<BddActionMapPageRef> Pages { get; set; } = [];

    public List<BddActionMapGroup> BddGroups { get; set; } = [];

    public List<string> UnmappedEventIds { get; set; } = [];

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
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

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ObjectRef { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TargetObjectRef { get; set; }

    public string Operation { get; set; } = string.Empty;
}
