using System.Text.Json;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Recording.Assistive;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Pure Assistive artifact builder (no UIA, WinForms, HTTP, or filesystem).
/// </summary>
public sealed class AssistiveArtifactBuilder
{
    public sealed class PageCandidateDocument
    {
        public int SchemaVersion { get; set; } = 1;
        public string PageId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string State { get; set; } = "candidate";
        public Dictionary<string, PageCandidateElement> Elements { get; set; } = new(StringComparer.Ordinal);
        public List<Dictionary<string, object?>>? Unresolved { get; set; }
    }

    public sealed class PageCandidateElement
    {
        public string? Description { get; set; }
        public Dictionary<string, object?> Locator { get; set; } = new(StringComparer.Ordinal);
        public PageCandidateQuality Quality { get; set; } = new();
        public PageCandidateSource Source { get; set; } = new();
    }

    public sealed class PageCandidateQuality
    {
        public string Grade { get; set; } = "medium";
        public List<string> Warnings { get; set; } = [];
    }

    public sealed class PageCandidateSource
    {
        public string Kind { get; set; } = "capture";
        public string? Path { get; set; }
        public Dictionary<string, object?> Metadata { get; set; } = new(StringComparer.Ordinal);
    }

    public sealed class BuildResult
    {
        public List<PageCandidateDocument> Pages { get; init; } = [];
        public BddActionMapDocument? BddActionMap { get; init; }
        public List<string> Warnings { get; init; } = [];
    }

    public BuildResult Build(
        string recordingId,
        string sourceRecordingFileName,
        DateTimeOffset createdAt,
        string? jiraKey,
        IReadOnlyList<RecordedAction> actions)
    {
        var warnings = new List<string>();
        var assistive = actions
            .Where(a => a.Mode == RecordingMode.Assistive)
            .OrderBy(a => a.Sequence ?? int.MaxValue)
            .ThenBy(a => a.Timestamp)
            .ToList();

        if (assistive.Count == 0)
            return new BuildResult();

        var pages = BuildPages(recordingId, sourceRecordingFileName, jiraKey, assistive, warnings);
        BddActionMapDocument? map = null;

        if (!string.IsNullOrWhiteSpace(jiraKey))
        {
            map = BuildBddMap(
                recordingId,
                sourceRecordingFileName,
                createdAt,
                jiraKey!,
                assistive,
                pages,
                warnings);
        }

        return new BuildResult
        {
            Pages = pages,
            BddActionMap = map,
            Warnings = warnings
        };
    }

    private static List<PageCandidateDocument> BuildPages(
        string recordingId,
        string sourceRecordingFileName,
        string? jiraKey,
        List<RecordedAction> assistive,
        List<string> warnings)
    {
        var byPage = new Dictionary<string, PageBuildState>(StringComparer.Ordinal);

        foreach (var action in assistive)
        {
            var pageId = action.PageId;
            if (string.IsNullOrWhiteSpace(pageId))
                continue;

            if (!byPage.TryGetValue(pageId, out var state))
            {
                state = new PageBuildState
                {
                    PageId = pageId,
                    Name = action.Window?.Title is { Length: > 0 } title ? title : pageId
                };
                byPage[pageId] = state;
            }

            RegisterElement(state, action, action.Element, action.ObjectRef, action.EventId, recordingId, sourceRecordingFileName, jiraKey, warnings);
            if (action.TargetElement != null)
                RegisterElement(state, action, action.TargetElement, action.TargetObjectRef, action.EventId, recordingId, sourceRecordingFileName, jiraKey, warnings);
        }

        return byPage.Values
            .OrderBy(p => p.PageId, StringComparer.Ordinal)
            .Select(ToDocument)
            .ToList();
    }

    private static Dictionary<string, object?> BuildLocator(ElementInfo element)
    {
        var locator = new Dictionary<string, object?>(StringComparer.Ordinal);
        var hasAutomationId = !string.IsNullOrWhiteSpace(element.AutomationId);
        var hasName = !string.IsNullOrWhiteSpace(element.Name);
        var hasClassName = !string.IsNullOrWhiteSpace(element.ClassName);
        var controlTypeKnown = Phase3KnownControlTypes.IsKnown(element.ControlType);

        if (hasAutomationId)
            locator["automationId"] = element.AutomationId;

        // Phase 3 rejects unrecognized controlType values. When AutomationId is present,
        // omit an unsupported controlType so the candidate remains valid.
        if (!string.IsNullOrWhiteSpace(element.ControlType) && controlTypeKnown)
            locator["controlType"] = element.ControlType;

        // Without AutomationId, name/className are only useful when paired with a known controlType.
        if (hasName && (hasAutomationId || controlTypeKnown))
            locator["name"] = element.Name;

        if (hasClassName && (hasAutomationId || controlTypeKnown))
            locator["className"] = element.ClassName;

        return locator;
    }

    private static void RegisterElement(
        PageBuildState state,
        RecordedAction action,
        ElementInfo? element,
        string? objectRef,
        string? eventId,
        string recordingId,
        string sourceRecordingFileName,
        string? jiraKey,
        List<string> warnings)
    {
        if (element == null)
            return;

        var hasAutomationId = !string.IsNullOrWhiteSpace(element.AutomationId);
        var controlTypeKnown = Phase3KnownControlTypes.IsKnown(element.ControlType);
        var locator = BuildLocator(element);

        if (!Phase3LocatorContract.IsValid(locator, out var contractError)
            || string.IsNullOrWhiteSpace(objectRef)
            || !objectRef.Contains('.'))
        {
            var reason = !controlTypeKnown && !hasAutomationId
                ? $"controlType '{element.ControlType}' is not recognized by Phase 3 locator validation."
                : string.IsNullOrWhiteSpace(contractError)
                    ? "Insufficient stable locator for Assistive page-object candidate."
                    : contractError;

            state.Unresolved.Add(new Dictionary<string, object?>
            {
                ["eventId"] = eventId,
                ["pageId"] = state.PageId,
                ["reason"] = reason,
                ["automationId"] = element.AutomationId,
                ["controlType"] = element.ControlType,
                ["name"] = element.Name
            });
            return;
        }

        var elementId = objectRef[(objectRef.IndexOf('.') + 1)..];
        if (!state.Elements.TryGetValue(elementId, out var entry))
        {
            entry = new ElementBuildState
            {
                ElementId = elementId,
                Description = $"Assistive recording target {ElementInfo.GetLabel(element)}",
                Locator = locator,
                Grade = "medium"
            };
            state.Elements[elementId] = entry;
        }

        if (!string.IsNullOrWhiteSpace(eventId))
            entry.EventIds.Add(eventId!);

        entry.SourcePath = sourceRecordingFileName;
        entry.RecordingId = recordingId;
        entry.JiraKey = jiraKey;
        entry.WindowTitle = action.Window?.Title;
    }

    private static PageCandidateDocument ToDocument(PageBuildState state)
    {
        // Grade automationId locators as strong only when unique within the page.
        var automationIdCounts = state.Elements.Values
            .Select(e => e.Locator.TryGetValue("automationId", out var id) ? id as string : null)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .GroupBy(id => id!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

        foreach (var element in state.Elements.Values)
        {
            if (element.Locator.TryGetValue("automationId", out var idObj)
                && idObj is string id
                && automationIdCounts.TryGetValue(id, out var count)
                && count == 1)
            {
                element.Grade = "strong";
            }
            else if (!element.Locator.ContainsKey("automationId"))
            {
                element.Grade = "medium";
                if (!element.Locator.ContainsKey("name") && !element.Locator.ContainsKey("className"))
                    element.Warnings.Add("Locator may be ambiguous; manual review required.");
            }
        }

        var doc = new PageCandidateDocument
        {
            PageId = state.PageId,
            Name = state.Name,
            Elements = state.Elements
                .OrderBy(kv => kv.Key, StringComparer.Ordinal)
                .ToDictionary(
                    kv => kv.Key,
                    kv => new PageCandidateElement
                    {
                        Description = kv.Value.Description,
                        Locator = kv.Value.Locator,
                        Quality = new PageCandidateQuality
                        {
                            Grade = kv.Value.Grade,
                            Warnings = kv.Value.Warnings
                        },
                        Source = new PageCandidateSource
                        {
                            Kind = "capture",
                            Path = kv.Value.SourcePath,
                            Metadata = BuildMetadata(kv.Value)
                        }
                    },
                    StringComparer.Ordinal)
        };

        if (state.Unresolved.Count > 0)
            doc.Unresolved = state.Unresolved;

        return doc;
    }

    private static Dictionary<string, object?> BuildMetadata(ElementBuildState element)
    {
        var metadata = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["recordingId"] = element.RecordingId,
            ["eventIds"] = element.EventIds.Distinct(StringComparer.Ordinal).OrderBy(x => x, StringComparer.Ordinal).ToList()
        };
        if (!string.IsNullOrWhiteSpace(element.JiraKey))
            metadata["jiraKey"] = element.JiraKey;
        if (!string.IsNullOrWhiteSpace(element.WindowTitle))
            metadata["windowTitle"] = element.WindowTitle;
        return metadata;
    }

    private static BddActionMapDocument BuildBddMap(
        string recordingId,
        string sourceRecordingFileName,
        DateTimeOffset createdAt,
        string jiraKey,
        List<RecordedAction> assistive,
        List<PageCandidateDocument> pages,
        List<string> warnings)
    {
        var jiraActions = assistive
            .Where(a => string.Equals(a.JiraKey, jiraKey, StringComparison.Ordinal))
            .ToList();

        var groups = new Dictionary<string, BddActionMapGroup>(StringComparer.Ordinal);
        var groupOrder = new List<string>();
        var unmapped = new List<string>();

        foreach (var action in jiraActions)
        {
            var eventId = action.EventId ?? string.Empty;
            var operation = RecordedActionOperationResolver.ResolveOperation(action);
            if (string.IsNullOrWhiteSpace(operation))
            {
                if (!string.IsNullOrWhiteSpace(eventId))
                    unmapped.Add(eventId);
                warnings.Add(
                    $"Event {eventId} has unsupported playback operation for actionType '{action.ActionType}'.");
                continue;
            }

            if (action.Bdd is null)
            {
                if (!string.IsNullOrWhiteSpace(eventId))
                    unmapped.Add(eventId);
                continue;
            }

            if (!groups.TryGetValue(action.Bdd.GroupId, out var group))
            {
                group = new BddActionMapGroup
                {
                    GroupId = action.Bdd.GroupId,
                    Statement = action.Bdd.Statement
                };
                groups[action.Bdd.GroupId] = group;
                groupOrder.Add(action.Bdd.GroupId);
            }

            var actionRef = new BddActionMapActionRef
            {
                EventId = eventId,
                Sequence = action.Sequence ?? 0,
                PageId = action.PageId ?? string.Empty,
                ObjectRef = action.ObjectRef,
                TargetObjectRef = action.TargetObjectRef,
                Operation = operation
            };

            if (string.IsNullOrWhiteSpace(action.ObjectRef))
            {
                warnings.Add(
                    $"Event {eventId} is missing objectRef; manual locator review is required.");
            }

            group.Actions.Add(actionRef);
        }

        return new BddActionMapDocument
        {
            SchemaVersion = 1,
            RecordingId = recordingId,
            JiraKey = jiraKey,
            SourceRecording = sourceRecordingFileName,
            CreatedAt = createdAt,
            Pages = pages.Select(p => new BddActionMapPageRef
            {
                PageId = p.PageId,
                WindowTitle = p.Name,
                File = $"page-objects/{p.PageId}.page.json"
            }).ToList(),
            BddGroups = groupOrder.Select(id => groups[id]).ToList(),
            UnmappedEventIds = unmapped,
            Warnings = warnings.Count == 0 ? null : warnings.Distinct(StringComparer.Ordinal).ToList()
        };
    }

    private sealed class PageBuildState
    {
        public required string PageId { get; init; }
        public required string Name { get; init; }
        public Dictionary<string, ElementBuildState> Elements { get; } = new(StringComparer.Ordinal);
        public List<Dictionary<string, object?>> Unresolved { get; } = [];
    }

    private sealed class ElementBuildState
    {
        public required string ElementId { get; init; }
        public string? Description { get; set; }
        public Dictionary<string, object?> Locator { get; set; } = new(StringComparer.Ordinal);
        public string Grade { get; set; } = "medium";
        public List<string> Warnings { get; } = [];
        public HashSet<string> EventIds { get; } = new(StringComparer.Ordinal);
        public string? SourcePath { get; set; }
        public string? RecordingId { get; set; }
        public string? JiraKey { get; set; }
        public string? WindowTitle { get; set; }
    }
}
