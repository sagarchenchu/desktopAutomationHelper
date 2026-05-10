using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Playback;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Replays actions from an Assistive recording by converting them to existing /ui operations.
/// </summary>
public sealed class PlaybackService : IPlaybackService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    private readonly IUiService _uiService;
    private readonly ILogger<PlaybackService> _logger;

    public PlaybackService(IUiService uiService, ILogger<PlaybackService> logger)
    {
        _uiService = uiService;
        _logger = logger;
    }

    public PlaybackResult Play(JsonElement payload)
    {
        var (recording, options) = ParsePayload(payload);
        var actions = recording.Actions ?? [];
        var result = new PlaybackResult { TotalActions = actions.Count };

        for (var i = 0; i < actions.Count; i++)
        {
            var action = actions[i];
            var actionResult = new PlaybackActionResult
            {
                Index = i,
                ActionType = action.ActionType,
                Description = action.Description
            };

            result.Actions.Add(actionResult);

            if (action.Mode != RecordingMode.Assistive)
            {
                MarkSkipped(result, actionResult, "Only Assistive recording actions are played back.");
                continue;
            }

            if (!TryBuildUiRequest(action, out var uiRequest, out var skipReason))
            {
                MarkSkipped(result, actionResult, skipReason ?? "Action cannot be mapped to a /ui operation.");
                continue;
            }

            actionResult.Operation = uiRequest.Operation;

            try
            {
                actionResult.Result = _uiService.Execute(uiRequest);
                actionResult.Success = true;
                result.ExecutedActions++;

                if (options.DelayMs > 0 && i < actions.Count - 1)
                    Thread.Sleep(options.DelayMs);
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
            {
                actionResult.Error = ex.Message;
                result.FailedActions++;
                _logger.LogWarning(ex, "Playback action {Index} failed. operation={Operation}", i, uiRequest.Operation);

                if (!options.ContinueOnError)
                {
                    result.Completed = false;
                    return result;
                }
            }
        }

        result.Completed = result.FailedActions == 0;
        return result;
    }

    private static void MarkSkipped(
        PlaybackResult result,
        PlaybackActionResult actionResult,
        string reason)
    {
        actionResult.Skipped = true;
        actionResult.SkipReason = reason;
        result.SkippedActions++;
    }

    private static (RecordingExport Recording, PlaybackOptions Options) ParsePayload(JsonElement payload)
    {
        if (payload.ValueKind == JsonValueKind.Undefined || payload.ValueKind == JsonValueKind.Null)
            throw new ArgumentException("Playback request body is required.");

        var options = new PlaybackOptions();
        RecordingExport? recording = null;

        if (payload.ValueKind == JsonValueKind.Array)
        {
            recording = new RecordingExport
            {
                Mode = RecordingMode.Assistive.ToString(),
                Actions = Deserialize<List<RecordedAction>>(payload) ?? []
            };
        }
        else if (payload.ValueKind == JsonValueKind.Object)
        {
            options.ContinueOnError = GetBoolean(payload, "continueOnError") ?? false;
            options.DelayMs = Math.Max(0, GetInt32(payload, "delayMs") ?? 0);

            if (TryGetProperty(payload, "recording", out var recordingElement))
            {
                recording = Deserialize<RecordingExport>(recordingElement);
            }
            else if (TryGetProperty(payload, "value", out var valueElement) &&
                     valueElement.ValueKind == JsonValueKind.Object &&
                     TryGetProperty(valueElement, "actions", out _))
            {
                recording = Deserialize<RecordingExport>(valueElement);
            }
            else if (TryGetProperty(payload, "actions", out _))
            {
                recording = Deserialize<RecordingExport>(payload);
            }
        }

        if (recording == null)
            throw new ArgumentException("Playback request must be a recording export JSON object, an actions array, or { recording: ... }.");

        return (recording, options);
    }

    private static T? Deserialize<T>(JsonElement element) =>
        JsonSerializer.Deserialize<T>(element.GetRawText(), JsonOptions);

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }

        value = default;
        return false;
    }

    private static bool? GetBoolean(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value))
            return null;

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => null
        };
    }

    private static int? GetInt32(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value))
            return null;

        return value.TryGetInt32(out var result) ? result : null;
    }

    private static bool TryBuildUiRequest(
        RecordedAction action,
        out UiRequest request,
        out string? skipReason)
    {
        request = new UiRequest();
        skipReason = null;

        var operation = ResolveOperation(action);
        if (operation == null)
        {
            skipReason = $"Unsupported recorded action type '{action.ActionType}'.";
            return false;
        }

        request.Operation = operation;
        request.Value = ResolveValue(action, operation);
        request.Locator = ResolveLocator(action, operation);
        request.ClickRegion = GetMetadataValue(action, "clickRegion")
            ?? GetMetadataValue(action, "headerClickRegion")
            ?? GetMetadataValue(action, "headerRegion");
        request.ItemRegion = GetMetadataValue(action, "itemClickRegion")
            ?? GetMetadataValue(action, "itemRegion");

        if (RequiresLocator(operation) && request.Locator == null)
        {
            skipReason = $"Recorded action '{action.ActionType}' does not include enough element information for playback.";
            return false;
        }

        return true;
    }

    private static string? ResolveOperation(RecordedAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.Operation))
            return action.Operation;

        return action.ActionType switch
        {
            ActionType.Click => "click",
            ActionType.MenuPathClick => "clicklogicalmenupath",
            ActionType.DoubleClick => "doubleclick",
            ActionType.RightClick => "rightclick",
            ActionType.Hover => "hover",
            ActionType.Select => "select",
            ActionType.Type => IsSendKeysValue(action.Value) ? "sendkeys" : "type",
            ActionType.TypeAndSelect => "typeandselect",
            ActionType.IsVisible => "isvisible",
            ActionType.IsClickable => "isclickable",
            ActionType.IsEnabled => "isenabled",
            ActionType.IsDisabled => "isenabled",
            ActionType.IsEditable => "iseditable",
            ActionType.GetTableHeaders => "gettableheaders",
            ActionType.GetTableData => "gettabledata",
            ActionType.IsChecked => "ischecked",
            ActionType.SelectCheckBox => ResolveCheckOperation(action),
            ActionType.ClearText => "clear",
            ActionType.GetValue => "getvalue",
            ActionType.Expand => "click",
            ActionType.Collapse => "click",
            ActionType.Maximize => "maximize",
            ActionType.Minimize => "minimize",
            ActionType.CloseWindow => "closewindow",
            ActionType.SwitchWindow => "switchwindow",
            ActionType.SetValue => "type",
            ActionType.Scroll => "scroll",
            _ => null
        };
    }

    private static string? ResolveCheckOperation(RecordedAction action)
    {
        if (action.Description?.Contains("Uncheck", StringComparison.OrdinalIgnoreCase) == true)
            return "uncheck";

        return "check";
    }

    private static string? ResolveValue(RecordedAction action, string operation)
    {
        if (string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase) &&
            action.MenuPath is { Count: > 0 })
        {
            return string.Join(">", action.MenuPath.Select(ElementInfo.GetLabel));
        }

        return action.Value;
    }

    private static UiLocator? ResolveLocator(RecordedAction action, string operation)
    {
        if (string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase))
            return action.TargetElement != null ? ToLocator(action.TargetElement) : null;

        if (string.Equals(operation, "selectdynamicmenuitem", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "selectdynamicmenupath", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "selectheaderdropdownitem", StringComparison.OrdinalIgnoreCase))
        {
            return ToLocator(action.TargetElement ?? action.Element);
        }

        if (string.Equals(operation, "switchwindow", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "maximize", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "minimize", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ToLocator(action.Element);
    }

    private static UiLocator? ToLocator(ElementInfo? element)
    {
        if (element == null)
            return null;

        if (!string.IsNullOrWhiteSpace(element.SuggestedXPath))
            return new UiLocator { XPath = element.SuggestedXPath };

        if (string.IsNullOrWhiteSpace(element.Name) &&
            string.IsNullOrWhiteSpace(element.AutomationId) &&
            string.IsNullOrWhiteSpace(element.ClassName) &&
            string.IsNullOrWhiteSpace(element.ControlType))
        {
            return null;
        }

        return new UiLocator
        {
            Name = string.IsNullOrWhiteSpace(element.Name) ? null : element.Name,
            AutomationId = string.IsNullOrWhiteSpace(element.AutomationId) ? null : element.AutomationId,
            ClassName = string.IsNullOrWhiteSpace(element.ClassName) ? null : element.ClassName,
            ControlType = string.IsNullOrWhiteSpace(element.ControlType) ? null : element.ControlType
        };
    }

    private static bool RequiresLocator(string operation)
    {
        return !string.Equals(operation, "switchwindow", StringComparison.OrdinalIgnoreCase) &&
               !string.Equals(operation, "maximize", StringComparison.OrdinalIgnoreCase) &&
               !string.Equals(operation, "minimize", StringComparison.OrdinalIgnoreCase) &&
               !string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetMetadataValue(RecordedAction action, string key)
    {
        if (action.Metadata == null)
            return null;

        return action.Metadata.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : null;
    }

    private static bool IsSendKeysValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var trimmed = value.Trim();
        return trimmed.StartsWith('{') && trimmed.EndsWith('}');
    }
}
