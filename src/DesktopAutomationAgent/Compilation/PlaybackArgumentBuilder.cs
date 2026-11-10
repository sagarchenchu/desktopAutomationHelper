using System.Text.Json;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Compilation;

/// <summary>
/// Builds Phase 2 plan step arguments from a recording action.
/// Mirrors <c>PlaybackService.TryBuildUiRequest</c> without referencing the driver.
/// </summary>
internal static class PlaybackArgumentBuilder
{
    private static readonly Regex ObjectRefPattern = new(
        @"^[a-z][a-z0-9-]{0,63}\.[a-z][a-z0-9-]{0,63}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static bool TryBuildArguments(
        string operation,
        RecordingAction action,
        bool preferObjectRefs,
        out Dictionary<string, JsonElement> arguments,
        out string? skipReason)
    {
        arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        skipReason = null;

        var value = ResolveValue(action, operation);
        if (!string.IsNullOrWhiteSpace(value))
            arguments["value"] = JsonSerializer.SerializeToElement(value);

        var clickRegion = GetMetadata(action, "clickRegion")
            ?? GetMetadata(action, "headerClickRegion")
            ?? GetMetadata(action, "headerRegion");
        if (!string.IsNullOrWhiteSpace(clickRegion))
            arguments["clickRegion"] = JsonSerializer.SerializeToElement(clickRegion);

        var itemRegion = GetMetadata(action, "itemClickRegion")
            ?? GetMetadata(action, "itemRegion");
        if (!string.IsNullOrWhiteSpace(itemRegion))
            arguments["itemRegion"] = JsonSerializer.SerializeToElement(itemRegion);

        if (TryBuildLocator(operation, action, preferObjectRefs, out var locator, out skipReason))
        {
            if (locator is not null)
                arguments["locator"] = locator.Value;
        }
        else
        {
            return false;
        }

        if (RequiresLocator(operation) && !arguments.ContainsKey("locator"))
        {
            skipReason =
                $"Recorded action '{action.EventId}' / operation '{operation}' does not include enough element information.";
            return false;
        }

        return true;
    }

    public static bool IsSensitiveOperation(string operation) =>
        operation is "type" or "sendkeys" or "typeandselect" or "setvalue";

    private static bool TryBuildLocator(
        string operation,
        RecordingAction action,
        bool preferObjectRefs,
        out JsonElement? locator,
        out string? skipReason)
    {
        locator = null;
        skipReason = null;

        if (IsWindowOrAlertOperation(operation))
            return true;

        if (preferObjectRefs)
        {
            var objectRef = ResolveObjectRefForOperation(operation, action);
            if (!string.IsNullOrWhiteSpace(objectRef) && ObjectRefPattern.IsMatch(objectRef))
            {
                locator = JsonSerializer.SerializeToElement(new Dictionary<string, string>
                {
                    ["$objectRef"] = objectRef
                });
                return true;
            }
        }

        var element = ResolveElementForOperation(operation, action);
        var inline = ToLocatorElement(element);
        if (inline is null)
        {
            if (RequiresLocator(operation))
            {
                skipReason =
                    $"Recorded action '{action.EventId}' / operation '{operation}' has no usable locator.";
                return false;
            }

            return true;
        }

        locator = inline;
        return true;
    }

    private static string? ResolveObjectRefForOperation(string operation, RecordingAction action)
    {
        if (string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase))
            return action.TargetObjectRef ?? action.ObjectRef;

        if (string.Equals(operation, "selectdynamicmenuitem", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operation, "selectdynamicmenupath", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operation, "selectheaderdropdownitem", StringComparison.OrdinalIgnoreCase))
        {
            return action.TargetObjectRef ?? action.ObjectRef;
        }

        return action.ObjectRef;
    }

    private static RecordingElement? ResolveElementForOperation(string operation, RecordingAction action)
    {
        if (string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase))
            return action.TargetElement;

        if (string.Equals(operation, "selectdynamicmenuitem", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operation, "selectdynamicmenupath", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operation, "selectheaderdropdownitem", StringComparison.OrdinalIgnoreCase))
        {
            return action.TargetElement ?? action.Element;
        }

        return action.Element;
    }

    private static string? ResolveValue(RecordingAction action, string operation)
    {
        if (string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase)
            && action.MenuPath is { Count: > 0 })
        {
            return string.Join(">", action.MenuPath.Select(GetLabel));
        }

        if (string.Equals(operation, "closewindow", StringComparison.OrdinalIgnoreCase))
            return action.Value ?? action.Element?.Name;

        return action.Value;
    }

    private static JsonElement? ToLocatorElement(RecordingElement? element)
    {
        if (element is null)
            return null;

        if (!string.IsNullOrWhiteSpace(element.SuggestedXPath))
        {
            return JsonSerializer.SerializeToElement(new Dictionary<string, string>
            {
                ["xpath"] = element.SuggestedXPath!
            });
        }

        var locator = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!string.IsNullOrWhiteSpace(element.Name))
            locator["name"] = element.Name!;
        if (!string.IsNullOrWhiteSpace(element.AutomationId))
            locator["automationId"] = element.AutomationId!;
        if (!string.IsNullOrWhiteSpace(element.ClassName))
            locator["className"] = element.ClassName!;
        if (!string.IsNullOrWhiteSpace(element.ControlType))
            locator["controlType"] = element.ControlType!;

        return locator.Count == 0 ? null : JsonSerializer.SerializeToElement(locator);
    }

    private static bool RequiresLocator(string operation) =>
        !IsWindowOrAlertOperation(operation)
        && !string.Equals(operation, "clicklogicalmenupath", StringComparison.OrdinalIgnoreCase);

    private static bool IsWindowOrAlertOperation(string operation) =>
        string.Equals(operation, "switchwindow", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "maximize", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "minimize", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "closewindow", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "alertok", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "alertcancel", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "alertclose", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "popupok", StringComparison.OrdinalIgnoreCase);

    private static string? GetMetadata(RecordingAction action, string key)
    {
        if (action.Metadata is null)
            return null;

        return action.Metadata.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : null;
    }

    private static string GetLabel(RecordingElement? info) =>
        !string.IsNullOrEmpty(info?.Name) ? info.Name! :
        !string.IsNullOrEmpty(info?.AutomationId) ? info.AutomationId! :
        !string.IsNullOrEmpty(info?.ControlType) ? info.ControlType! :
        "(element)";
}
