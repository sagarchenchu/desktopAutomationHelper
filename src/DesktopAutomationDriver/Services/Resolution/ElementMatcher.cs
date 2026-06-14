using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;

namespace DesktopAutomationDriver.Services.Resolution;

public static class ElementMatcher
{
    private const int DefaultNearPointTolerancePixels = 5;

    public static void Match(ElementCandidate candidate, UiLocator locator, string globalMatchMode = "exact")
    {
        var snapshot = candidate.Snapshot;
        var matchMode = locator.MatchMode ?? globalMatchMode;

        // 1. ControlType
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var expectedCt = NormalizeControlTypeAlias(locator.ControlType);
            if (string.Equals(snapshot.ControlType, expectedCt, StringComparison.OrdinalIgnoreCase))
            {
                candidate.MatchReasons.Add("controlType exact match");
            }
            else
            {
                candidate.RejectReasons.Add($"controlType mismatch: expected='{expectedCt}', actual='{snapshot.ControlType}'");
            }
        }

        // 2. ClassName
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            var mode = locator.ClassNameMatchMode ?? matchMode;
            if (EvaluateMatch(snapshot.ClassName, locator.ClassName, mode))
            {
                candidate.MatchReasons.Add($"className matched ({mode})");
            }
            else
            {
                candidate.RejectReasons.Add($"className mismatch: expected='{locator.ClassName}', actual='{snapshot.ClassName}' (mode={mode})");
            }
        }
        if (!string.IsNullOrWhiteSpace(locator.ClassNameRegex))
        {
            if (EvaluateRegex(snapshot.ClassName, locator.ClassNameRegex))
            {
                candidate.MatchReasons.Add("classNameRegex match");
            }
            else
            {
                candidate.RejectReasons.Add($"classNameRegex mismatch: pattern='{locator.ClassNameRegex}', actual='{snapshot.ClassName}'");
            }
        }

        // 3. Name
        var targetName = locator.Name ?? locator.Title;
        if (!string.IsNullOrWhiteSpace(targetName))
        {
            var mode = locator.NameMatchMode ?? matchMode;
            if (EvaluateMatch(snapshot.Name, targetName, mode))
            {
                candidate.MatchReasons.Add($"name matched ({mode})");
            }
            else
            {
                candidate.RejectReasons.Add($"name mismatch: expected='{targetName}', actual='{snapshot.Name}' (mode={mode})");
            }
        }
        if (!string.IsNullOrWhiteSpace(locator.NameRegex))
        {
            if (EvaluateRegex(snapshot.Name, locator.NameRegex))
            {
                candidate.MatchReasons.Add("nameRegex match");
            }
            else
            {
                candidate.RejectReasons.Add($"nameRegex mismatch: pattern='{locator.NameRegex}', actual='{snapshot.Name}'");
            }
        }

        // 4. AutomationId
        var targetAid = locator.AutomationId ?? locator.AutoId;
        if (!string.IsNullOrWhiteSpace(targetAid))
        {
            var mode = locator.AutomationIdMatchMode ?? "exact"; // automationId is usually exact
            if (EvaluateMatch(snapshot.AutomationId, targetAid, mode))
            {
                candidate.MatchReasons.Add($"automationId matched ({mode})");
            }
            else
            {
                candidate.RejectReasons.Add($"automationId mismatch: expected='{targetAid}', actual='{snapshot.AutomationId}' (mode={mode})");
            }
        }
        var targetAidRegex = locator.AutomationIdRegex ?? locator.AutoIdRegex;
        if (!string.IsNullOrWhiteSpace(targetAidRegex))
        {
            if (EvaluateRegex(snapshot.AutomationId, targetAidRegex))
            {
                candidate.MatchReasons.Add("automationIdRegex match");
            }
            else
            {
                candidate.RejectReasons.Add($"automationIdRegex mismatch: pattern='{targetAidRegex}', actual='{snapshot.AutomationId}'");
            }
        }

        // 5. Value
        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var mode = locator.ValueMatchMode ?? matchMode;
            if (EvaluateMatch(snapshot.Value, locator.Value, mode))
            {
                candidate.MatchReasons.Add($"value matched ({mode})");
            }
            else
            {
                candidate.RejectReasons.Add($"value mismatch: expected='{locator.Value}', actual='{snapshot.Value}' (mode={mode})");
            }
        }
        if (!string.IsNullOrWhiteSpace(locator.ValueRegex))
        {
            if (EvaluateRegex(snapshot.Value, locator.ValueRegex))
            {
                candidate.MatchReasons.Add("valueRegex match");
            }
            else
            {
                candidate.RejectReasons.Add($"valueRegex mismatch: pattern='{locator.ValueRegex}', actual='{snapshot.Value}'");
            }
        }

        // 6. Text
        if (!string.IsNullOrWhiteSpace(locator.Text))
        {
            var mode = locator.TextMatchMode ?? matchMode;
            if (EvaluateMatch(snapshot.Text, locator.Text, mode))
            {
                candidate.MatchReasons.Add($"text matched ({mode})");
            }
            else
            {
                candidate.RejectReasons.Add($"text mismatch: expected='{locator.Text}', actual='{snapshot.Text}' (mode={mode})");
            }
        }

        // 7. General Text/Value Search matching (Name, ValuePattern.Value, TextPattern text, LegacyIAccessible name/value)
        // If a bestMatch or general query without field is specified
        var targetBestMatch = locator.BestMatch;
        if (!string.IsNullOrWhiteSpace(targetBestMatch))
        {
            bool anyMatch =
                EvaluateMatch(snapshot.Name, targetBestMatch, "contains") ||
                EvaluateMatch(snapshot.Value, targetBestMatch, "contains") ||
                EvaluateMatch(snapshot.Text, targetBestMatch, "contains") ||
                EvaluateMatch(snapshot.LegacyName, targetBestMatch, "contains") ||
                EvaluateMatch(snapshot.LegacyValue, targetBestMatch, "contains");

            if (anyMatch)
            {
                candidate.MatchReasons.Add("bestMatch text overlap match");
            }
            else
            {
                candidate.RejectReasons.Add($"bestMatch mismatch: '{targetBestMatch}' not found in any text property");
            }
        }

        // 8. Native / Identity fields
        if (locator.ProcessId.HasValue || locator.Pid.HasValue)
        {
            var pid = locator.ProcessId ?? locator.Pid!.Value;
            if (snapshot.ProcessId == pid)
            {
                candidate.MatchReasons.Add("processId matched");
            }
            else
            {
                candidate.RejectReasons.Add($"processId mismatch: expected={pid}, actual={snapshot.ProcessId}");
            }
        }
        if (locator.Hwnd.HasValue)
        {
            if (snapshot.Hwnd == locator.Hwnd.Value)
            {
                candidate.MatchReasons.Add("hwnd matched");
            }
            else
            {
                candidate.RejectReasons.Add($"hwnd mismatch: expected={locator.Hwnd.Value}, actual={snapshot.Hwnd}");
            }
        }
        if (locator.ControlId.HasValue)
        {
            if (snapshot.ControlId == locator.ControlId.Value)
            {
                candidate.MatchReasons.Add("controlId matched");
            }
            else
            {
                candidate.RejectReasons.Add($"controlId mismatch: expected={locator.ControlId.Value}, actual={snapshot.ControlId}");
            }
        }
        if (!string.IsNullOrWhiteSpace(locator.FrameworkId))
        {
            if (string.Equals(snapshot.FrameworkId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
            {
                candidate.MatchReasons.Add("frameworkId matched");
            }
            else
            {
                candidate.RejectReasons.Add($"frameworkId mismatch: expected='{locator.FrameworkId}', actual='{snapshot.FrameworkId}'");
            }
        }
        if (!string.IsNullOrWhiteSpace(locator.RuntimeId))
        {
            // Runtime ID is exact Match
            // Let's assume candidate Snapshot runtimeId is represented in snapshot (or we can extract it if needed)
            // Wait, does Snapshot have RuntimeId? Let's check: Snapshot doesn't have RuntimeId in our newly updated class!
            // Wait, let's check: in the updated Snapshot class, did we include RuntimeId?
            // Ah! The problem statement in Section 4 "ElementSnapshot" did NOT have RuntimeId!
            // Wait, let's verify if we need to match it, we can get it from element itself or from the snapshot if we add it,
            // or match it directly from candidate.Element which is available!
            var rtId = UiService.SafeRuntimeIdString(candidate.Element);
            if (string.Equals(rtId, locator.RuntimeId, StringComparison.Ordinal))
            {
                candidate.MatchReasons.Add("runtimeId matched");
            }
            else
            {
                candidate.RejectReasons.Add($"runtimeId mismatch: expected='{locator.RuntimeId}', actual='{rtId}'");
            }
        }

        // 9. State filters
        if (locator.Visible.HasValue)
        {
            if (snapshot.IsVisible == locator.Visible.Value)
            {
                candidate.MatchReasons.Add($"visible={snapshot.IsVisible} matched");
            }
            else
            {
                candidate.RejectReasons.Add($"visible mismatch: expected={locator.Visible.Value}, actual={snapshot.IsVisible}");
            }
        }
        if (locator.Enabled.HasValue)
        {
            if (snapshot.IsEnabled == locator.Enabled.Value)
            {
                candidate.MatchReasons.Add($"enabled={snapshot.IsEnabled} matched");
            }
            else
            {
                candidate.RejectReasons.Add($"enabled mismatch: expected={locator.Enabled.Value}, actual={snapshot.IsEnabled}");
            }
        }
        if (locator.Offscreen.HasValue)
        {
            if (snapshot.IsOffscreen == locator.Offscreen.Value)
            {
                candidate.MatchReasons.Add($"offscreen={snapshot.IsOffscreen} matched");
            }
            else
            {
                candidate.RejectReasons.Add($"offscreen mismatch: expected={locator.Offscreen.Value}, actual={snapshot.IsOffscreen}");
            }
        }
        if (locator.IncludeOffscreen.HasValue && locator.IncludeOffscreen.Value == false)
        {
            if (snapshot.IsOffscreen == true)
            {
                candidate.RejectReasons.Add("offscreen=true rejected because includeOffscreen=false");
            }
        }

        // 10. Rectangle/Spatial filters
        MatchRectangleFilters(candidate, locator);
    }

    private static void MatchRectangleFilters(ElementCandidate candidate, UiLocator locator)
    {
        var rectObj = candidate.Element.BoundingRectangle;
        if (rectObj.IsEmpty)
        {
            if (locator.Left.HasValue || locator.Top.HasValue || locator.Right.HasValue || locator.Bottom.HasValue ||
                locator.Width.HasValue || locator.Height.HasValue || locator.NearX.HasValue || locator.NearY.HasValue)
            {
                candidate.RejectReasons.Add("element bounding rectangle is empty but spatial filters are requested");
            }
            return;
        }

        var tolerance = locator.Tolerance ?? 0;

        if (locator.Left.HasValue)
        {
            if (Math.Abs(rectObj.Left - locator.Left.Value) > tolerance)
                candidate.RejectReasons.Add($"left spatial mismatch: expected={locator.Left.Value}, actual={rectObj.Left} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("left spatial match");
        }

        if (locator.Top.HasValue)
        {
            if (Math.Abs(rectObj.Top - locator.Top.Value) > tolerance)
                candidate.RejectReasons.Add($"top spatial mismatch: expected={locator.Top.Value}, actual={rectObj.Top} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("top spatial match");
        }

        if (locator.Right.HasValue)
        {
            if (Math.Abs(rectObj.Right - locator.Right.Value) > tolerance)
                candidate.RejectReasons.Add($"right spatial mismatch: expected={locator.Right.Value}, actual={rectObj.Right} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("right spatial match");
        }

        if (locator.Bottom.HasValue)
        {
            if (Math.Abs(rectObj.Bottom - locator.Bottom.Value) > tolerance)
                candidate.RejectReasons.Add($"bottom spatial mismatch: expected={locator.Bottom.Value}, actual={rectObj.Bottom} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("bottom spatial match");
        }

        if (locator.Width.HasValue)
        {
            if (Math.Abs(rectObj.Width - locator.Width.Value) > tolerance)
                candidate.RejectReasons.Add($"width spatial mismatch: expected={locator.Width.Value}, actual={rectObj.Width} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("width spatial match");
        }

        if (locator.Height.HasValue)
        {
            if (Math.Abs(rectObj.Height - locator.Height.Value) > tolerance)
                candidate.RejectReasons.Add($"height spatial mismatch: expected={locator.Height.Value}, actual={rectObj.Height} (tolerance={tolerance})");
            else
                candidate.MatchReasons.Add("height spatial match");
        }

        if (locator.NearX.HasValue && locator.NearY.HasValue)
        {
            var nearTolerance = locator.Tolerance ?? DefaultNearPointTolerancePixels;

            var containsNearPoint =
                locator.NearX.Value >= rectObj.Left - nearTolerance &&
                locator.NearX.Value <= rectObj.Right + nearTolerance &&
                locator.NearY.Value >= rectObj.Top - nearTolerance &&
                locator.NearY.Value <= rectObj.Bottom + nearTolerance;

            if (!containsNearPoint)
                candidate.RejectReasons.Add($"nearPoint spatial mismatch: expected near ({locator.NearX.Value},{locator.NearY.Value}), element rect=[{rectObj.Left},{rectObj.Top},{rectObj.Right},{rectObj.Bottom}] (nearTolerance={nearTolerance})");
            else
                candidate.MatchReasons.Add("nearPoint spatial match");
        }

        if (locator.ContainsPoint == true && locator.NearX.HasValue && locator.NearY.HasValue)
        {
            var containsPoint =
                locator.NearX.Value >= rectObj.Left &&
                locator.NearX.Value <= rectObj.Right &&
                locator.NearY.Value >= rectObj.Top &&
                locator.NearY.Value <= rectObj.Bottom;

            if (!containsPoint)
                candidate.RejectReasons.Add($"containsPoint spatial mismatch: point ({locator.NearX.Value},{locator.NearY.Value}) not inside element rect=[{rectObj.Left},{rectObj.Top},{rectObj.Right},{rectObj.Bottom}]");
            else
                candidate.MatchReasons.Add("containsPoint spatial match");
        }

        if (locator.IntersectsRectangle == true && locator.Left.HasValue && locator.Top.HasValue && locator.Right.HasValue && locator.Bottom.HasValue)
        {
            var intersects =
                !(rectObj.Left > locator.Right.Value ||
                  rectObj.Right < locator.Left.Value ||
                  rectObj.Top > locator.Bottom.Value ||
                  rectObj.Bottom < locator.Top.Value);

            if (!intersects)
                candidate.RejectReasons.Add("intersectsRectangle spatial mismatch");
            else
                candidate.MatchReasons.Add("intersectsRectangle spatial match");
        }
    }

    private static string NormalizeControlTypeAlias(string controlType)
    {
        if (string.IsNullOrWhiteSpace(controlType)) return string.Empty;
        var norm = controlType.Trim().ToLowerInvariant();
        switch (norm)
        {
            case "dialog":
                return "Window";
            case "button":
                return "Button";
            case "text":
                return "Text";
            case "pane":
                return "Pane";
            case "list item":
            case "listitem":
                return "ListItem";
            default:
                if (norm.Length > 0)
                {
                    return char.ToUpperInvariant(norm[0]) + norm.Substring(1);
                }
                return controlType;
        }
    }

    private static bool EvaluateMatch(string? actual, string? expected, string mode)
    {
        if (expected == null) return true;
        if (actual == null) return false;

        var act = actual.Trim();
        var exp = expected.Trim();

        switch (mode.ToLowerInvariant())
        {
            case "contains":
                return act.Contains(exp, StringComparison.OrdinalIgnoreCase);
            case "startswith":
                return act.StartsWith(exp, StringComparison.OrdinalIgnoreCase);
            case "endswith":
                return act.EndsWith(exp, StringComparison.OrdinalIgnoreCase);
            case "regex":
                return EvaluateRegex(act, exp);
            case "exact":
            default:
                return string.Equals(act, exp, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static bool EvaluateRegex(string? actual, string? pattern)
    {
        if (string.IsNullOrEmpty(pattern)) return true;
        if (actual == null) return false;
        try
        {
            return Regex.IsMatch(actual, pattern, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
        }
        catch
        {
            return false;
        }
    }
}
