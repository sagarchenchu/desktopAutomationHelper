using System;
using System.Drawing;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution;

public static class ElementScoring
{
    private const int NearParentDistanceThresholdPixels = 500;

    public static (int Score, string Reason) Score(
        AutomationElement element,
        UiLocator locator,
        ElementSearchOptions options,
        int? sessionProcessId)
    {
        int score = 0;
        var reasons = new System.Collections.Generic.List<string>();

        // 1. hwnd exact: +100
        if (locator.Hwnd.HasValue)
        {
            try
            {
                var handle = element.Properties.NativeWindowHandle.ValueOrDefault;
                if (handle != IntPtr.Zero && handle.ToInt64() == locator.Hwnd.Value)
                {
                    score += 100;
                    reasons.Add("hwnd-exact");
                }
            }
            catch {}
        }

        // 2. runtimeId exact: +90
        if (!string.IsNullOrEmpty(locator.RuntimeId))
        {
            try
            {
                var rtId = UiService.SafeRuntimeIdString(element);
                if (string.Equals(rtId, locator.RuntimeId, StringComparison.Ordinal))
                {
                    score += 90;
                    reasons.Add("runtimeId-exact");
                }
            }
            catch {}
        }

        // 3. automationId / autoId exact: +80
        var targetAid = locator.AutomationId ?? locator.AutoId;
        if (!string.IsNullOrEmpty(targetAid))
        {
            try
            {
                var aid = UiService.SafeElementAutomationId(element);
                if (string.Equals(aid, targetAid, StringComparison.OrdinalIgnoreCase))
                {
                    score += 80;
                    reasons.Add("automationId-exact");
                }
            }
            catch {}
        }

        // 4. controlType exact: +70
        if (!string.IsNullOrEmpty(locator.ControlType))
        {
            try
            {
                var ct = UiService.SafeElementControlType(element);
                if (string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase))
                {
                    score += 70;
                    reasons.Add("controlType-exact");
                }
            }
            catch {}
        }

        // 5. name exact: +60
        // 6. name contains: +45
        var targetName = locator.Name ?? locator.Title;
        if (!string.IsNullOrEmpty(targetName))
        {
            try
            {
                var name = UiService.SafeElementName(element);
                if (string.Equals(name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    score += 60;
                    reasons.Add("name-exact");
                }
                else if (name != null && name.Contains(targetName, StringComparison.OrdinalIgnoreCase))
                {
                    score += 45;
                    reasons.Add("name-contains");
                }
            }
            catch {}
        }

        // 7. className exact: +40
        if (!string.IsNullOrEmpty(locator.ClassName))
        {
            try
            {
                var cn = UiService.SafeElementClassName(element);
                if (string.Equals(cn, locator.ClassName, StringComparison.OrdinalIgnoreCase))
                {
                    score += 40;
                    reasons.Add("className-exact");
                }
            }
            catch {}
        }

        // 8. value exact: +35
        if (!string.IsNullOrEmpty(locator.Value))
        {
            try
            {
                var val = UiService.SafeElementValue(element);
                if (string.Equals(val, locator.Value, StringComparison.OrdinalIgnoreCase))
                {
                    score += 35;
                    reasons.Add("value-exact");
                }
            }
            catch {}
        }

        // 9. visible: +25
        try
        {
            var offscreen = UiService.SafeIsOffscreen(element) ?? false;
            var rect = element.BoundingRectangle;
            var isVisible = !offscreen && !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
            if (isVisible)
            {
                score += 25;
                reasons.Add("visible");
            }
        }
        catch {}

        // 10. enabled: +20
        try
        {
            var enabled = UiService.SafeIsEnabled(element) ?? false;
            if (enabled)
            {
                score += 20;
                reasons.Add("enabled");
            }
        }
        catch {}

        // 11. same process: +20
        if (sessionProcessId.HasValue)
        {
            try
            {
                var pid = UiService.SafeProcessId(element);
                if (pid == sessionProcessId.Value)
                {
                    score += 20;
                    reasons.Add("same-process");
                }
            }
            catch {}
        }

        // 12. near parent rectangle: +15
        if (options.Parent != null)
        {
            try
            {
                var rect = element.BoundingRectangle;
                var parentRect = options.Parent.BoundingRectangle;
                if (!rect.IsEmpty && !parentRect.IsEmpty)
                {
                    var center = new Point((int)(rect.Left + rect.Width / 2), (int)(rect.Top + rect.Height / 2));
                    var parentCenter = new Point((int)(parentRect.Left + parentRect.Width / 2), (int)(parentRect.Top + parentRect.Height / 2));
                    double dist = Math.Sqrt(Math.Pow(center.X - parentCenter.X, 2) + Math.Pow(center.Y - parentCenter.Y, 2));
                    if (dist < NearParentDistanceThresholdPixels)
                    {
                        score += 15;
                        reasons.Add("near-parent");
                    }
                }
            }
            catch {}
        }

        // 13. content element: +10
        try
        {
            if (element.Properties.IsContentElement.ValueOrDefault)
            {
                score += 10;
                reasons.Add("content-element");
            }
        }
        catch {}

        // 14. offscreen: -50 when includeOffscreen=false
        try
        {
            var offscreen = UiService.SafeIsOffscreen(element) ?? false;
            if (offscreen && !options.IncludeOffscreen)
            {
                score -= 50;
                reasons.Add("offscreen-penalty");
            }
        }
        catch {}

        // 15. empty rectangle: -25
        try
        {
            var rect = element.BoundingRectangle;
            if (rect.IsEmpty || rect.Width == 0 || rect.Height == 0)
            {
                score -= 25;
                reasons.Add("empty-rectangle-penalty");
            }
        }
        catch {}

        return (score, string.Join(";", reasons));
    }
}
