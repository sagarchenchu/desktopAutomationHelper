using System;
using System.Collections.Generic;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Services;

namespace DesktopAutomationDriver.Services.Resolution;

public static class ElementTextExtractor
{
    public static string GetName(AutomationElement element)
    {
        return UiService.SafeElementName(element);
    }

    public static string GetValue(AutomationElement element)
    {
        return UiService.SafeElementValue(element);
    }

    public static string GetText(AutomationElement element)
    {
        return UiService.SafeElementText(element);
    }

    public static string GetLegacyName(AutomationElement element)
    {
        try
        {
            if (element.Patterns.LegacyIAccessible.IsSupported)
            {
                return element.Patterns.LegacyIAccessible.Pattern.Name.ValueOrDefault ?? "";
            }
        }
        catch { }
        return "";
    }

    public static string GetLegacyValue(AutomationElement element)
    {
        try
        {
            if (element.Patterns.LegacyIAccessible.IsSupported)
            {
                return element.Patterns.LegacyIAccessible.Pattern.Value.ValueOrDefault ?? "";
            }
        }
        catch { }
        return "";
    }

    public static List<string> GetAllPossibleTexts(AutomationElement element)
    {
        var list = new List<string>();
        AddIfNotEmpty(list, GetName(element));
        AddIfNotEmpty(list, GetValue(element));
        AddIfNotEmpty(list, GetText(element));
        AddIfNotEmpty(list, GetLegacyName(element));
        AddIfNotEmpty(list, GetLegacyValue(element));
        return list;
    }

    private static void AddIfNotEmpty(List<string> list, string val)
    {
        if (!string.IsNullOrWhiteSpace(val))
        {
            var trimmed = val.Trim();
            if (!list.Contains(trimmed))
            {
                list.Add(trimmed);
            }
        }
    }
}
