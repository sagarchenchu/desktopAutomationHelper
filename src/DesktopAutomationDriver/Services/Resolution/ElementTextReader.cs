using System;
using System.Collections.Generic;
using System.Linq;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public static class ElementTextReader
{
    public static List<string> ReadAllText(AutomationElement element)
    {
        var texts = new List<string>();
        if (element == null) return texts;

        AddTextIfNotEmpty(texts, UiService.SafeElementName(element));
        AddTextIfNotEmpty(texts, UiService.SafeElementValue(element));
        AddTextIfNotEmpty(texts, UiService.SafeElementText(element));

        try
        {
            var descendants = element.FindAllDescendants();
            foreach (var child in descendants)
            {
                AddTextIfNotEmpty(texts, UiService.SafeElementName(child));
                AddTextIfNotEmpty(texts, UiService.SafeElementValue(child));
                AddTextIfNotEmpty(texts, UiService.SafeElementText(child));
            }
        }
        catch { }

        return texts.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static void AddTextIfNotEmpty(List<string> list, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            list.Add(value.Trim());
        }
    }
}
