using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class ListItemActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "ListItem" && (operation == "select" || operation == "click");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var listItem = element.AsListBoxItem();
        if (request.Operation == "select" || request.Operation == "click")
        {
            listItem.Select();
            return new { success = true, strategy = "listitem-select" };
        }
        return new { success = false };
    }
}
