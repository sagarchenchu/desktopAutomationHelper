using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class MenuItemActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "MenuItem" && (operation == "click" || operation == "clickmenu" || operation == "expand" || operation == "collapse");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var menuItem = element.AsMenuItem();
        if (request.Operation == "expand")
        {
            menuItem.Expand();
            return new { success = true };
        }
        if (request.Operation == "collapse")
        {
            menuItem.Collapse();
            return new { success = true };
        }
        menuItem.Click();
        return new { success = true, strategy = "menuitem-click" };
    }
}
