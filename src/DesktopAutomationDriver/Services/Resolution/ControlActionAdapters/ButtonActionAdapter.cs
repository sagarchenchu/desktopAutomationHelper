using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class ButtonActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "Button" && (operation == "click" || operation == "doubleclick" || operation == "rightclick");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        if (element.Patterns.Invoke.IsSupported)
        {
            element.Patterns.Invoke.Pattern.Invoke();
            return new { success = true, strategy = "uia-invoke" };
        }
        element.Click();
        return new { success = true, strategy = "physical-click" };
    }
}
