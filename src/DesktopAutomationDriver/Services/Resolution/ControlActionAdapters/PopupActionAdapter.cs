using System;
using System.Linq;
using System.Threading;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class PopupActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return (snapshot.ControlType == "Window" || snapshot.ClassName == "#32770") && 
               (operation == "popupaction" || operation == "popuptext");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var win = element.AsWindow();
        if (request.Operation == "popuptext")
        {
            var texts = ElementTextExtractor.GetAllPossibleTexts(win);
            return new { success = true, texts };
        }

        var action = request.Action?.ToLowerInvariant() ?? "button";
        if (action == "close")
        {
            win.Close();
            return new { success = true, action = "close" };
        }

        if (action == "enter")
        {
            win.Focus();
            Thread.Sleep(50);
            FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
            return new { success = true, action = "enter" };
        }

        if (action == "escape")
        {
            win.Focus();
            Thread.Sleep(50);
            FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ESCAPE);
            return new { success = true, action = "escape" };
        }

        if (action == "button")
        {
            var buttonName = request.Button ?? "OK";
            var buttons = buttonName.Split('|');
            foreach (var b in buttons)
            {
                var btnElement = win.FindFirstDescendant(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button).And(cf.ByName(b)));
                if (btnElement != null)
                {
                    btnElement.AsButton().Click();
                    return new { success = true, action = "button-click", buttonClicked = b };
                }
            }
            throw new InvalidOperationException($"No buttons matching '{buttonName}' found inside popup.");
        }

        return new { success = false, message = $"Unknown action '{action}'" };
    }
}
