using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class EditActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "Edit" && (operation == "type" || operation == "clear" || operation == "getvalue");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var edit = element.AsTextBox();
        if (request.Operation == "clear")
        {
            edit.Text = "";
            return new { success = true };
        }
        if (request.Operation == "type")
        {
            edit.Enter(request.Value ?? "");
            return new { success = true };
        }
        return edit.Text;
    }
}
