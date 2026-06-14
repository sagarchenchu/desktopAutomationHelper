using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class CheckBoxActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "CheckBox" && (operation == "check" || operation == "uncheck" || operation == "toggle" || operation == "ischecked");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var checkBox = element.AsCheckBox();
        if (request.Operation == "ischecked")
        {
            return checkBox.IsChecked;
        }

        bool? shouldCheck = null;
        if (request.Operation == "check") shouldCheck = true;
        else if (request.Operation == "uncheck") shouldCheck = false;

        if (shouldCheck.HasValue)
        {
            checkBox.IsChecked = shouldCheck.Value;
            return new { success = true, isChecked = shouldCheck.Value };
        }

        // toggle
        var targetState = !(checkBox.IsChecked ?? false);
        checkBox.IsChecked = targetState;
        return new { success = true, isChecked = targetState };
    }
}
