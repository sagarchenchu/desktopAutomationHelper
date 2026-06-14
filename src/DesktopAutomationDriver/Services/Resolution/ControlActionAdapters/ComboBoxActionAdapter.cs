using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class ComboBoxActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "ComboBox" && (operation == "select" || operation == "getvalue");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var combo = element.AsComboBox();
        if (request.Operation == "getvalue")
        {
            return combo.Value;
        }
        if (!string.IsNullOrWhiteSpace(request.Value))
        {
            combo.Select(request.Value);
            return new { success = true, strategy = "combobox-select-by-value" };
        }
        if (request.Index.HasValue)
        {
            combo.Select(request.Index.Value);
            return new { success = true, strategy = "combobox-select-by-index" };
        }
        return new { success = false, message = "Neither Value nor Index specified" };
    }
}
