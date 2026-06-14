using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public interface IControlActionAdapter
{
    bool CanHandle(ElementSnapshot snapshot, string operation);
    object Execute(AutomationElement element, UiRequest request);
}
