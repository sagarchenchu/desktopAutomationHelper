using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;

namespace DesktopAutomationDriver.Services;

internal static class TypeCapabilityHelper
{
    public static bool IsTypeCapableElement(AutomationElement? element, ConditionFactory? conditionFactory)
    {
        if (element == null)
            return false;

        try
        {
            var ct = element.ControlType;

            if (ct == ControlType.Edit ||
                ct == ControlType.ComboBox ||
                ct == ControlType.Document)
            {
                return true;
            }

            if (ct == ControlType.Pane ||
                ct == ControlType.Custom ||
                ct == ControlType.Text)
            {
                if (IsFocusable(element))
                    return true;

                if (HasEditableChild(element, conditionFactory))
                    return true;
            }

            try
            {
                if (element.Patterns.Value.IsSupported)
                    return true;
            }
            catch
            {
                // Best effort only.
            }

            try
            {
                if (element.Patterns.Text.IsSupported)
                    return true;
            }
            catch
            {
                // Best effort only.
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    public static bool IsFocusable(AutomationElement element)
    {
        try
        {
            return element.Properties.IsKeyboardFocusable.ValueOrDefault;
        }
        catch
        {
            return false;
        }
    }

    public static bool HasEditableChild(AutomationElement element, ConditionFactory? conditionFactory)
    {
        try
        {
            if (conditionFactory == null)
                return false;

            var editChild = element.FindFirstDescendant(conditionFactory.ByControlType(ControlType.Edit));
            if (editChild != null)
                return true;

            var comboChild = element.FindFirstDescendant(conditionFactory.ByControlType(ControlType.ComboBox));
            if (comboChild != null)
                return true;

            return false;
        }
        catch
        {
            return false;
        }
    }

    public static bool TryFocusElement(AutomationElement element)
    {
        try
        {
            element.Focus();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static bool ShouldClickBeforeTyping(AutomationElement element)
    {
        try
        {
            var ct = element.ControlType;
            return ct == ControlType.Pane ||
                   ct == ControlType.Custom ||
                   ct == ControlType.Text;
        }
        catch
        {
            return false;
        }
    }
}
