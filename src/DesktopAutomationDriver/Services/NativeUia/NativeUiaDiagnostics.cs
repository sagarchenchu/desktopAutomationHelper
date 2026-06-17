using System.Drawing;

namespace DesktopAutomationDriver.Services.NativeUia;

internal static class NativeUiaDiagnostics
{
    public static object CandidateDiagnostic(int index, NativeUiaElementSnapshot snapshot)
    {
        return new
        {
            index,
            automationId = snapshot.AutomationId,
            name = snapshot.Name,
            className = snapshot.ClassName,
            controlType = snapshot.ControlType,
            frameworkId = snapshot.FrameworkId,
            processId = snapshot.ProcessId,
            rectangle = ToRectangleObject(snapshot.BoundingRectangle),
            isEnabled = snapshot.IsEnabled,
            isOffscreen = snapshot.IsOffscreen,
            supportedPatterns = snapshot.SupportedPatterns
        };
    }

    public static object ToRectangleObject(object? boundingRectangle)
    {
        if (boundingRectangle is Rectangle rect)
        {
            return new
            {
                left = rect.Left,
                top = rect.Top,
                right = rect.Right,
                bottom = rect.Bottom,
                width = rect.Width,
                height = rect.Height
            };
        }

        return new { left = 0, top = 0, right = 0, bottom = 0, width = 0, height = 0 };
    }

    public static object ComboSummary(NativeUiaElementSnapshot snapshot) => new
    {
        automationId = snapshot.AutomationId,
        name = snapshot.Name,
        controlType = snapshot.ControlType,
        className = snapshot.ClassName,
        frameworkId = snapshot.FrameworkId,
        processId = snapshot.ProcessId,
        rectangle = ToRectangleObject(snapshot.BoundingRectangle),
        isEnabled = snapshot.IsEnabled,
        isOffscreen = snapshot.IsOffscreen,
        supportedPatterns = snapshot.SupportedPatterns,
        value = snapshot.Value
    };
}
