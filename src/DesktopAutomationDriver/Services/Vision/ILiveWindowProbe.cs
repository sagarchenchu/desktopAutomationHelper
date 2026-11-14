using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Fresh Win32/UIA probe of a window's HWND/PID/bounds/DPI.
/// Authorization must use this — never trust client-supplied window identity.
/// </summary>
public interface ILiveWindowProbe
{
    DebugWindowIdentity? ProbeByHandle(IntPtr windowHandle);
}
