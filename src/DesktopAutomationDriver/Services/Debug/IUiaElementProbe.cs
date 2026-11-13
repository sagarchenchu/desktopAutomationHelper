namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Abstraction over live UIA point/tree probes so Debug Mode logic stays testable
/// without embedding feature code in RecordingOverlayWindow or UiService.
/// </summary>
public interface IUiaElementProbe
{
    DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25);

    int CountLocatorMatches(IReadOnlyDictionary<string, object?> locator, DebugWindowIdentity? scope);

    HoverProbeResult ProbeHover(Point screenPoint);
}
