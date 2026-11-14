namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Abstraction over live UIA point/tree probes so Debug Mode logic stays testable
/// without embedding feature code in RecordingOverlayWindow or UiService.
/// </summary>
public interface IUiaElementProbe
{
    DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25);

    /// <summary>
    /// Live locator match count. Must report <see cref="LocatorMatchProbeResult.ProbeSucceeded"/>=false
    /// when the probe cannot run (missing scope, UIA failure) — never invent uniqueness from stale captures.
    /// </summary>
    LocatorMatchProbeResult CountLocatorMatches(
        IReadOnlyDictionary<string, object?> locator,
        DebugWindowIdentity? scope);

    HoverProbeResult ProbeHover(Point screenPoint);
}
