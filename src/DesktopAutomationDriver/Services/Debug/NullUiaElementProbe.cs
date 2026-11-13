namespace DesktopAutomationDriver.Services.Debug;

/// <summary>Default probe used when no live UIA session is available (CI / headless hosts).</summary>
public sealed class NullUiaElementProbe : IUiaElementProbe
{
    public DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25) =>
        null;

    public int CountLocatorMatches(IReadOnlyDictionary<string, object?> locator, DebugWindowIdentity? scope) =>
        0;

    public HoverProbeResult ProbeHover(Point screenPoint) =>
        new()
        {
            ScreenPoint = screenPoint,
            NoOp = true,
            TooltipText = null
        };
}
