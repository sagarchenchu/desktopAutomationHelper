namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Captures the coordinate mapping details used to resolve a right-click target.
/// </summary>
public class PointerContextInfo
{
    public PointerCoordinateInfo? HookPoint { get; set; }
    public PointerCoordinateInfo? CursorPoint { get; set; }
    public PointerCoordinateInfo? ResolvedPoint { get; set; }
    public bool CoordinateMismatch { get; set; }
    public bool UsedCursorFallback { get; set; }
    public bool HookPointInTarget { get; set; }
    public bool CursorPointInTarget { get; set; }
}
