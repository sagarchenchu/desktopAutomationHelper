namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Captures the bounds of the screen that was active when recording started.
/// </summary>
public class ScreenResolutionInfo
{
    public string? DeviceName { get; set; }
    public bool IsPrimary { get; set; }
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int WorkingAreaX { get; set; }
    public int WorkingAreaY { get; set; }
    public int WorkingAreaWidth { get; set; }
    public int WorkingAreaHeight { get; set; }
}
