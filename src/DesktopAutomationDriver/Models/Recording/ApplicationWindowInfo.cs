namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Captures the launched application's initial window position and size.
/// </summary>
public class ApplicationWindowInfo
{
    public string? Title { get; set; }
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public string WindowState { get; set; } = "unknown";
    public bool IsMaximized { get; set; }
    public bool IsMinimized { get; set; }
    public bool IsFullScreen { get; set; }
}
