using DesktopAutomationDriver.Models.Recording;

namespace DesktopAutomationDriver.Models.Playback;

/// <summary>
/// Result for one recorded action during playback.
/// </summary>
public class PlaybackActionResult
{
    public int Index { get; set; }
    public bool Success { get; set; }
    public bool Skipped { get; set; }
    public string? SkipReason { get; set; }
    public string? Operation { get; set; }
    public ActionType? ActionType { get; set; }
    public string? Description { get; set; }
    public string? Error { get; set; }
    public object? Result { get; set; }
}
