namespace DesktopAutomationDriver.Models.Playback;

/// <summary>
/// Summary returned by POST /playback after replaying submitted recording JSON.
/// </summary>
public class PlaybackResult
{
    public int TotalActions { get; set; }
    public int ExecutedActions { get; set; }
    public int SkippedActions { get; set; }
    public int FailedActions { get; set; }
    public bool Completed { get; set; }
    public List<PlaybackActionResult> Actions { get; set; } = [];
}
