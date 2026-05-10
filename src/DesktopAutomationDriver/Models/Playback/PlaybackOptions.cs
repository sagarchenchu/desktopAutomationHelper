namespace DesktopAutomationDriver.Models.Playback;

/// <summary>
/// Options accepted alongside a playback recording payload.
/// </summary>
public class PlaybackOptions
{
    /// <summary>When true, continue after a failed action. Defaults to false.</summary>
    public bool ContinueOnError { get; set; }

    /// <summary>Optional delay between successfully executed actions.</summary>
    public int DelayMs { get; set; }
}
