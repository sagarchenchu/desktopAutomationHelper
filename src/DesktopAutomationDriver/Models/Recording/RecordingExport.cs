namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// The full recording export that is written to a JSON file when recording stops.
/// </summary>
public class RecordingExport
{
    /// <summary>UTC time recording started.</summary>
    public DateTimeOffset StartedAt { get; set; }

    /// <summary>UTC time recording stopped, or null if still running.</summary>
    public DateTimeOffset? StoppedAt { get; set; }

    /// <summary>The last active recording mode at the time of export.</summary>
    public string Mode { get; set; } = string.Empty;

    /// <summary>Path of the JSON file that was written, if any.</summary>
    public string? ExportedFilePath { get; set; }

    /// <summary>All actions recorded during this session.</summary>
    public List<RecordedAction> Actions { get; set; } = [];
}
