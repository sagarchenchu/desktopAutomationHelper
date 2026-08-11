namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Request body for POST /record/start.
/// </summary>
public class StartRecordingRequest
{
    /// <summary>
    /// Full path to the application executable that should be launched when recording starts.
    /// Optional – if omitted the overlay is shown without launching an application.
    /// </summary>
    public string? ExePath { get; set; }

    /// <summary>
    /// Output directory where the recorded JSON file will be saved.
    /// The driver creates a deterministic <c>recording_&lt;timestamp&gt;.json</c> filename
    /// inside this directory. This value is always treated as a directory (not a file path).
    /// Optional – defaults to %TEMP%\DesktopAutomationHelper\Recordings\.
    /// </summary>
    public string? OutputPath { get; set; }

    /// <summary>
    /// Number of seconds the recording session should remain active before it is
    /// automatically stopped and the JSON file is written.
    /// Optional – if omitted the session runs until the user presses Ctrl+S or
    /// POST /record/stop is called.
    /// </summary>
    public int? WaitSeconds { get; set; }
}
