namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Result returned by <see cref="Services.IRecordingService.StartRecording"/>.
/// </summary>
public class StartRecordingResult
{
    /// <summary>Non-null when starting failed (e.g. a session is already active).</summary>
    public string? Error { get; set; }

    /// <summary>
    /// Information about the application that was launched, or null when no
    /// <c>exePath</c> was provided.
    /// </summary>
    public LaunchInfo? Launch { get; set; }

    /// <summary>
    /// Resolved output path that will be used for the JSON export file,
    /// reflecting the caller-supplied value or the default temp directory.
    /// </summary>
    public string? OutputPath { get; set; }
}
