namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Details about the application that was launched when recording started.
/// </summary>
public class LaunchInfo
{
    /// <summary>Whether the application launched successfully.</summary>
    public bool Success { get; set; }

    /// <summary>OS process identifier of the launched application.</summary>
    public int? ProcessId { get; set; }

    /// <summary>Title of the application's main window (best-effort, may be empty).</summary>
    public string? WindowTitle { get; set; }

    /// <summary>Initial size, position and state of the launched application's main window.</summary>
    public ApplicationWindowInfo? Window { get; set; }

    /// <summary>Error message if the launch failed.</summary>
    public string? Error { get; set; }
}
