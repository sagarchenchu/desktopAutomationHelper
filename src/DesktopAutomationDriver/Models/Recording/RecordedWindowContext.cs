namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Immutable top-level window context captured before an Assistive action executes.
/// </summary>
public sealed class RecordedWindowContext
{
    public string? Title { get; set; }

    public string? NormalizedTitle { get; set; }

    public int? ProcessId { get; set; }

    /// <summary>Diagnostic HWND string such as <c>0x001A04F2</c>. Not used as a locator.</summary>
    public string? NativeWindowHandle { get; set; }
}
