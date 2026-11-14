using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Server-side capture of window identity, optional UIA target under cursor, and window-scoped bytes.
/// Callers must not supply forged screenshots or window identity for issuance.
/// </summary>
public interface IVisionCaptureService
{
    VisionServerCapture? Capture(IntPtr windowHandle, Point? cursorScreenPoint = null);
}

public sealed class VisionServerCapture
{
    public required DebugWindowIdentity Window { get; init; }
    public DebugUiaNode? SemanticTarget { get; init; }
    public required byte[] PngBytes { get; init; }
    public Point? Cursor { get; init; }
}
