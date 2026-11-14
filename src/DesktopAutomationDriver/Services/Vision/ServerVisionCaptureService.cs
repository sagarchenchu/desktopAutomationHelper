using System.Text;
using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Captures live window + UIA under cursor via probes. PNG bytes are a deterministic
/// server-built envelope for the fake provider (real OCR adapters can replace later).
/// </summary>
public sealed class ServerVisionCaptureService : IVisionCaptureService
{
    private readonly ILiveWindowProbe _windowProbe;
    private readonly IUiaElementProbe _uiaProbe;

    public ServerVisionCaptureService(ILiveWindowProbe windowProbe, IUiaElementProbe uiaProbe)
    {
        _windowProbe = windowProbe ?? throw new ArgumentNullException(nameof(windowProbe));
        _uiaProbe = uiaProbe ?? throw new ArgumentNullException(nameof(uiaProbe));
    }

    public VisionServerCapture? Capture(IntPtr windowHandle, Point? cursorScreenPoint = null)
    {
        if (windowHandle == IntPtr.Zero)
            return null;

        var window = _windowProbe.ProbeByHandle(windowHandle);
        if (window is null)
            return null;

        DebugUiaNode? target = null;
        if (cursorScreenPoint is { } cursor)
        {
            var tree = _uiaProbe.CaptureAtPoint(cursor, maxAncestry: 4, maxChildren: 8);
            target = tree?.Target;
        }

        // Server-built envelope — never accept client PNG for issuance.
        var png = Encoding.UTF8.GetBytes(
            $"server-vision-capture:{window.WindowHandleHex}:{window.ProcessId}:{window.Bounds.Width}x{window.Bounds.Height}");

        return new VisionServerCapture
        {
            Window = window,
            SemanticTarget = target,
            PngBytes = png,
            Cursor = cursorScreenPoint
        };
    }
}

/// <summary>Test double for server capture.</summary>
public sealed class ScriptedVisionCaptureService : IVisionCaptureService
{
    private readonly Func<IntPtr, Point?, VisionServerCapture?> _capture;

    public ScriptedVisionCaptureService(Func<IntPtr, Point?, VisionServerCapture?> capture)
    {
        _capture = capture ?? throw new ArgumentNullException(nameof(capture));
    }

    public ScriptedVisionCaptureService(VisionServerCapture capture)
        : this((hwnd, _) => hwnd == capture.Window.WindowHandle ? capture : null)
    {
    }

    public VisionServerCapture? Capture(IntPtr windowHandle, Point? cursorScreenPoint = null) =>
        _capture(windowHandle, cursorScreenPoint);
}
