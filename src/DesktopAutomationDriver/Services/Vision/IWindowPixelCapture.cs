using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Captures real window pixels as PNG (not HWND/PID text placeholders).</summary>
public interface IWindowPixelCapture
{
    /// <summary>Returns PNG bytes for the window's on-screen bounds, or null on failure.</summary>
    byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds);
}

/// <summary>
/// CopyFromScreen capture of the live window <em>client</em> area.
/// Client-area avoids title-bar active/inactive chrome churn after interactive presence dialogs.
/// </summary>
public sealed class ScreenWindowPixelCapture : IWindowPixelCapture
{
    public byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds)
    {
        if (windowHandle == IntPtr.Zero)
            return null;

        // Prefer client-area screen rect so presence MessageBox focus changes do not
        // invalidate fingerprints via non-client chrome.
        var captureBounds = TryGetClientScreenBounds(windowHandle) ?? bounds;
        if (captureBounds.IsEmpty || captureBounds.Width <= 0 || captureBounds.Height <= 0)
            return null;

        try
        {
            using var bitmap = new Bitmap(captureBounds.Width, captureBounds.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(
                    captureBounds.X,
                    captureBounds.Y,
                    0,
                    0,
                    captureBounds.Size,
                    CopyPixelOperation.SourceCopy);
            }

            using var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Png);
            var bytes = stream.ToArray();
            // Guard against empty / non-image placeholders.
            if (bytes.Length < 32)
                return null;
            return bytes;
        }
        catch
        {
            return null;
        }
    }

    private static Rectangle? TryGetClientScreenBounds(IntPtr windowHandle)
    {
        try
        {
            if (!GetClientRect(windowHandle, out var client) || client.Right <= client.Left || client.Bottom <= client.Top)
                return null;

            var topLeft = new POINT { X = client.Left, Y = client.Top };
            var bottomRight = new POINT { X = client.Right, Y = client.Bottom };
            if (!ClientToScreen(windowHandle, ref topLeft) || !ClientToScreen(windowHandle, ref bottomRight))
                return null;

            return Rectangle.FromLTRB(topLeft.X, topLeft.Y, bottomRight.X, bottomRight.Y);
        }
        catch
        {
            return null;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);
}

/// <summary>Test double returning scripted PNG bytes.</summary>
public sealed class ScriptedWindowPixelCapture : IWindowPixelCapture
{
    private readonly Func<IntPtr, Rectangle, byte[]?> _capture;

    public ScriptedWindowPixelCapture(Func<IntPtr, Rectangle, byte[]?> capture)
    {
        _capture = capture ?? throw new ArgumentNullException(nameof(capture));
    }

    public ScriptedWindowPixelCapture(byte[] pngBytes)
        : this((_, _) => pngBytes)
    {
    }

    public byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds) =>
        _capture(windowHandle, bounds);
}
