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
/// Captures real window client pixels via PrintWindow (not HWND/PID text).
/// PrintWindow avoids CopyFromScreen composition/cursor/occlusion churn that
/// otherwise invalidates approve fingerprints after the operator confirmation dialog.
/// </summary>
public sealed class ScreenWindowPixelCapture : IWindowPixelCapture
{
    private const uint PwClientOnly = 0x1;
    private const uint PwRenderFullContent = 0x2;

    public byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds)
    {
        if (windowHandle == IntPtr.Zero)
            return null;

        if (!GetClientRect(windowHandle, out var client)
            || client.Right <= client.Left
            || client.Bottom <= client.Top)
        {
            if (bounds.IsEmpty || bounds.Width <= 0 || bounds.Height <= 0)
                return null;
            return CaptureViaCopyFromScreen(bounds);
        }

        var width = client.Right - client.Left;
        var height = client.Bottom - client.Top;
        if (width <= 0 || height <= 0)
            return null;

        try
        {
            using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                var hdc = graphics.GetHdc();
                var printed = false;
                try
                {
                    printed = PrintWindow(windowHandle, hdc, PwClientOnly | PwRenderFullContent)
                        || PrintWindow(windowHandle, hdc, PwClientOnly)
                        || PrintWindow(windowHandle, hdc, PwRenderFullContent)
                        || PrintWindow(windowHandle, hdc, 0);
                }
                finally
                {
                    graphics.ReleaseHdc(hdc);
                }

                if (!printed)
                    return CaptureViaCopyFromScreen(TryGetClientScreenBounds(windowHandle) ?? bounds);
            }

            using var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Png);
            var bytes = stream.ToArray();
            if (bytes.Length < 32)
                return null;
            return bytes;
        }
        catch
        {
            return CaptureViaCopyFromScreen(TryGetClientScreenBounds(windowHandle) ?? bounds);
        }
    }

    private static byte[]? CaptureViaCopyFromScreen(Rectangle captureBounds)
    {
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

    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
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
