using System.Drawing;
using System.Drawing.Imaging;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Captures real window pixels as PNG (not HWND/PID text placeholders).</summary>
public interface IWindowPixelCapture
{
    /// <summary>Returns PNG bytes for the window's on-screen bounds, or null on failure.</summary>
    byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds);
}

/// <summary>CopyFromScreen capture of the live window rectangle.</summary>
public sealed class ScreenWindowPixelCapture : IWindowPixelCapture
{
    public byte[]? CapturePng(IntPtr windowHandle, Rectangle bounds)
    {
        if (windowHandle == IntPtr.Zero)
            return null;
        if (bounds.IsEmpty || bounds.Width <= 0 || bounds.Height <= 0)
            return null;

        try
        {
            using var bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(
                    bounds.X,
                    bounds.Y,
                    0,
                    0,
                    bounds.Size,
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
