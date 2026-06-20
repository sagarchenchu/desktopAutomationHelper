using System.Drawing;
using System.Drawing.Imaging;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

internal static class NativeUiaScreenshot
{
    public static (bool success, string? base64, string? path, int width, int height, string? error) CaptureElement(
        IUIAutomationElement element,
        NativeUiaAutomation uia,
        string? outputPath)
    {
        var rect = uia.GetBoundingRectangle(element);
        if (!rect.HasValue || rect.Value.Width <= 0 || rect.Value.Height <= 0)
            return (false, null, null, 0, 0, "Element bounding rectangle is empty or unavailable.");

        try
        {
            using var bitmap = new Bitmap(rect.Value.Width, rect.Value.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(
                    rect.Value.Left,
                    rect.Value.Top,
                    0,
                    0,
                    new Size(rect.Value.Width, rect.Value.Height));
            }

            string? savedPath = null;
            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                bitmap.Save(outputPath, ImageFormat.Png);
                savedPath = outputPath;
            }

            using var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Png);
            var base64 = Convert.ToBase64String(stream.ToArray());

            return (true, base64, savedPath, rect.Value.Width, rect.Value.Height, null);
        }
        catch (Exception ex)
        {
            return (false, null, null, 0, 0, ex.Message);
        }
    }
}
