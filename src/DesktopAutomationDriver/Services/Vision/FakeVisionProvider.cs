using System.Security.Cryptography;
using System.Text;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Deterministic fake OCR/vision provider for CI. Never auto-promotes.</summary>
public sealed class FakeVisionProvider : IVisionProvider
{
    public string Name => "fake";

    public IReadOnlyList<VisionCandidateTarget> Analyze(VisionScreenshot screenshot, string? query = null)
    {
        ArgumentNullException.ThrowIfNull(screenshot);

        var label = string.IsNullOrWhiteSpace(query) ? "fake-target" : query.Trim();
        var fingerprint = Convert.ToHexString(SHA256.HashData(screenshot.PngBytes))[..16].ToLowerInvariant();

        // Place a synthetic candidate in the upper-left quadrant of the window.
        var bounds = new Rectangle(
            screenshot.WindowBounds.X + Math.Max(4, screenshot.WindowBounds.Width / 8),
            screenshot.WindowBounds.Y + Math.Max(4, screenshot.WindowBounds.Height / 8),
            Math.Max(8, screenshot.WindowBounds.Width / 6),
            Math.Max(8, screenshot.WindowBounds.Height / 10));

        return
        [
            new VisionCandidateTarget
            {
                Label = label,
                Bounds = bounds,
                Confidence = 0.91,
                Kind = "ocr",
                AutoPromoted = false,
                Window = screenshot.Window,
                ProducedAtUtc = DateTimeOffset.UtcNow,
                ScreenshotFingerprint = fingerprint + ":" + Encoding.UTF8.GetByteCount(label)
            }
        ];
    }
}
