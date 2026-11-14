using System.Security.Cryptography;
using System.Text;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Stable fingerprint for approve-time content revalidation.</summary>
public static class VisionFingerprint
{
    public static string Compute(byte[] pngBytes, string? label, Rectangle bounds)
    {
        ArgumentNullException.ThrowIfNull(pngBytes);
        var labelPart = (label ?? string.Empty).Trim();
        var material = Convert.ToHexString(SHA256.HashData(pngBytes))[..16].ToLowerInvariant();
        return $"{material}:{labelPart.Length}:{bounds.X},{bounds.Y},{bounds.Width},{bounds.Height}";
    }

    public static bool Matches(string? expected, string? actual) =>
        !string.IsNullOrWhiteSpace(expected)
        && string.Equals(expected, actual, StringComparison.Ordinal);
}
