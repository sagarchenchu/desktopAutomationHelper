using System.Text;
using DesktopAutomationDriver.Services.Assistive;
using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Window-scoped screenshots only (not unrestricted desktop) unless explicitly requested.
/// Avoids persisting sensitive content by default.
/// </summary>
public sealed class VisionScreenshotService
{
    private static readonly string[] SensitiveMarkers =
    [
        "password", "ssn", "social security", "credit card", "cvv", "secret", "api key", "bearer "
    ];

    private readonly VisionFallbackOptions _options;

    public VisionScreenshotService(VisionFallbackOptions? options = null)
    {
        _options = options ?? new VisionFallbackOptions();
    }

    /// <summary>
    /// Creates a window-scoped screenshot envelope from already-captured PNG bytes (testable).
    /// </summary>
    public VisionScreenshot CreateWindowScoped(
        byte[] pngBytes,
        DebugWindowIdentity window,
        string? ocrOrUiTextHint = null,
        string? persistDirectory = null,
        bool allowUnrestrictedDesktop = false)
    {
        ArgumentNullException.ThrowIfNull(pngBytes);
        ArgumentNullException.ThrowIfNull(window);

        if (allowUnrestrictedDesktop)
            throw new InvalidOperationException(
                "Unrestricted desktop screenshots are not enabled in Phase 8; window scope is required.");

        var sensitive = DetectSensitive(ocrOrUiTextHint) || DetectSensitive(window.WindowTitle);
        var reason = sensitive ? "Matched sensitive marker in window title or hint text." : null;

        string? path = null;
        var persisted = false;
        if (!sensitive || _options.PersistSensitiveScreenshots)
        {
            if (!string.IsNullOrWhiteSpace(persistDirectory) && (!sensitive || _options.PersistSensitiveScreenshots))
            {
                if (sensitive && !_options.PersistSensitiveScreenshots)
                {
                    // unreachable due to outer condition, kept for clarity
                }
                else if (!sensitive)
                {
                    var root = Path.GetFullPath(persistDirectory);
                    Directory.CreateDirectory(root);
                    AssistivePathSafety.EnsureWritablePathInside(root, root);
                    path = Path.Combine(root, $"vision-{DateTimeOffset.UtcNow:yyyyMMddHHmmssfff}.png");
                    AssistivePathSafety.EnsureWritablePathInside(path, root);
                    File.WriteAllBytes(path, pngBytes);
                    persisted = true;
                }
            }
        }

        if (sensitive && !_options.PersistSensitiveScreenshots)
        {
            persisted = false;
            path = null;
        }

        return new VisionScreenshot
        {
            PngBytes = pngBytes,
            WindowBounds = window.Bounds,
            Window = window,
            CapturedAtUtc = DateTimeOffset.UtcNow,
            ContainsSensitiveContent = sensitive,
            SensitivityReason = reason,
            Persisted = persisted,
            PersistedPath = path,
            Scope = "window"
        };
    }

    public static string? RedactSensitiveText(string? text, bool enabled = true)
    {
        if (!enabled || string.IsNullOrEmpty(text))
            return text;

        var result = text;
        foreach (var marker in SensitiveMarkers)
        {
            // crude redaction of marker words for persisted OCR snippets
            result = ReplaceInsensitive(result, marker, "[REDACTED]");
        }

        return result;
    }

    public static bool DetectSensitive(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;
        foreach (var marker in SensitiveMarkers)
        {
            if (text.Contains(marker, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static string ReplaceInsensitive(string input, string oldValue, string newValue)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue))
            return input;

        var sb = new StringBuilder();
        var index = 0;
        while (index < input.Length)
        {
            var found = input.IndexOf(oldValue, index, StringComparison.OrdinalIgnoreCase);
            if (found < 0)
            {
                sb.Append(input.AsSpan(index));
                break;
            }

            sb.Append(input.AsSpan(index, found - index));
            sb.Append(newValue);
            index = found + oldValue.Length;
        }

        return sb.ToString();
    }
}
