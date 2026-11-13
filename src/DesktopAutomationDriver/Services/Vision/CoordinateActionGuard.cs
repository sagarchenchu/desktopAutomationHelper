using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Revalidates window identity, DPI/scale, bounds, and screenshot freshness before any coordinate action.
/// Never uses a saved coordinate after window/layout change. Destructive actions remain denied.
/// </summary>
public sealed class CoordinateActionGuard
{
    private readonly VisionFallbackOptions _options;

    public CoordinateActionGuard(VisionFallbackOptions? options = null)
    {
        _options = options ?? new VisionFallbackOptions();
    }

    public CoordinateActionDecision Evaluate(CoordinateActionRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Candidate);
        ArgumentNullException.ThrowIfNull(request.CurrentWindow);

        if (request.IsDestructive
            || IsDestructiveAction(request.Action))
        {
            return Deny("Destructive coordinate actions are denied.");
        }

        if (request.Candidate.AutoPromoted)
            return Deny("Vision candidates must never be auto-promoted.");

        if (request.Candidate.Confidence < _options.MinimumConfidence)
        {
            return Deny(
                $"Vision confidence {request.Candidate.Confidence:0.###} is below minimum {_options.MinimumConfidence:0.###}.");
        }

        var ageMs = (request.NowUtc - request.Candidate.ProducedAtUtc).TotalMilliseconds;
        if (ageMs < 0 || ageMs > _options.MaxScreenshotAgeMilliseconds)
        {
            return Deny(
                $"Stale vision screenshot/candidate rejected (age {ageMs:0}ms; max {_options.MaxScreenshotAgeMilliseconds}ms).");
        }

        if (!WindowIdentityMatches(request.Candidate.Window, request.CurrentWindow))
            return Deny("Window identity changed; saved coordinates are invalid.");

        if (!DpiScaleMatches(request.Candidate.Window, request.CurrentWindow))
            return Deny("DPI/scale changed; saved coordinates are invalid.");

        if (!BoundsCompatible(request.Candidate.Window.Bounds, request.CurrentWindow.Bounds))
            return Deny("Window bounds/layout changed; saved coordinates are invalid.");

        if (!request.CurrentWindow.Bounds.Contains(Center(request.Candidate.Bounds))
            && !request.CurrentWindow.Bounds.IntersectsWith(request.Candidate.Bounds))
        {
            return Deny("Candidate bounds fall outside the current window.");
        }

        return new CoordinateActionDecision
        {
            Allowed = true,
            Reason = "Coordinate action revalidated.",
            ScreenPoint = Center(request.Candidate.Bounds)
        };
    }

    public static bool WindowIdentityMatches(DebugWindowIdentity a, DebugWindowIdentity b)
    {
        if (a.WindowHandle != IntPtr.Zero && b.WindowHandle != IntPtr.Zero && a.WindowHandle != b.WindowHandle)
            return false;
        if (a.ProcessId is int pa && b.ProcessId is int pb && pa != pb)
            return false;
        if (!string.IsNullOrWhiteSpace(a.ProcessName)
            && !string.IsNullOrWhiteSpace(b.ProcessName)
            && !string.Equals(a.ProcessName, b.ProcessName, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    public static bool DpiScaleMatches(DebugWindowIdentity a, DebugWindowIdentity b, double tolerance = 0.01)
    {
        return Math.Abs(a.DpiScaleX - b.DpiScaleX) <= tolerance
            && Math.Abs(a.DpiScaleY - b.DpiScaleY) <= tolerance;
    }

    public static bool BoundsCompatible(Rectangle original, Rectangle current, int maxDelta = 2)
    {
        return Math.Abs(original.X - current.X) <= maxDelta
            && Math.Abs(original.Y - current.Y) <= maxDelta
            && Math.Abs(original.Width - current.Width) <= maxDelta
            && Math.Abs(original.Height - current.Height) <= maxDelta;
    }

    public static bool IsDestructiveAction(string? action)
    {
        if (string.IsNullOrWhiteSpace(action))
            return false;
        return action.Equals("delete", StringComparison.OrdinalIgnoreCase)
            || action.Equals("close", StringComparison.OrdinalIgnoreCase)
            || action.Equals("submit", StringComparison.OrdinalIgnoreCase)
            || action.Equals("confirm", StringComparison.OrdinalIgnoreCase)
            || action.Equals("destructive", StringComparison.OrdinalIgnoreCase);
    }

    private static Point Center(Rectangle r) =>
        new(r.X + Math.Max(0, r.Width / 2), r.Y + Math.Max(0, r.Height / 2));

    private static CoordinateActionDecision Deny(string reason) =>
        new() { Allowed = false, Reason = reason, ScreenPoint = null };
}
