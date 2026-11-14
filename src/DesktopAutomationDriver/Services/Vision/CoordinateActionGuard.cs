using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Fail-closed authorization for coordinate/vision actions.
/// Requires click point (and preferably full candidate bounds) inside the live window;
/// requires HWND + PID identity; denies destructive actions and destructive candidate labels.
/// </summary>
public sealed class CoordinateActionGuard
{
    private readonly VisionFallbackOptions _options;

    public CoordinateActionGuard(VisionFallbackOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    public CoordinateActionDecision Evaluate(CoordinateActionRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Candidate);
        ArgumentNullException.ThrowIfNull(request.CurrentWindow);

        if (request.Candidate.AutoPromoted)
            return Deny("Auto-promoted vision candidates are denied.");

        if (DestructiveActionHeuristics.IsDestructive(
                request.Action,
                request.Candidate.Label,
                request.IsDestructive))
        {
            return Deny("Destructive or unknown coordinate actions and destructive candidate labels are denied.");
        }

        // Age is measured from FreshnessAnchorUtc (token/revalidation time), not original issue,
        // so humans are not forced through inspect+approve+authorize inside MaxScreenshotAgeMs alone.
        if (!request.SkipAgeCheck)
        {
            var anchor = request.FreshnessAnchorUtc ?? request.Candidate.ProducedAtUtc;
            var ageMs = (request.NowUtc - anchor).TotalMilliseconds;
            if (ageMs < 0 || ageMs > _options.MaxScreenshotAgeMilliseconds)
                return Deny($"Stale vision candidate (age {ageMs:0}ms > {_options.MaxScreenshotAgeMilliseconds}ms).");
        }

        if (request.Candidate.Confidence < _options.MinimumConfidence)
            return Deny($"Confidence {request.Candidate.Confidence:0.###} below minimum {_options.MinimumConfidence:0.###}.");

        if (!WindowIdentityMatches(request.Candidate.Window, request.CurrentWindow))
            return Deny("Window identity mismatch (HWND and ProcessId required and must match).");

        if (!DpiMatches(request.Candidate.Window, request.CurrentWindow))
            return Deny("DPI scale changed since candidate capture.");

        if (!LayoutMatches(request.Candidate.Window, request.CurrentWindow))
            return Deny("Window layout/bounds changed since candidate capture.");

        var currentBounds = request.CurrentWindow.Bounds;
        if (currentBounds.IsEmpty || currentBounds.Width <= 0 || currentBounds.Height <= 0)
            return Deny("Current window bounds are unavailable.");

        var candidateBounds = request.Candidate.Bounds;
        if (candidateBounds.IsEmpty || candidateBounds.Width <= 0 || candidateBounds.Height <= 0)
            return Deny("Candidate bounds are empty.");

        // Prefer full containment of candidate bounds; always require the click point inside the window.
        if (!currentBounds.Contains(candidateBounds))
            return Deny("Candidate bounds are not fully inside the current window.");

        var center = new Point(
            candidateBounds.X + candidateBounds.Width / 2,
            candidateBounds.Y + candidateBounds.Height / 2);
        if (!currentBounds.Contains(center))
            return Deny("Click point is outside the current window bounds.");

        return new CoordinateActionDecision
        {
            Allowed = true,
            Reason = "Authorized.",
            ScreenPoint = center
        };
    }

    /// <summary>Revalidation at operator approve time — skips screenshot-age so humans have CandidateTtl.</summary>
    public CoordinateActionDecision EvaluateForApproval(CoordinateActionRequest request) =>
        Evaluate(request with { SkipAgeCheck = true });

    private static CoordinateActionDecision Deny(string reason) =>
        new() { Allowed = false, Reason = reason, ScreenPoint = null };

    /// <summary>
    /// Fail closed: both HWND and ProcessId must be present (non-zero) on candidate and current, and match.
    /// </summary>
    internal static bool WindowIdentityMatches(DebugWindowIdentity a, DebugWindowIdentity b)
    {
        if (a.WindowHandle == IntPtr.Zero || b.WindowHandle == IntPtr.Zero)
            return false;
        if (a.WindowHandle != b.WindowHandle)
            return false;

        if (a.ProcessId is null or <= 0 || b.ProcessId is null or <= 0)
            return false;
        if (a.ProcessId != b.ProcessId)
            return false;

        return true;
    }

    private static bool DpiMatches(DebugWindowIdentity a, DebugWindowIdentity b) =>
        NearlyEqual(a.DpiScaleX, b.DpiScaleX) && NearlyEqual(a.DpiScaleY, b.DpiScaleY);

    private static bool LayoutMatches(DebugWindowIdentity a, DebugWindowIdentity b)
    {
        const int tolerance = 2;
        return Math.Abs(a.Bounds.X - b.Bounds.X) <= tolerance
               && Math.Abs(a.Bounds.Y - b.Bounds.Y) <= tolerance
               && Math.Abs(a.Bounds.Width - b.Bounds.Width) <= tolerance
               && Math.Abs(a.Bounds.Height - b.Bounds.Height) <= tolerance;
    }

    private static bool NearlyEqual(double x, double y) => Math.Abs(x - y) < 0.01;
}
