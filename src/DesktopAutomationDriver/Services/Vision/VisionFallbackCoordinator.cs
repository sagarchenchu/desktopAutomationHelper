using System.Security.Cryptography;
using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Gated vision fallback: server capture → OCR candidate → store → human presence
/// → re-capture vs original baseline fingerprint → action-bound one-use token → authorize.
/// </summary>
public sealed class VisionFallbackCoordinator
{
    private readonly VisionFallbackPolicy _policy;
    private readonly VisionScreenshotService _screenshots;
    private readonly IVisionProvider _provider;
    private readonly CoordinateActionGuard _guard;
    private readonly VisionCandidateStore _store;
    private readonly ILiveWindowProbe _windowProbe;
    private readonly IVisionCaptureService _capture;
    private readonly IOperatorPresenceService _presence;

    public VisionFallbackCoordinator(
        VisionFallbackPolicy policy,
        VisionScreenshotService screenshots,
        IVisionProvider provider,
        CoordinateActionGuard guard,
        VisionCandidateStore store,
        ILiveWindowProbe windowProbe,
        IVisionCaptureService capture,
        IOperatorPresenceService presence)
    {
        _policy = policy ?? throw new ArgumentNullException(nameof(policy));
        _screenshots = screenshots ?? throw new ArgumentNullException(nameof(screenshots));
        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        _guard = guard ?? throw new ArgumentNullException(nameof(guard));
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _windowProbe = windowProbe ?? throw new ArgumentNullException(nameof(windowProbe));
        _capture = capture ?? throw new ArgumentNullException(nameof(capture));
        _presence = presence ?? throw new ArgumentNullException(nameof(presence));
    }

    public static VisionFallbackCoordinator ForTests(
        VisionFallbackOptions? options = null,
        VisionCandidateStore? store = null,
        ILiveWindowProbe? windowProbe = null,
        IVisionCaptureService? capture = null,
        IOperatorPresenceService? presence = null)
    {
        options ??= new VisionFallbackOptions { Enabled = false, Provider = "fake" };
        store ??= new VisionCandidateStore(options);
        windowProbe ??= new ScriptedLiveWindowProbe();
        capture ??= new ScriptedVisionCaptureService((_, _) => null);
        presence ??= new ScriptedOperatorPresenceService { Allow = true };
        var policy = new VisionFallbackPolicy(options);
        return new VisionFallbackCoordinator(
            policy,
            new VisionScreenshotService(options),
            new FakeVisionProvider(),
            new CoordinateActionGuard(options),
            store,
            windowProbe,
            capture,
            presence);
    }

    public VisionFallbackOptions Options => _policy.Options;
    public VisionCandidateStore Store => _store;

    public VisionEligibility EvaluateEligibility(
        DebugUiaNode? semanticTarget,
        bool policyExplicitlyPermits = false) =>
        _policy.Evaluate(semanticTarget, policyExplicitlyPermits);

    public VisionEligibility EvaluateEligibility(bool uiaReportsInaccessible, bool policyOverride = false)
    {
        if (!_policy.Options.Enabled)
            return VisionEligibility.DisabledByDefault;

        if (uiaReportsInaccessible && _policy.Options.AllowWhenUiaInaccessible)
            return VisionEligibility.AllowedUiaInaccessible;

        if (policyOverride && _policy.Options.AllowWhenPolicyPermits)
            return VisionEligibility.AllowedByPolicy;

        if (!uiaReportsInaccessible)
            return VisionEligibility.DeniedUiaAccessible;

        return VisionEligibility.Denied;
    }

    public (VisionEligibility Eligibility, VisionScreenshot? Screenshot, IReadOnlyList<VisionCandidateTarget> Candidates)
        TryProduceCandidatesFromHandle(
            IntPtr windowHandle,
            Point? cursorScreenPoint = null,
            string? query = null,
            bool policyExplicitlyPermits = false,
            string? persistDirectory = null)
    {
        if (!_policy.Options.Enabled)
            return (VisionEligibility.DisabledByDefault, null, []);

        var capture = _capture.Capture(windowHandle, cursorScreenPoint);
        if (capture is null)
            return (VisionEligibility.Denied, null, []);

        return TryProduceCandidates(
            capture.SemanticTarget,
            capture.Window,
            capture.PngBytes,
            query: query,
            policyExplicitlyPermits: policyExplicitlyPermits,
            persistDirectory: persistDirectory,
            cursorScreenPoint: capture.Cursor);
    }

    public (VisionEligibility Eligibility, VisionScreenshot? Screenshot, IReadOnlyList<VisionCandidateTarget> Candidates)
        TryProduceCandidates(
            DebugUiaNode? semanticTarget,
            DebugWindowIdentity window,
            byte[] windowPngBytes,
            string? query = null,
            bool policyExplicitlyPermits = false,
            string? persistDirectory = null,
            Point? cursorScreenPoint = null)
    {
        var eligibility = _policy.Evaluate(semanticTarget, policyExplicitlyPermits);
        if (eligibility is VisionEligibility.DisabledByDefault
            or VisionEligibility.Denied
            or VisionEligibility.DeniedUiaAccessible)
        {
            return (eligibility, null, []);
        }

        var issuedAt = DateTimeOffset.UtcNow;
        var screenshot = _screenshots.CreateWindowScoped(
            windowPngBytes,
            window,
            ocrOrUiTextHint: query,
            persistDirectory: persistDirectory);

        var issued = _provider.Analyze(screenshot, query)
            .Where(c => !c.AutoPromoted)
            .Where(c => c.Confidence >= _policy.Options.MinimumConfidence)
            .Select(c =>
            {
                var fingerprint = VisionFingerprint.Compute(windowPngBytes, c.Label, c.Bounds);
                return _store.Issue(
                    c with
                    {
                        ProducedAtUtc = issuedAt,
                        AutoPromoted = false,
                        Window = window,
                        ScreenshotFingerprint = fingerprint
                    },
                    issuedAt,
                    cursorScreenPoint).Snapshot;
            })
            .ToList();

        return (eligibility, screenshot, issued);
    }

    /// <summary>
    /// Operator approval: requires driver-owned human presence, then re-captures pixels and
    /// revalidates fingerprint/label/bounds against the <em>original</em> window baseline
    /// (not live-vs-live), then issues a short-lived one-use execution token.
    /// </summary>
    public VisionApprovalResult ApproveCandidate(string candidateId, string action)
    {
        if (!_policy.Options.Enabled)
            return VisionApprovalResult.Deny("Vision fallback disabled (Vision:Enabled=false).");

        if (string.IsNullOrWhiteSpace(candidateId) || !_store.TryGet(candidateId, out var issued))
            return VisionApprovalResult.Deny("Unknown or expired vision candidate id.");

        if (issued.AutoPromoted || issued.Snapshot.AutoPromoted)
            return VisionApprovalResult.Deny("Auto-promoted vision candidates are denied.");

        if (issued.Consumed)
            return VisionApprovalResult.Deny("Vision candidate already consumed.");

        var normalizedAction = VisionCandidateStore.NormalizeAction(action);
        if (DestructiveActionHeuristics.IsDestructive(normalizedAction, issued.Snapshot.Label, false))
            return VisionApprovalResult.Deny("Destructive or unknown actions cannot be approved.");

        // Genuine human presence — ACL file self-read is not sufficient.
        if (!_presence.TryConfirmVisionApproval(
                new OperatorPresenceRequest(issued.Id, normalizedAction, issued.Snapshot.Label),
                out var presenceDeny))
        {
            return VisionApprovalResult.Deny(
                string.IsNullOrWhiteSpace(presenceDeny)
                    ? "Operator presence confirmation required."
                    : presenceDeny);
        }

        // Re-capture before issuing token — do not trust stale HWND-only identity.
        var recapture = _capture.Capture(
            issued.Snapshot.Window.WindowHandle,
            issued.CursorScreenPoint);
        if (recapture is null)
            return VisionApprovalResult.Deny("Approval re-capture failed (window/cursor binding).");

        var baseline = issued.Snapshot;
        var live = recapture.Window;
        var freshFingerprint = VisionFingerprint.Compute(
            recapture.PngBytes,
            baseline.Label,
            baseline.Bounds);
        if (!VisionFingerprint.Matches(baseline.ScreenshotFingerprint, freshFingerprint))
            return VisionApprovalResult.Deny("Approval revalidation failed: content fingerprint changed.");

        // Compare live probe against the ORIGINAL window identity/DPI/layout — not live-vs-live.
        var revalidation = _guard.EvaluateForApproval(new CoordinateActionRequest
        {
            Candidate = baseline,
            CurrentWindow = live,
            NowUtc = DateTimeOffset.UtcNow,
            Action = normalizedAction
        });
        if (!revalidation.Allowed)
            return VisionApprovalResult.Deny($"Approval revalidation failed: {revalidation.Reason}");

        var token = "vtx-" + Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();
        var approvedAt = DateTimeOffset.UtcNow;

        // Keep matched fingerprint; refresh live window only after original-baseline checks pass.
        issued.Snapshot = baseline with
        {
            Window = live,
            ScreenshotFingerprint = freshFingerprint
        };

        if (!_store.TryBindApproval(candidateId, normalizedAction, token, approvedAt, out var bound) || bound is null)
            return VisionApprovalResult.Deny("Failed to bind approval / execution token.");

        return new VisionApprovalResult
        {
            Allowed = true,
            Reason = "Approved.",
            CandidateId = bound.Id,
            Action = bound.ApprovedAction,
            ExecutionToken = bound.ExecutionToken,
            ExecutionTokenExpiresAtUtc = bound.ExecutionTokenExpiresAtUtc,
            ScreenPoint = revalidation.ScreenPoint
        };
    }

    /// <summary>
    /// Authorize with a short-lived one-use execution token. Validates action/match first
    /// (wrong action must not burn the token), then atomically claims before Allowed=true.
    /// Concurrent double-pass: only one claim wins; consume failure ⇒ Denied.
    /// </summary>
    public CoordinateActionDecision AuthorizeCoordinateAction(
        string executionToken,
        string? action = null,
        bool isDestructive = false)
    {
        if (!_policy.Options.Enabled)
            return Deny("Vision fallback disabled (Vision:Enabled=false).");

        if (string.IsNullOrWhiteSpace(executionToken)
            || !_store.TryGetExecutionGrant(executionToken, out var grant))
        {
            return Deny("Unknown, expired, consumed, or mismatched execution token.");
        }

        if (!_store.TryGet(grant.CandidateId, out var issued))
            return Deny("Vision candidate missing for execution token.");

        if (issued.AutoPromoted || issued.Snapshot.AutoPromoted)
            return Deny("Auto-promoted vision candidates are denied.");

        var requestedAction = VisionCandidateStore.NormalizeAction(action ?? grant.Action);
        if (!string.Equals(requestedAction, grant.Action, StringComparison.Ordinal))
            return Deny("Action does not match operator-approved action.");

        // Verify submitted token still equals issued token on the candidate.
        if (!string.Equals(issued.ExecutionToken, executionToken, StringComparison.Ordinal)
            || !string.Equals(grant.IssuedToken, executionToken, StringComparison.Ordinal))
        {
            return Deny("Execution token does not match issued token.");
        }

        var live = _windowProbe.ProbeByHandle(issued.Snapshot.Window.WindowHandle);
        if (live is null)
            return Deny("Fresh window probe failed (HWND/PID/bounds/DPI unavailable).");

        var now = DateTimeOffset.UtcNow;
        var decision = _guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = issued.Snapshot,
            CurrentWindow = live,
            NowUtc = now,
            Action = requestedAction,
            IsDestructive = isDestructive,
            FreshnessAnchorUtc = issued.ApprovedAtUtc ?? issued.Snapshot.ProducedAtUtc
        });

        if (!decision.Allowed)
            return decision;

        // Atomic one-use claim — concurrent callers: only one TryRemove wins.
        if (!_store.TryClaimExecutionToken(executionToken, out _))
            return Deny("Execution token already claimed or consume race lost.");

        if (!_store.TryConsumeAfterClaim(issued.Id, executionToken))
            return Deny("Failed to consume execution token after authorization.");

        return decision;
    }

    public CoordinateActionDecision AuthorizeCoordinateActionByCandidateId(
        string candidateId,
        string action = "click",
        bool isDestructive = false)
    {
        if (!_store.TryGet(candidateId, out var issued)
            || string.IsNullOrWhiteSpace(issued.ExecutionToken))
        {
            return Deny("Unknown candidate or missing execution token; operator approve required.");
        }

        return AuthorizeCoordinateAction(issued.ExecutionToken, action, isDestructive);
    }

    public CoordinateActionDecision AuthorizeGatedPointerAction(
        string executionToken,
        string action,
        bool isDestructive = false) =>
        AuthorizeCoordinateAction(executionToken, action, isDestructive);

    private static CoordinateActionDecision Deny(string reason) =>
        new() { Allowed = false, Reason = reason, ScreenPoint = null };
}

public sealed class VisionApprovalResult
{
    public bool Allowed { get; init; }
    public string Reason { get; init; } = string.Empty;
    public string? CandidateId { get; init; }
    public string? Action { get; init; }
    public string? ExecutionToken { get; init; }
    public DateTimeOffset? ExecutionTokenExpiresAtUtc { get; init; }
    public Point? ScreenPoint { get; init; }

    public static VisionApprovalResult Deny(string reason) =>
        new() { Allowed = false, Reason = reason };
}

public sealed class ScriptedLiveWindowProbe : ILiveWindowProbe
{
    private readonly Func<IntPtr, DebugWindowIdentity?> _probe;

    public ScriptedLiveWindowProbe(Func<IntPtr, DebugWindowIdentity?>? probe = null)
    {
        _probe = probe ?? (_ => null);
    }

    public ScriptedLiveWindowProbe(DebugWindowIdentity identity)
        : this(hwnd => hwnd == identity.WindowHandle ? identity : null)
    {
    }

    public DebugWindowIdentity? ProbeByHandle(IntPtr windowHandle) => _probe(windowHandle);
}
