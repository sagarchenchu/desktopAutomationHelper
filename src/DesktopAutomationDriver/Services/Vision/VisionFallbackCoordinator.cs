using System.Security.Cryptography;
using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Gated vision fallback: server capture → OCR candidate → store → operator approve (action-bound
/// one-use execution token) → authorize with fresh probe. Automation cannot self-approve.
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

    public VisionFallbackCoordinator(
        VisionFallbackPolicy policy,
        VisionScreenshotService screenshots,
        IVisionProvider provider,
        CoordinateActionGuard guard,
        VisionCandidateStore store,
        ILiveWindowProbe windowProbe,
        IVisionCaptureService capture)
    {
        _policy = policy ?? throw new ArgumentNullException(nameof(policy));
        _screenshots = screenshots ?? throw new ArgumentNullException(nameof(screenshots));
        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        _guard = guard ?? throw new ArgumentNullException(nameof(guard));
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _windowProbe = windowProbe ?? throw new ArgumentNullException(nameof(windowProbe));
        _capture = capture ?? throw new ArgumentNullException(nameof(capture));
    }

    public static VisionFallbackCoordinator ForTests(
        VisionFallbackOptions? options = null,
        VisionCandidateStore? store = null,
        ILiveWindowProbe? windowProbe = null,
        IVisionCaptureService? capture = null)
    {
        options ??= new VisionFallbackOptions { Enabled = false, Provider = "fake" };
        store ??= new VisionCandidateStore(options);
        windowProbe ??= new ScriptedLiveWindowProbe();
        capture ??= new ScriptedVisionCaptureService((_, _) => null);
        var policy = new VisionFallbackPolicy(options);
        return new VisionFallbackCoordinator(
            policy,
            new VisionScreenshotService(options),
            new FakeVisionProvider(),
            new CoordinateActionGuard(options),
            store,
            windowProbe,
            capture);
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

    /// <summary>
    /// Server-side issuance from HWND (+ optional cursor). Client image/window/UIA flags are not trusted.
    /// </summary>
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
            persistDirectory: persistDirectory);
    }

    /// <summary>
    /// Internal/test path when window + bytes are already server-captured (e.g. Debug session UIA).
    /// </summary>
    public (VisionEligibility Eligibility, VisionScreenshot? Screenshot, IReadOnlyList<VisionCandidateTarget> Candidates)
        TryProduceCandidates(
            DebugUiaNode? semanticTarget,
            DebugWindowIdentity window,
            byte[] windowPngBytes,
            string? query = null,
            bool policyExplicitlyPermits = false,
            string? persistDirectory = null)
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
            .Select(c => _store.Issue(
                c with
                {
                    ProducedAtUtc = issuedAt,
                    AutoPromoted = false,
                    Window = window
                },
                issuedAt).Snapshot)
            .ToList();

        return (eligibility, screenshot, issued);
    }

    /// <summary>
    /// Operator approval: bind action, re-capture/revalidate window, issue short-lived one-use execution token.
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

        var live = _windowProbe.ProbeByHandle(issued.Snapshot.Window.WindowHandle);
        if (live is null)
            return VisionApprovalResult.Deny("Fresh window probe failed during approval revalidation.");

        var revalidation = _guard.EvaluateForApproval(new CoordinateActionRequest
        {
            Candidate = issued.Snapshot,
            CurrentWindow = live,
            NowUtc = DateTimeOffset.UtcNow,
            Action = normalizedAction
        });
        if (!revalidation.Allowed)
            return VisionApprovalResult.Deny($"Approval revalidation failed: {revalidation.Reason}");

        var token = "vtx-" + Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();
        var approvedAt = DateTimeOffset.UtcNow;
        if (!_store.TryBindApproval(candidateId, normalizedAction, token, approvedAt, out var bound) || bound is null)
            return VisionApprovalResult.Deny("Failed to bind approval / execution token.");

        // Refresh snapshot window to the revalidated live probe for subsequent authorize.
        bound.Snapshot = bound.Snapshot with { Window = live, ProducedAtUtc = approvedAt };

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
    /// Authorize with a short-lived one-use execution token (action-bound). Consumes on success.
    /// </summary>
    public CoordinateActionDecision AuthorizeCoordinateAction(
        string executionToken,
        string? action = null,
        bool isDestructive = false)
    {
        if (!_policy.Options.Enabled)
            return Deny("Vision fallback disabled (Vision:Enabled=false).");

        if (string.IsNullOrWhiteSpace(executionToken)
            || !_store.TryGetByExecutionToken(executionToken, out var issued))
        {
            return Deny("Unknown, expired, or unissued execution token.");
        }

        if (issued.AutoPromoted || issued.Snapshot.AutoPromoted)
            return Deny("Auto-promoted vision candidates are denied.");

        if (issued.Consumed || issued.ExecutionTokenConsumed)
            return Deny("Execution token already consumed.");

        if (issued.ExecutionTokenExpiresAtUtc is { } exp && exp < DateTimeOffset.UtcNow)
            return Deny("Execution token expired.");

        if (string.IsNullOrWhiteSpace(issued.ApprovedAction))
            return Deny("Vision candidate is not approved; operator approval required.");

        var requestedAction = VisionCandidateStore.NormalizeAction(action ?? issued.ApprovedAction);
        if (!string.Equals(requestedAction, issued.ApprovedAction, StringComparison.Ordinal))
            return Deny("Action does not match operator-approved action.");

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
            // Freshness from approval/revalidation — not original issue time.
            FreshnessAnchorUtc = issued.ApprovedAtUtc ?? issued.Snapshot.ProducedAtUtc
        });

        if (decision.Allowed)
            _store.TryConsume(issued.Id, executionToken);

        return decision;
    }

    /// <summary>Legacy candidate-id authorize path — requires a live execution token on the candidate.</summary>
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

/// <summary>Test double that echoes a scripted window identity for a given HWND.</summary>
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
