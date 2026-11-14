using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Orchestrates gated vision fallback: policy → window screenshot → OCR candidate → store → coordinate guard.
/// Never replaces working UIA locators and never auto-promotes.
/// Authorization requires Vision:Enabled, opaque server-issued id, human approval, and a fresh window probe.
/// </summary>
public sealed class VisionFallbackCoordinator
{
    private readonly VisionFallbackPolicy _policy;
    private readonly VisionScreenshotService _screenshots;
    private readonly IVisionProvider _provider;
    private readonly CoordinateActionGuard _guard;
    private readonly VisionCandidateStore _store;
    private readonly ILiveWindowProbe _windowProbe;

    public VisionFallbackCoordinator(
        VisionFallbackPolicy policy,
        VisionScreenshotService screenshots,
        IVisionProvider provider,
        CoordinateActionGuard guard,
        VisionCandidateStore store,
        ILiveWindowProbe windowProbe)
    {
        _policy = policy ?? throw new ArgumentNullException(nameof(policy));
        _screenshots = screenshots ?? throw new ArgumentNullException(nameof(screenshots));
        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        _guard = guard ?? throw new ArgumentNullException(nameof(guard));
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _windowProbe = windowProbe ?? throw new ArgumentNullException(nameof(windowProbe));
    }

    public static VisionFallbackCoordinator ForTests(
        VisionFallbackOptions? options = null,
        VisionCandidateStore? store = null,
        ILiveWindowProbe? windowProbe = null)
    {
        options ??= new VisionFallbackOptions { Enabled = false, Provider = "fake" };
        store ??= new VisionCandidateStore();
        windowProbe ??= new ScriptedLiveWindowProbe();
        var policy = new VisionFallbackPolicy(options);
        return new VisionFallbackCoordinator(
            policy,
            new VisionScreenshotService(options),
            new FakeVisionProvider(),
            new CoordinateActionGuard(options),
            store,
            windowProbe);
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
    /// Produce vision candidates only when policy allows. Issues opaque server ids (unapproved).
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
                c with { ProducedAtUtc = issuedAt, AutoPromoted = false },
                issuedAt).Snapshot)
            .ToList();

        return (eligibility, screenshot, issued);
    }

    public bool TryApproveCandidate(string candidateId, out VisionIssuedCandidate? issued) =>
        _store.TryApprove(candidateId, out issued);

    /// <summary>
    /// Authorize by opaque server-issued candidate id. Uses server clock + fresh window probe.
    /// Gated on Vision:Enabled. Rejects unknown, unapproved, and auto-promoted candidates.
    /// </summary>
    public CoordinateActionDecision AuthorizeCoordinateAction(
        string candidateId,
        string action = "click",
        bool isDestructive = false)
    {
        if (!_policy.Options.Enabled)
        {
            return Deny("Vision fallback disabled (Vision:Enabled=false).");
        }

        if (string.IsNullOrWhiteSpace(candidateId) || !_store.TryGet(candidateId, out var issued))
            return Deny("Unknown or unissued vision candidate id.");

        if (issued.AutoPromoted || issued.Snapshot.AutoPromoted)
            return Deny("Auto-promoted vision candidates are denied.");

        if (!issued.Approved)
            return Deny("Vision candidate is not approved; human approval required.");

        var live = _windowProbe.ProbeByHandle(issued.Snapshot.Window.WindowHandle);
        if (live is null)
            return Deny("Fresh window probe failed (HWND/PID/bounds/DPI unavailable).");

        // Server clock only — never accept client NowUtc / ProducedAtUtc.
        return _guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = issued.Snapshot,
            CurrentWindow = live,
            NowUtc = DateTimeOffset.UtcNow,
            Action = action,
            IsDestructive = isDestructive
        });
    }

    /// <summary>Legacy overload retained for unit tests of the guard path with an explicit probe window.</summary>
    public CoordinateActionDecision AuthorizeCoordinateAction(CoordinateActionRequest request)
    {
        if (!_policy.Options.Enabled)
            return Deny("Vision fallback disabled (Vision:Enabled=false).");

        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.Candidate.Id)
            || !_store.TryGet(request.Candidate.Id, out var issued))
        {
            return Deny("Unknown or unissued vision candidate id.");
        }

        if (!issued.Approved)
            return Deny("Vision candidate is not approved; human approval required.");

        if (issued.AutoPromoted || request.Candidate.AutoPromoted)
            return Deny("Auto-promoted vision candidates are denied.");

        // Re-probe live window from stored HWND — ignore caller CurrentWindow / NowUtc.
        return AuthorizeCoordinateAction(issued.Id, request.Action, request.IsDestructive);
    }

    /// <summary>
    /// Hover / double-click share the same fail-closed authorize-by-id path.
    /// </summary>
    public CoordinateActionDecision AuthorizeGatedPointerAction(
        string candidateId,
        string action,
        bool isDestructive = false) =>
        AuthorizeCoordinateAction(candidateId, action, isDestructive);

    private static CoordinateActionDecision Deny(string reason) =>
        new() { Allowed = false, Reason = reason, ScreenPoint = null };
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
