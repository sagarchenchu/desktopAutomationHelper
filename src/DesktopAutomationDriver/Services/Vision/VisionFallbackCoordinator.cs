using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Orchestrates gated vision fallback: policy → window screenshot → OCR candidate → coordinate guard.
/// Never replaces working UIA locators and never auto-promotes.
/// </summary>
public sealed class VisionFallbackCoordinator
{
    private readonly VisionFallbackPolicy _policy;
    private readonly VisionScreenshotService _screenshots;
    private readonly IVisionProvider _provider;
    private readonly CoordinateActionGuard _guard;

    public VisionFallbackCoordinator(
        VisionFallbackPolicy policy,
        VisionScreenshotService screenshots,
        IVisionProvider provider,
        CoordinateActionGuard guard)
    {
        _policy = policy ?? throw new ArgumentNullException(nameof(policy));
        _screenshots = screenshots ?? throw new ArgumentNullException(nameof(screenshots));
        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        _guard = guard ?? throw new ArgumentNullException(nameof(guard));
    }

    public static VisionFallbackCoordinator ForTests(VisionFallbackOptions? options = null)
    {
        options ??= new VisionFallbackOptions { Enabled = false, Provider = "fake" };
        var policy = new VisionFallbackPolicy(options);
        return new VisionFallbackCoordinator(
            policy,
            new VisionScreenshotService(options),
            new FakeVisionProvider(),
            new CoordinateActionGuard(options));
    }

    public VisionEligibility EvaluateEligibility(
        DebugUiaNode? semanticTarget,
        bool policyExplicitlyPermits = false) =>
        _policy.Evaluate(semanticTarget, policyExplicitlyPermits);

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

        var screenshot = _screenshots.CreateWindowScoped(
            windowPngBytes,
            window,
            ocrOrUiTextHint: query,
            persistDirectory: persistDirectory);

        var candidates = _provider.Analyze(screenshot, query)
            .Where(c => !c.AutoPromoted)
            .Where(c => c.Confidence >= _policy.Options.MinimumConfidence)
            .ToList();

        return (eligibility, screenshot, candidates);
    }

    public CoordinateActionDecision AuthorizeCoordinateAction(CoordinateActionRequest request) =>
        _guard.Evaluate(request);
}
