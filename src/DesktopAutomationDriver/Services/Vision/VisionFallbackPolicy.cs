using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Vision never replaces working UIA locators. Off by default; gated on inaccessibility or explicit policy.
/// </summary>
public sealed class VisionFallbackPolicy
{
    private readonly VisionFallbackOptions _options;

    public VisionFallbackPolicy(VisionFallbackOptions? options = null)
    {
        _options = options ?? new VisionFallbackOptions();
    }

    public VisionFallbackOptions Options => _options;

    public VisionEligibility Evaluate(DebugUiaNode? semanticTarget, bool policyExplicitlyPermits = false)
    {
        if (!_options.Enabled)
            return VisionEligibility.DisabledByDefault;

        if (policyExplicitlyPermits && _options.AllowWhenPolicyPermits)
            return VisionEligibility.AllowedByPolicy;

        if (_options.AllowWhenUiaInaccessible
            && semanticTarget is not null
            && UiaTreeCaptureService.IsSemanticallyInaccessible(semanticTarget))
        {
            return VisionEligibility.AllowedUiaInaccessible;
        }

        if (semanticTarget is not null && !UiaTreeCaptureService.IsSemanticallyInaccessible(semanticTarget))
            return VisionEligibility.DeniedUiaAccessible;

        return VisionEligibility.Denied;
    }

    public bool IsEnabled => _options.Enabled;
}
