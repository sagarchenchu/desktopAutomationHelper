using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Vision;

/// <summary>
/// Agent-side gate for vision fallback. Disabled by default; never replaces working UIA.
/// </summary>
public sealed class VisionPolicyEvaluator
{
    private readonly VisionOptions _options;

    public VisionPolicyEvaluator(VisionOptions? options = null)
    {
        _options = options ?? new VisionOptions();
    }

    public static VisionPolicyEvaluator FromAgentOptions(AgentOptions options) =>
        new(options.Vision);

    public bool IsEnabled => _options.Enabled;

    public bool MayUseVision(bool uiaInaccessible, bool policyExplicitlyPermits = false)
    {
        if (!_options.Enabled)
            return false;

        if (policyExplicitlyPermits && _options.AllowWhenPolicyPermits)
            return true;

        if (uiaInaccessible && _options.AllowWhenUiaInaccessible)
            return true;

        return false;
    }

    public VisionOptions Options => _options;
}
