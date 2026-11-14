using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Vision;

namespace DesktopAutomationAgent.Tests;

public class Phase8VisionPolicyTests
{
    [Fact]
    public void VisionOptions_DefaultDisabled()
    {
        var options = new AgentOptions();
        Assert.False(options.Vision.Enabled);
        Assert.False(options.Vision.PersistSensitiveScreenshots);
        Assert.Equal("fake", options.Vision.Provider);
    }

    [Fact]
    public void VisionPolicyEvaluator_RespectsGates()
    {
        var disabled = new VisionPolicyEvaluator();
        Assert.False(disabled.MayUseVision(uiaInaccessible: true));

        var enabled = new VisionPolicyEvaluator(new VisionOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true,
            AllowWhenPolicyPermits = true
        });
        Assert.True(enabled.MayUseVision(uiaInaccessible: true));
        Assert.False(enabled.MayUseVision(uiaInaccessible: false));
        Assert.True(enabled.MayUseVision(uiaInaccessible: false, policyExplicitlyPermits: true));
    }

    [Fact]
    public void VisionOptionsValidator_RejectsNonFakeProvider()
    {
        var vision = new VisionOptions { Provider = "azure-ocr" };
        var ex = Assert.Throws<AgentConfigurationException>(() =>
            AgentOptionsValidator.ValidateVisionOptions(vision));
        Assert.Contains("fake", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HoverProbeRules_AreNotUiaCapability()
    {
        Assert.False(HoverProbeRules.IsUiaCapability);
        Assert.Equal("hover-noop", HoverProbeRules.DescribeOutcome(false, false, null, true));
        Assert.Equal("hover-tooltip:Hint", HoverProbeRules.DescribeOutcome(false, false, "Hint", false));
        Assert.Equal("hover-structure-changed", HoverProbeRules.DescribeOutcome(true, false, null, false));
    }
}
