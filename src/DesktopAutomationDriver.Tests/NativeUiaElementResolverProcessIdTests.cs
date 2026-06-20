using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaElementResolverProcessIdTests
{
    [Fact]
    public void ToNativeLocator_SessionProcessId_DoesNotEnforceProcessIdMatch()
    {
        var locator = NativeUiaElementResolver.ToNativeLocator(
            new UiRequest
            {
                Operation = "clickmenuuia",
                Locator = new UiLocator
                {
                    Name = "DQA",
                    ControlType = "Menu",
                    MatchMode = "contains"
                }
            },
            4242);

        Assert.Equal(4242, locator.ProcessId);
        Assert.False(locator.EnforceProcessIdMatch);
    }

    [Fact]
    public void ToNativeLocator_ExplicitRequestProcessId_EnforcesProcessIdMatch()
    {
        var locator = NativeUiaElementResolver.ToNativeLocator(
            new UiRequest
            {
                Operation = "clickmenuuia",
                ProcessId = 1111,
                Locator = new UiLocator
                {
                    Name = "DQA",
                    ControlType = "Menu"
                }
            },
            4242);

        Assert.Equal(1111, locator.ProcessId);
        Assert.True(locator.EnforceProcessIdMatch);
    }

    [Fact]
    public void ToNativeLocator_ExplicitLocatorProcessId_EnforcesProcessIdMatch()
    {
        var locator = NativeUiaElementResolver.ToNativeLocator(
            new UiRequest
            {
                Operation = "clickuia",
                Locator = new UiLocator
                {
                    Name = "DQA",
                    ProcessId = 2222
                }
            },
            4242);

        Assert.Equal(2222, locator.ProcessId);
        Assert.True(locator.EnforceProcessIdMatch);
    }
}
