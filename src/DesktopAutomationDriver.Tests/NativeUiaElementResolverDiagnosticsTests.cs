using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaElementResolverDiagnosticsTests
{
    [Theory]
    [InlineData("activeWindow", "activeWindow")]
    [InlineData("processWindows", "processWindows")]
    [InlineData("desktopChildren", "desktopChildren")]
    [InlineData("desktop", "desktopChildren")]
    public void ResolveRootMode_NormalizesRequestRoot(string root, string expected)
    {
        var mode = NativeUiaElementResolver.ResolveRootMode(new UiRequest { Root = root });
        Assert.Equal(expected, mode);
    }

    [Fact]
    public void ResolveRootMode_UsesSearchRootWhenRootMissing()
    {
        var mode = NativeUiaElementResolver.ResolveRootMode(new UiRequest { SearchRoot = "process-windows" });
        Assert.Equal("processWindows", mode);
    }
}
