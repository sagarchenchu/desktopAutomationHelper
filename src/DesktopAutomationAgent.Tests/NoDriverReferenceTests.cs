using System.Xml.Linq;

namespace DesktopAutomationAgent.Tests;

public class NoDriverReferenceTests
{
    [Fact]
    public void AgentProject_DoesNotReferenceDriver()
    {
        var csproj = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..",
            "DesktopAutomationAgent", "DesktopAutomationAgent.csproj"));

        Assert.True(File.Exists(csproj), csproj);
        var xml = XDocument.Load(csproj);
        var refs = xml.Descendants("ProjectReference")
            .Select(e => (string?)e.Attribute("Include") ?? string.Empty)
            .ToList();

        Assert.DoesNotContain(refs, r => r.Contains("DesktopAutomationDriver", StringComparison.OrdinalIgnoreCase));
    }
}
