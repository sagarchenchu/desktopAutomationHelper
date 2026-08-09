using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Tests;

public class WorkspaceManagerTests
{
    [Fact]
    public void Initialize_IsIdempotentAndDoesNotOverwrite()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);

        var first = workspace.Initialize();
        Assert.Contains(first.CreatedPaths, p => p.Replace('\\', '/') == "suites/smoke.json");

        var marker = "custom-user-content";
        var smokePath = Path.Combine(workspace.RootPath, "suites", "smoke.json");
        File.WriteAllText(smokePath, marker);

        var second = workspace.Initialize();
        Assert.Contains(second.SkippedExistingPaths, p => p.Replace('\\', '/') == "suites/smoke.json");
        Assert.Equal(marker, File.ReadAllText(smokePath));

        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void ResolveSafePath_RejectsTraversal()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var workspace = TestSupport.CreateWorkspace(options);

        Assert.Throws<WorkspaceException>(() => workspace.ResolveSafePath("../outside.json"));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void ResolveSafePath_RejectsAbsolutePathOutsideRoot()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var workspace = TestSupport.CreateWorkspace(options);
        var outside = Path.Combine(Path.GetTempPath(), "da-agent-outside-" + Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(outside, "{}");

        Assert.Throws<WorkspaceException>(() => workspace.ResolveSafePath(outside));

        File.Delete(outside);
        Directory.Delete(options.Workspace.Root, recursive: true);
    }
}
