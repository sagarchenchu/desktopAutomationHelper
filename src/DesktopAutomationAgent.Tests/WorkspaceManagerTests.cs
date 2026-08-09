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
        Assert.Contains(first.CreatedPaths, p => p.Replace('\\', '/') == "schemas/plan.schema.json");
        Assert.Contains(first.CreatedPaths, p => p.Replace('\\', '/') == "plans/example.plan.json");

        var marker = "custom-user-content";
        var smokePath = Path.Combine(workspace.RootPath, "suites", "smoke.json");
        File.WriteAllText(smokePath, marker);

        var second = workspace.Initialize();
        Assert.Contains(second.SkippedExistingPaths, p => p.Replace('\\', '/') == "suites/smoke.json");
        Assert.Contains(second.SkippedExistingPaths, p => p.Replace('\\', '/') == "schemas/plan.schema.json");
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

    [Fact]
    public void IsInsideRoot_UsesOsAwarePathComparison()
    {
        if (OperatingSystem.IsWindows())
        {
            Assert.True(WorkspaceManager.PathsEqual(@"C:\Workspace", @"c:\workspace"));
        }
        else
        {
            Assert.False(WorkspaceManager.PathsEqual("/workspace/automation", "/workspace/Automation"));

            var options = TestSupport.CreateOptions(workspaceRoot: "/tmp/da-agent-case-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(options.Workspace.Root);
            var workspace = TestSupport.CreateWorkspace(options);

            var alteredCase = options.Workspace.Root.ToUpperInvariant();
            if (!string.Equals(alteredCase, options.Workspace.Root, StringComparison.Ordinal))
            {
                Assert.Throws<WorkspaceException>(() =>
                    workspace.ResolveSafePath(Path.Combine(alteredCase, "suites", "smoke.json")));
            }

            Directory.Delete(options.Workspace.Root, recursive: true);
        }
    }

    [Fact]
    public void Initialize_WrapsIoFailuresAsWorkspaceException()
    {
        var options = TestSupport.CreateOptions(
            workspaceRoot: Path.Combine(Path.GetTempPath(), "da-agent-file-" + Guid.NewGuid().ToString("N") + ".blocker"));
        File.WriteAllText(options.Workspace.Root, "not-a-directory");

        var workspace = TestSupport.CreateWorkspace(options);
        var ex = Assert.Throws<WorkspaceException>(() => workspace.Initialize());
        Assert.Contains("Failed to initialize workspace", ex.Message, StringComparison.OrdinalIgnoreCase);

        File.Delete(options.Workspace.Root);
    }
}
