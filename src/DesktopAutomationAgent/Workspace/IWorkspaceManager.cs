namespace DesktopAutomationAgent.Workspace;

public interface IWorkspaceManager
{
    string RootPath { get; }

    WorkspaceInitResult Initialize();

    void EnsureInitialized();

    string ResolveSafePath(string relativeOrAbsolutePath);
}

public sealed class WorkspaceInitResult
{
    public required string RootPath { get; init; }

    public IReadOnlyList<string> CreatedPaths { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> SkippedExistingPaths { get; init; } = Array.Empty<string>();
}
