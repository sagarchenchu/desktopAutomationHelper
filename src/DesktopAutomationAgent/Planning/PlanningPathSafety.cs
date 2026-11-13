using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Fail-closed link-aware containment for planning reads/writes under the workspace.
/// </summary>
internal static class PlanningPathSafety
{
    private const int MaxSymlinkDepth = 32;

    public static string PlanningRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "planning"));

    public static void EnsureWritablePathInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot))
        {
            throw new PlanningPathException(
                $"path '{fullPath}' resolves outside the allowed directory.");
        }

        EnsureNotSymlinkEscape(normalizedPath, normalizedRoot);
    }

    public static void EnsureReadablePathInside(string fullPath, string containmentRoot) =>
        EnsureWritablePathInside(fullPath, containmentRoot);

    public static void EnsureNotSymlinkEscape(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        EnsureAncestorDirectoriesSafe(normalizedPath, normalizedRoot);

        var resolved = ResolveFinalPath(
            normalizedPath,
            normalizedRoot,
            new HashSet<string>(PathComparer),
            depth: 0);

        if (!IsInsideDirectory(resolved, normalizedRoot))
        {
            throw new PlanningPathException(
                $"path '{fullPath}' resolves outside the allowed directory via a link chain.");
        }

        EnsureAncestorDirectoriesSafe(resolved, normalizedRoot);
    }

    private static string ResolveFinalPath(
        string fullPath,
        string containmentRoot,
        HashSet<string> visited,
        int depth)
    {
        if (depth > MaxSymlinkDepth)
        {
            throw new PlanningPathException(
                $"path '{fullPath}' exceeds the maximum symbolic-link resolution depth.");
        }

        var normalized = Path.GetFullPath(fullPath);
        if (!visited.Add(normalized))
        {
            throw new PlanningPathException(
                $"path '{fullPath}' participates in a symbolic-link cycle.");
        }

        EnsureAncestorDirectoriesSafe(normalized, containmentRoot);

        var directoryInfo = new DirectoryInfo(normalized);
        if (directoryInfo.Exists)
        {
            if (!IsReparsePoint(directoryInfo))
                return normalized;

            var dirTarget = TryGetLinkTarget(directoryInfo);
            if (string.IsNullOrWhiteSpace(dirTarget))
            {
                throw new PlanningPathException(
                    $"path '{fullPath}' is a directory reparse point that cannot be validated safely.");
            }

            var dirNext = Path.GetFullPath(
                Path.IsPathRooted(dirTarget)
                    ? dirTarget
                    : Path.Combine(directoryInfo.Parent?.FullName ?? normalized, dirTarget));

            if (!IsInsideDirectory(dirNext, containmentRoot))
            {
                throw new PlanningPathException(
                    $"directory link '{fullPath}' resolves outside the allowed directory.");
            }

            return ResolveFinalPath(dirNext, containmentRoot, visited, depth + 1);
        }

        var fileInfo = new FileInfo(normalized);
        if (!fileInfo.Exists)
            return normalized;

        if (!IsReparsePoint(fileInfo))
            return normalized;

        var linkTarget = TryGetLinkTarget(fileInfo);
        if (string.IsNullOrWhiteSpace(linkTarget))
        {
            throw new PlanningPathException(
                $"path '{fullPath}' is a reparse point that cannot be validated safely.");
        }

        var next = Path.GetFullPath(
            Path.IsPathRooted(linkTarget)
                ? linkTarget
                : Path.Combine(fileInfo.Directory!.FullName, linkTarget));

        if (!IsInsideDirectory(next, containmentRoot))
        {
            throw new PlanningPathException(
                $"symbolic link '{fullPath}' resolves outside the allowed directory.");
        }

        return ResolveFinalPath(next, containmentRoot, visited, depth + 1);
    }

    private static void EnsureAncestorDirectoriesSafe(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var current = new DirectoryInfo(Path.GetDirectoryName(fullPath) ?? normalizedRoot);

        while (true)
        {
            if (current.Exists && IsReparsePoint(current))
            {
                var linkTarget = TryGetLinkTarget(current);
                if (string.IsNullOrWhiteSpace(linkTarget))
                {
                    throw new PlanningPathException(
                        $"path '{fullPath}' traverses a directory reparse point '{current.FullName}' that cannot be validated safely.");
                }

                var resolved = Path.GetFullPath(
                    Path.IsPathRooted(linkTarget)
                        ? linkTarget
                        : Path.Combine(current.Parent?.FullName ?? normalizedRoot, linkTarget));

                if (!IsInsideDirectory(resolved, containmentRoot))
                {
                    throw new PlanningPathException(
                        $"path '{fullPath}' traverses directory link '{current.FullName}' that resolves outside the allowed directory.");
                }
            }

            if (WorkspaceManager.PathsEqual(current.FullName, normalizedRoot))
                break;

            if (current.Parent is null)
                break;

            current = current.Parent;
        }
    }

    private static bool IsReparsePoint(FileSystemInfo info)
    {
        if (!info.Exists)
            return false;

        try
        {
            if (info.Attributes.HasFlag(FileAttributes.ReparsePoint))
                return true;

            return info.LinkTarget is not null;
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException or IOException)
        {
            return false;
        }
    }

    private static string? TryGetLinkTarget(FileSystemInfo info)
    {
        try
        {
            return info.LinkTarget;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    private static bool IsInsideDirectory(string fullPath, string directoryRoot)
    {
        var root = Path.GetFullPath(directoryRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (WorkspaceManager.PathsEqual(normalized, root))
            return true;

        var relative = Path.GetRelativePath(root, normalized);
        return !Path.IsPathRooted(relative)
               && relative != ".."
               && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
               && !relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal);
    }

    private static StringComparer PathComparer =>
        OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
}

public sealed class PlanningPathException : Exception
{
    public PlanningPathException(string message) : base(message)
    {
    }
}
