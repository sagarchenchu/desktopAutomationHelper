using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Fail-closed link-aware containment for planning reads/writes under the workspace.
/// Every path component is resolved recursively so chained directory links cannot escape.
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

        EnsurePathComponentsResolvedInside(normalizedPath, normalizedRoot);

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

        EnsurePathComponentsResolvedInside(resolved, normalizedRoot);
    }

    /// <summary>
    /// Walk every path component from the containment root to the leaf and recursively
    /// resolve directory reparse points (a → b → /outside must fail).
    /// </summary>
    private static void EnsurePathComponentsResolvedInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot)
            && !WorkspaceManager.PathsEqual(normalizedPath, normalizedRoot))
        {
            throw new PlanningPathException(
                $"path '{fullPath}' resolves outside the allowed directory.");
        }

        var relative = Path.GetRelativePath(normalizedRoot, normalizedPath);
        if (relative == ".")
            return;

        var parts = relative.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        var current = normalizedRoot;
        foreach (var part in parts)
        {
            current = Path.GetFullPath(Path.Combine(current, part));
            var directoryInfo = new DirectoryInfo(current);
            if (!directoryInfo.Exists)
                continue;

            if (!IsReparsePoint(directoryInfo))
            {
                if (!IsInsideDirectory(directoryInfo.FullName, normalizedRoot))
                {
                    throw new PlanningPathException(
                        $"path '{fullPath}' traverses '{directoryInfo.FullName}' outside the allowed directory.");
                }

                continue;
            }

            var finalTarget = RequireResolvedLinkTarget(directoryInfo, fullPath);
            if (!IsInsideDirectory(finalTarget, normalizedRoot))
            {
                throw new PlanningPathException(
                    $"path '{fullPath}' traverses directory link '{directoryInfo.FullName}' that resolves outside the allowed directory.");
            }

            // Continue walking on the lexical component so subsequent segments are checked;
            // final-target containment already rejected escapes through the full chain.
            EnsureResolvedDirectoryChainInside(finalTarget, normalizedRoot, fullPath);
        }
    }

    private static void EnsureResolvedDirectoryChainInside(
        string resolvedDirectory,
        string containmentRoot,
        string originalPath)
    {
        var info = new DirectoryInfo(resolvedDirectory);
        if (!info.Exists)
            return;

        if (!IsInsideDirectory(info.FullName, containmentRoot))
        {
            throw new PlanningPathException(
                $"path '{originalPath}' resolves through '{resolvedDirectory}' outside the allowed directory.");
        }

        if (!IsReparsePoint(info))
            return;

        var final = RequireResolvedLinkTarget(info, originalPath);
        if (!IsInsideDirectory(final, containmentRoot))
        {
            throw new PlanningPathException(
                $"path '{originalPath}' resolves through '{resolvedDirectory}' outside the allowed directory via a link chain.");
        }
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

        EnsurePathComponentsResolvedInside(normalized, containmentRoot);

        var directoryInfo = new DirectoryInfo(normalized);
        if (directoryInfo.Exists)
        {
            if (!IsReparsePoint(directoryInfo))
                return normalized;

            var dirFinal = RequireResolvedLinkTarget(directoryInfo, fullPath);
            if (!IsInsideDirectory(dirFinal, containmentRoot))
            {
                throw new PlanningPathException(
                    $"directory link '{fullPath}' resolves outside the allowed directory.");
            }

            return ResolveFinalPath(dirFinal, containmentRoot, visited, depth + 1);
        }

        var fileInfo = new FileInfo(normalized);
        if (!fileInfo.Exists)
            return normalized;

        if (!IsReparsePoint(fileInfo))
            return normalized;

        var fileFinal = RequireResolvedLinkTarget(fileInfo, fullPath);
        if (!IsInsideDirectory(fileFinal, containmentRoot))
        {
            throw new PlanningPathException(
                $"symbolic link '{fullPath}' resolves outside the allowed directory.");
        }

        return ResolveFinalPath(fileFinal, containmentRoot, visited, depth + 1);
    }

    /// <summary>
    /// Fail-closed link resolution: any inability to fully resolve a reparse point rejects the path.
    /// </summary>
    private static string RequireResolvedLinkTarget(FileSystemInfo info, string originalPath)
    {
        try
        {
            var target = info.ResolveLinkTarget(returnFinalTarget: true);
            if (target is null)
            {
                throw new PlanningPathException(
                    $"path '{originalPath}' is a reparse point that cannot be validated safely.");
            }

            return Path.GetFullPath(target.FullName);
        }
        catch (PlanningPathException)
        {
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new PlanningPathException(
                $"path '{originalPath}' is a reparse point that cannot be validated safely.");
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
            throw new PlanningPathException(
                $"path '{info.FullName}' could not be inspected for reparse points.");
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
