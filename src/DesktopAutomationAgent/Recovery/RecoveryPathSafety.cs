using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Link-aware containment for recovery candidate writes under object-repository/candidates.
/// </summary>
internal static class RecoveryPathSafety
{
    private const int MaxSymlinkDepth = 32;

    public static string CandidatesRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "object-repository", "candidates"));

    public static string PagesRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "object-repository", "pages"));

    public static void EnsureWritableUnderCandidates(string fullPath, string workspaceRoot)
    {
        var candidates = CandidatesRoot(workspaceRoot);
        var pages = PagesRoot(workspaceRoot);
        EnsureWritablePathInside(fullPath, candidates);
        EnsureNotSymlinkEscape(fullPath, candidates);

        var normalized = Path.GetFullPath(fullPath);
        if (IsInsideDirectory(normalized, pages) || WorkspaceManager.PathsEqual(normalized, pages))
        {
            throw new RecoveryPathException(
                $"path '{fullPath}' must not resolve under object-repository/pages.");
        }
    }

    public static void EnsureWritablePathInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot)
            && !WorkspaceManager.PathsEqual(normalizedPath, normalizedRoot))
        {
            throw new RecoveryPathException(
                $"path '{fullPath}' resolves outside the allowed directory.");
        }

        EnsureNotSymlinkEscape(normalizedPath, normalizedRoot);
    }

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

        if (!IsInsideDirectory(resolved, normalizedRoot)
            && !WorkspaceManager.PathsEqual(resolved, normalizedRoot))
        {
            throw new RecoveryPathException(
                $"path '{fullPath}' resolves outside the allowed directory via a link chain.");
        }

        EnsurePathComponentsResolvedInside(resolved, normalizedRoot);
    }

    private static void EnsurePathComponentsResolvedInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot)
            && !WorkspaceManager.PathsEqual(normalizedPath, normalizedRoot))
        {
            throw new RecoveryPathException(
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
                if (!IsInsideDirectory(directoryInfo.FullName, normalizedRoot)
                    && !WorkspaceManager.PathsEqual(directoryInfo.FullName, normalizedRoot))
                {
                    throw new RecoveryPathException(
                        $"path '{fullPath}' traverses '{directoryInfo.FullName}' outside the allowed directory.");
                }

                continue;
            }

            var finalTarget = RequireResolvedLinkTarget(directoryInfo, fullPath);
            if (!IsInsideDirectory(finalTarget, normalizedRoot)
                && !WorkspaceManager.PathsEqual(finalTarget, normalizedRoot))
            {
                throw new RecoveryPathException(
                    $"path '{fullPath}' traverses directory link '{directoryInfo.FullName}' that resolves outside the allowed directory.");
            }
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
            throw new RecoveryPathException(
                $"path '{fullPath}' exceeds the maximum symbolic-link resolution depth.");
        }

        var normalized = Path.GetFullPath(fullPath);
        if (!visited.Add(normalized))
        {
            throw new RecoveryPathException(
                $"path '{fullPath}' participates in a symbolic-link cycle.");
        }

        EnsurePathComponentsResolvedInside(normalized, containmentRoot);

        var directoryInfo = new DirectoryInfo(normalized);
        if (directoryInfo.Exists)
        {
            if (!IsReparsePoint(directoryInfo))
                return normalized;

            var dirFinal = RequireResolvedLinkTarget(directoryInfo, fullPath);
            if (!IsInsideDirectory(dirFinal, containmentRoot)
                && !WorkspaceManager.PathsEqual(dirFinal, containmentRoot))
            {
                throw new RecoveryPathException(
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
        if (!IsInsideDirectory(fileFinal, containmentRoot)
            && !WorkspaceManager.PathsEqual(fileFinal, containmentRoot))
        {
            throw new RecoveryPathException(
                $"symbolic link '{fullPath}' resolves outside the allowed directory.");
        }

        return ResolveFinalPath(fileFinal, containmentRoot, visited, depth + 1);
    }

    private static string RequireResolvedLinkTarget(FileSystemInfo info, string originalPath)
    {
        try
        {
            var target = info.ResolveLinkTarget(returnFinalTarget: true);
            if (target is null)
            {
                throw new RecoveryPathException(
                    $"path '{originalPath}' is a reparse point that cannot be validated safely.");
            }

            return Path.GetFullPath(target.FullName);
        }
        catch (RecoveryPathException)
        {
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new RecoveryPathException(
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
            throw new RecoveryPathException(
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

public sealed class RecoveryPathException : Exception
{
    public RecoveryPathException(string message) : base(message)
    {
    }
}
