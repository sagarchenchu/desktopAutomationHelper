using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Knowledge;

internal static class KnowledgePathSafety
{
    private const int MaxSymlinkDepth = 32;

    public static void EnsureWritablePathInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot))
        {
            throw new KnowledgePathException(
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
            throw new KnowledgePathException(
                $"path '{fullPath}' resolves outside the allowed directory via a link chain.");
        }

        EnsureAncestorDirectoriesSafe(resolved, normalizedRoot);
    }

    public static string SanitizePathSegment(string segment, string label)
    {
        if (string.IsNullOrWhiteSpace(segment))
            throw new KnowledgePathException($"{label} is required.");

        var trimmed = segment.Trim();
        if (trimmed is "." or ".."
            || trimmed.Contains('/')
            || trimmed.Contains('\\')
            || trimmed.Contains(':')
            || trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new KnowledgePathException($"{label} '{segment}' is not a safe path segment.");
        }

        return trimmed;
    }

    public static string CandidatesRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "knowledge", "candidates"));

    public static string ApprovedRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "knowledge", "approved"));

    public static string ObjectRepositoryPagesRoot(string workspaceRoot) =>
        Path.GetFullPath(Path.Combine(workspaceRoot, "object-repository", "pages"));

    /// <summary>
    /// Candidate outputs and approval sources must live under knowledge/candidates.
    /// Rejects knowledge/approved and object-repository/pages.
    /// </summary>
    public static void EnsureUnderCandidates(string fullPath, string workspaceRoot)
    {
        var normalized = Path.GetFullPath(fullPath);
        var candidates = CandidatesRoot(workspaceRoot);
        try
        {
            EnsureWritablePathInside(normalized, candidates);
        }
        catch (KnowledgePathException ex)
        {
            throw new KnowledgePathException(
                $"path '{fullPath}' must resolve under knowledge/candidates. {ex.Message}");
        }

        RejectForbiddenKnowledgeRoots(normalized, workspaceRoot);
    }

    public static void EnsureCandidateSource(string fullPath, string workspaceRoot)
    {
        var normalized = Path.GetFullPath(fullPath);
        var candidates = CandidatesRoot(workspaceRoot);
        try
        {
            EnsureReadablePathInside(normalized, candidates);
        }
        catch (KnowledgePathException ex)
        {
            throw new KnowledgePathException(
                $"approval source '{fullPath}' must resolve under knowledge/candidates. {ex.Message}");
        }

        RejectForbiddenKnowledgeRoots(normalized, workspaceRoot);
    }

    private static void RejectForbiddenKnowledgeRoots(string fullPath, string workspaceRoot)
    {
        var approved = ApprovedRoot(workspaceRoot);
        var pages = ObjectRepositoryPagesRoot(workspaceRoot);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (IsInsideOrEqual(normalized, approved))
        {
            throw new KnowledgePathException(
                $"path '{fullPath}' resolves under knowledge/approved and cannot be used as a candidate path.");
        }

        if (IsInsideOrEqual(normalized, pages))
        {
            throw new KnowledgePathException(
                $"path '{fullPath}' resolves under object-repository/pages and is forbidden for knowledge operations.");
        }
    }

    private static bool IsInsideOrEqual(string fullPath, string directoryRoot)
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

    private static string ResolveFinalPath(
        string fullPath,
        string containmentRoot,
        HashSet<string> visited,
        int depth)
    {
        if (depth > MaxSymlinkDepth)
        {
            throw new KnowledgePathException(
                $"path '{fullPath}' exceeds the maximum symbolic-link resolution depth.");
        }

        var normalized = Path.GetFullPath(fullPath);
        if (!visited.Add(normalized))
        {
            throw new KnowledgePathException(
                $"path '{fullPath}' participates in a symbolic-link cycle.");
        }

        EnsureAncestorDirectoriesSafe(normalized, containmentRoot);

        // Directory reparse points (junctions / directory symlinks): resolve when the target
        // stays inside containment; reject escape. Previously FileInfo skipped directories entirely.
        var directoryInfo = new DirectoryInfo(normalized);
        if (directoryInfo.Exists)
        {
            if (!IsReparsePoint(directoryInfo))
                return normalized;

            var dirTarget = TryGetLinkTarget(directoryInfo);
            if (string.IsNullOrWhiteSpace(dirTarget))
            {
                throw new KnowledgePathException(
                    $"path '{fullPath}' is a directory reparse point that cannot be validated safely.");
            }

            var dirNext = Path.GetFullPath(
                Path.IsPathRooted(dirTarget)
                    ? dirTarget
                    : Path.Combine(directoryInfo.Parent?.FullName ?? normalized, dirTarget));

            if (!IsInsideDirectory(dirNext, containmentRoot))
            {
                throw new KnowledgePathException(
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
            throw new KnowledgePathException(
                $"path '{fullPath}' is a reparse point that cannot be validated safely.");
        }

        var next = Path.GetFullPath(
            Path.IsPathRooted(linkTarget)
                ? linkTarget
                : Path.Combine(fileInfo.Directory!.FullName, linkTarget));

        if (!IsInsideDirectory(next, containmentRoot))
        {
            throw new KnowledgePathException(
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
                // Safe directory symlink: allow only when the link target stays inside containment.
                var linkTarget = TryGetLinkTarget(current);
                if (string.IsNullOrWhiteSpace(linkTarget))
                {
                    throw new KnowledgePathException(
                        $"path '{fullPath}' traverses a directory reparse point '{current.FullName}' that cannot be validated safely.");
                }

                var resolved = Path.GetFullPath(
                    Path.IsPathRooted(linkTarget)
                        ? linkTarget
                        : Path.Combine(current.Parent?.FullName ?? normalizedRoot, linkTarget));

                if (!IsInsideDirectory(resolved, containmentRoot))
                {
                    throw new KnowledgePathException(
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

public sealed class KnowledgePathException : Exception
{
    public KnowledgePathException(string message) : base(message)
    {
    }
}
