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
                throw new KnowledgePathException(
                    $"path '{fullPath}' traverses a symbolic link or reparse point directory '{current.FullName}'.");
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
