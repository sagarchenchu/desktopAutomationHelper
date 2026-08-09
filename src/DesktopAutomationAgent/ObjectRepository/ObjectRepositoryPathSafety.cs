using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.ObjectRepository;

internal static class ObjectRepositoryPathSafety
{
    public static void EnsureNotSymlinkEscape(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var current = new DirectoryInfo(Path.GetDirectoryName(fullPath) ?? normalizedRoot);
        var fileInfo = new FileInfo(fullPath);

        while (true)
        {
            if (current.Exists && IsSymbolicLink(current))
            {
                throw new RepositoryPathException(
                    $"path '{fullPath}' traverses a symbolic link or reparse point directory '{current.FullName}'.");
            }

            if (WorkspaceManager.PathsEqual(current.FullName, normalizedRoot))
                break;

            if (current.Parent is null)
                break;

            current = current.Parent;
        }

        if (fileInfo.Exists && IsSymbolicLink(fileInfo))
        {
            var linkTarget = fileInfo.LinkTarget;
            if (!string.IsNullOrWhiteSpace(linkTarget))
            {
                var resolvedTarget = Path.GetFullPath(
                    Path.IsPathRooted(linkTarget)
                        ? linkTarget
                        : Path.Combine(fileInfo.Directory!.FullName, linkTarget));

                if (!IsInsideDirectory(resolvedTarget, normalizedRoot))
                {
                    throw new RepositoryPathException(
                        $"symbolic link '{fullPath}' resolves outside the allowed directory.");
                }
            }
            else if (fileInfo.Attributes.HasFlag(FileAttributes.ReparsePoint))
            {
                throw new RepositoryPathException(
                    $"path '{fullPath}' is a reparse point that cannot be validated safely.");
            }
        }
    }

    public static void EnsureWritablePathInside(string fullPath, string containmentRoot)
    {
        var normalizedRoot = Path.GetFullPath(containmentRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedPath = Path.GetFullPath(fullPath);

        if (!IsInsideDirectory(normalizedPath, normalizedRoot))
        {
            throw new RepositoryPathException(
                $"path '{fullPath}' resolves outside the allowed directory.");
        }

        EnsureNotSymlinkEscape(normalizedPath, normalizedRoot);

        // Validate every existing ancestor; reject junctions/reparse points before CreateDirectory.
        var directory = Path.GetDirectoryName(normalizedPath);
        while (!string.IsNullOrWhiteSpace(directory))
        {
            var info = new DirectoryInfo(directory);
            if (info.Exists && IsSymbolicLink(info))
            {
                throw new RepositoryPathException(
                    $"output path '{fullPath}' traverses symbolic link or reparse point '{info.FullName}'.");
            }

            if (WorkspaceManager.PathsEqual(info.FullName, normalizedRoot))
                break;

            directory = info.Parent?.FullName;
        }
    }

    private static bool IsSymbolicLink(FileSystemInfo info)
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
}
