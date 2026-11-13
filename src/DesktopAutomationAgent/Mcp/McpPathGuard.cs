using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// Workspace containment for MCP file arguments, including symlink / junction / chained-link escape rejection.
/// </summary>
internal static class McpPathGuard
{
    private const int MaxSymlinkDepth = 32;

    public static string RequireWorkspacePath(IWorkspaceManager workspace, string? path, string argumentName)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new McpToolException(McpErrorCode.Validation, $"{argumentName} is required.");

        if (path.Length > McpResponseBudget.DefaultPathMaxChars)
        {
            throw new McpToolException(
                McpErrorCode.Validation,
                $"{argumentName} exceeds maximum length of {McpResponseBudget.DefaultPathMaxChars}.");
        }

        if (path.Contains("..", StringComparison.Ordinal)
            || path.Contains('\0')
            || path.Contains('|'))
        {
            throw new McpToolException(
                McpErrorCode.PathTraversal,
                $"Path traversal or unsafe path rejected for {argumentName}.",
                path);
        }

        try
        {
            var resolved = workspace.ResolveSafePath(path);
            EnsureNotSymlinkEscape(resolved, workspace.RootPath);
            return resolved;
        }
        catch (WorkspaceException ex)
        {
            var code = ex.Message.Contains("traversal", StringComparison.OrdinalIgnoreCase)
                || ex.Message.Contains("outside the workspace", StringComparison.OrdinalIgnoreCase)
                ? McpErrorCode.PathTraversal
                : McpErrorCode.Validation;
            throw new McpToolException(code, ex.Message);
        }
        catch (McpPathException ex)
        {
            throw new McpToolException(McpErrorCode.PathTraversal, ex.Message);
        }
    }

    public static string? OptionalWorkspacePath(IWorkspaceManager workspace, string? path, string argumentName) =>
        string.IsNullOrWhiteSpace(path) ? null : RequireWorkspacePath(workspace, path, argumentName);

    /// <summary>
    /// Validates an already workspace-resolved path (and any existing file) for link escape before I/O.
    /// </summary>
    public static void EnsureReadablePathInside(string fullPath, string containmentRoot)
    {
        try
        {
            var normalizedRoot = Path.GetFullPath(containmentRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var normalizedPath = Path.GetFullPath(fullPath);

            if (!IsInsideDirectory(normalizedPath, normalizedRoot))
            {
                throw new McpPathException(
                    $"path '{fullPath}' resolves outside the allowed directory.");
            }

            EnsureNotSymlinkEscape(normalizedPath, normalizedRoot);
        }
        catch (McpPathException ex)
        {
            throw new McpToolException(McpErrorCode.PathTraversal, ex.Message);
        }
    }

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
            throw new McpPathException(
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
            throw new McpPathException(
                $"path '{fullPath}' exceeds the maximum symbolic-link resolution depth.");
        }

        var normalized = Path.GetFullPath(fullPath);
        if (!visited.Add(normalized))
        {
            throw new McpPathException(
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
                throw new McpPathException(
                    $"path '{fullPath}' is a directory reparse point that cannot be validated safely.");
            }

            var dirNext = Path.GetFullPath(
                Path.IsPathRooted(dirTarget)
                    ? dirTarget
                    : Path.Combine(directoryInfo.Parent?.FullName ?? normalized, dirTarget));

            if (!IsInsideDirectory(dirNext, containmentRoot))
            {
                throw new McpPathException(
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
            throw new McpPathException(
                $"path '{fullPath}' is a reparse point that cannot be validated safely.");
        }

        var next = Path.GetFullPath(
            Path.IsPathRooted(linkTarget)
                ? linkTarget
                : Path.Combine(fileInfo.Directory!.FullName, linkTarget));

        if (!IsInsideDirectory(next, containmentRoot))
        {
            throw new McpPathException(
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
                    throw new McpPathException(
                        $"path '{fullPath}' traverses a directory reparse point '{current.FullName}' that cannot be validated safely.");
                }

                var resolved = Path.GetFullPath(
                    Path.IsPathRooted(linkTarget)
                        ? linkTarget
                        : Path.Combine(current.Parent?.FullName ?? normalizedRoot, linkTarget));

                if (!IsInsideDirectory(resolved, containmentRoot))
                {
                    throw new McpPathException(
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

internal sealed class McpPathException : Exception
{
    public McpPathException(string message) : base(message)
    {
    }
}
