using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Mcp;

internal static class McpPathGuard
{
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
            return workspace.ResolveSafePath(path);
        }
        catch (WorkspaceException ex)
        {
            var code = ex.Message.Contains("traversal", StringComparison.OrdinalIgnoreCase)
                || ex.Message.Contains("outside the workspace", StringComparison.OrdinalIgnoreCase)
                ? McpErrorCode.PathTraversal
                : McpErrorCode.Validation;
            throw new McpToolException(code, ex.Message);
        }
    }

    public static string? OptionalWorkspacePath(IWorkspaceManager workspace, string? path, string argumentName) =>
        string.IsNullOrWhiteSpace(path) ? null : RequireWorkspacePath(workspace, path, argumentName);
}
