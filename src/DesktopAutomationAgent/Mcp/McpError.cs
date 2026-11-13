using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Mcp;

/// <summary>Stable MCP error codes for structured tool failures.</summary>
public static class McpErrorCode
{
    public const string Validation = "validation";
    public const string PathTraversal = "path_traversal";
    public const string StaleRef = "stale_ref";
    public const string Ambiguous = "ambiguous";
    public const string Denied = "denied";
    public const string DriverUnavailable = "driver_unavailable";
    public const string Cancelled = "cancelled";
    public const string Truncated = "truncated";
    public const string Internal = "internal";
}

public sealed class McpErrorPayload
{
    public bool Success { get; init; }

    public string ErrorCode { get; init; } = McpErrorCode.Internal;

    public string Message { get; init; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Detail { get; init; }
}

/// <summary>Thrown by the MCP facade; mapped to structured CallToolResult errors.</summary>
public sealed class McpToolException : Exception
{
    public McpToolException(string errorCode, string message, string? detail = null)
        : base(message)
    {
        ErrorCode = errorCode;
        Detail = detail;
    }

    public string ErrorCode { get; }

    public string? Detail { get; }
}
