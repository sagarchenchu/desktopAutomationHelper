using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace DesktopAutomationAgent.Mcp;

/// <summary>Bounds MCP tool JSON payloads so clients never receive unbounded trees or dumps.</summary>
public static class McpResponseBudget
{
    public const int DefaultMaxChars = 48_000;
    public const int AbsoluteMaxChars = 200_000;
    public const int DefaultMaxNodes = 200;
    public const int AbsoluteMaxNodes = 500;
    public const int DefaultMaxDepth = 8;
    public const int AbsoluteMaxDepth = 12;
    public const int DefaultMaxChildren = 50;
    public const int AbsoluteMaxChildren = 200;
    public const int DefaultSearchLimit = 25;
    public const int AbsoluteSearchLimit = 100;
    public const int DefaultPathMaxChars = 512;

    public static string SerializeBounded(object payload, int maxChars = DefaultMaxChars)
    {
        maxChars = Math.Clamp(maxChars, 1_024, AbsoluteMaxChars);
        var json = JsonSerializer.Serialize(payload, McpJson.Options);
        if (json.Length <= maxChars)
            return json;

        // Prefer structured truncation marker over raw mid-string cuts.
        var node = JsonNode.Parse(json);
        if (node is JsonObject obj)
        {
            TrimObject(obj, maxChars);
            obj["truncated"] = true;
            obj["truncationReason"] = "response_budget";
            var trimmed = obj.ToJsonString(McpJson.NodeOptions);
            if (trimmed.Length <= maxChars)
                return trimmed;
            json = trimmed;
        }

        var builder = new StringBuilder(maxChars);
        builder.Append(json.AsSpan(0, Math.Max(0, maxChars - 96)));
        builder.Append("…\",\"truncated\":true,\"truncationReason\":\"response_budget_hard_cut\"}");
        return builder.ToString();
    }

    private static void TrimObject(JsonObject obj, int maxChars)
    {
        // Drop the largest array/object properties first until under budget.
        while (obj.ToJsonString(McpJson.NodeOptions).Length > maxChars)
        {
            string? heaviestKey = null;
            var heaviestLen = 0;
            foreach (var (key, value) in obj)
            {
                if (key is "truncated" or "truncationReason" or "success" or "errorCode" or "message")
                    continue;
                var len = value?.ToJsonString(McpJson.NodeOptions).Length ?? 0;
                if (len > heaviestLen)
                {
                    heaviestLen = len;
                    heaviestKey = key;
                }
            }

            if (heaviestKey is null || heaviestLen == 0)
                break;

            if (obj[heaviestKey] is JsonArray arr && arr.Count > 0)
            {
                var keep = Math.Max(1, arr.Count / 2);
                while (arr.Count > keep)
                    arr.RemoveAt(arr.Count - 1);
                continue;
            }

            obj.Remove(heaviestKey);
        }
    }
}
