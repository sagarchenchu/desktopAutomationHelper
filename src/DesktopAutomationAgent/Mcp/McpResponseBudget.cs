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
    public const int MaxTrimIterations = 10_000;

    /// <summary>
    /// Serializes <paramref name="payload"/> to compact JSON that is always parseable and
    /// never exceeds <paramref name="maxChars"/>. Throws <see cref="McpToolException"/> with
    /// <see cref="McpErrorCode.Truncated"/> when a valid bounded envelope cannot be produced.
    /// </summary>
    public static string SerializeBounded(object payload, int maxChars = DefaultMaxChars)
    {
        maxChars = Math.Clamp(maxChars, 256, AbsoluteMaxChars);

        string json;
        try
        {
            json = JsonSerializer.Serialize(payload, McpJson.Options);
        }
        catch (Exception ex) when (ex is JsonException or NotSupportedException or ArgumentException)
        {
            throw new McpToolException(
                McpErrorCode.Truncated,
                "Response payload could not be serialized under the MCP budget.",
                ex.Message);
        }

        if (json.Length <= maxChars)
        {
            EnsureValidJson(json);
            return json;
        }

        JsonNode? node;
        try
        {
            node = JsonNode.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new McpToolException(
                McpErrorCode.Truncated,
                "Response payload produced invalid JSON before budgeting.",
                ex.Message);
        }

        if (node is not JsonObject obj)
        {
            return BuildAndValidateEnvelope(
                new
                {
                    success = false,
                    errorCode = McpErrorCode.Truncated,
                    message = "Non-object response exceeded budget.",
                    truncated = true,
                    truncationReason = "response_budget"
                },
                maxChars);
        }

        obj["truncated"] = true;
        obj["truncationReason"] = "response_budget";

        var previousLength = int.MaxValue;
        for (var iteration = 0; iteration < MaxTrimIterations; iteration++)
        {
            var current = obj.ToJsonString(McpJson.NodeOptions);
            if (current.Length <= maxChars)
            {
                EnsureValidJson(current);
                return current;
            }

            if (current.Length >= previousLength)
            {
                // No progress — fall through to minimal envelope.
                break;
            }

            previousLength = current.Length;
            if (!TryTrimOnce(obj))
                break;
        }

        return BuildAndValidateEnvelope(
            new
            {
                success = false,
                errorCode = McpErrorCode.Truncated,
                message = "Response exceeded budget after structured truncation.",
                truncated = true,
                truncationReason = "response_budget_exhausted"
            },
            maxChars);
    }

    private static string BuildAndValidateEnvelope(object envelope, int maxChars)
    {
        var json = JsonSerializer.Serialize(envelope, McpJson.Options);
        if (json.Length > maxChars)
        {
            throw new McpToolException(
                McpErrorCode.Truncated,
                $"Response budget ({maxChars}) is too small for a valid JSON error envelope.");
        }

        EnsureValidJson(json);
        return json;
    }

    private static void EnsureValidJson(string json)
    {
        try
        {
            using var _ = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new McpToolException(
                McpErrorCode.Truncated,
                "Bounded response was not valid JSON.",
                ex.Message);
        }
    }

    /// <summary>
    /// Performs one progressive trim step. Returns false when nothing more can be removed.
    /// </summary>
    private static bool TryTrimOnce(JsonObject obj)
    {
        string? heaviestKey = null;
        JsonNode? heaviestValue = null;
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
                heaviestValue = value;
            }
        }

        if (heaviestKey is null || heaviestLen == 0)
            return false;

        switch (heaviestValue)
        {
            case JsonArray arr when arr.Count > 0:
            {
                var keep = Math.Max(0, arr.Count / 2);
                if (keep == arr.Count)
                    keep = arr.Count - 1;
                while (arr.Count > keep)
                    arr.RemoveAt(arr.Count - 1);
                return true;
            }
            case JsonObject child:
            {
                if (TryTrimOnce(child))
                    return true;
                obj.Remove(heaviestKey);
                return true;
            }
            case JsonValue value:
            {
                if (value.TryGetValue<string>(out var text) && text.Length > 32)
                {
                    var keep = Math.Max(16, text.Length / 2);
                    obj[heaviestKey] = text[..keep] + "…";
                    return true;
                }

                obj.Remove(heaviestKey);
                return true;
            }
            default:
                obj.Remove(heaviestKey);
                return true;
        }
    }
}
