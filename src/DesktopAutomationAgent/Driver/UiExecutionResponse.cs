using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Driver;

public sealed class UiExecutionResponse
{
    public bool Success { get; set; }

    public JsonElement? Value { get; set; }

    public string? Error { get; set; }

    public string? Reason { get; set; }

    /// <summary>Typed driver execution state (e.g. notStarted). Preferred over free-text errors.</summary>
    public string? ExecutionState { get; set; }

    /// <summary>Typed driver failure code (e.g. locatorNotFoundBeforeInvoke).</summary>
    public string? FailureCode { get; set; }

    public string? ScreenshotPath { get; set; }

    public JsonElement? Locator { get; set; }

    public JsonElement? Candidates { get; set; }

    public JsonElement? Suggestions { get; set; }

    [JsonIgnore]
    public int? HttpStatusCode { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
