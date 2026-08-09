using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Driver;

public sealed class UiExecutionResponse
{
    public bool Success { get; set; }

    public JsonElement? Value { get; set; }

    public string? Error { get; set; }

    public string? Reason { get; set; }

    public string? ScreenshotPath { get; set; }

    public JsonElement? Locator { get; set; }

    public JsonElement? Candidates { get; set; }

    public JsonElement? Suggestions { get; set; }

    [JsonIgnore]
    public int? HttpStatusCode { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
