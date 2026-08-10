using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class PageReference
{
    public string PageId { get; set; } = string.Empty;

    public string File { get; set; } = string.Empty;

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
