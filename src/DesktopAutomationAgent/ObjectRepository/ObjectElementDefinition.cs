using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectElementDefinition
{
    public string? Description { get; set; }

    public ObjectLocator? Locator { get; set; }

    /// <summary>
    /// Approved alternate locators for bounded recovery (Phase 7). Never auto-generated into active pages.
    /// </summary>
    public List<ObjectLocator>? AlternateLocators { get; set; }

    public ObjectQuality? Quality { get; set; }

    public ObjectSource? Source { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
