using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Suites;

public sealed class SuiteManifest
{
    public int SchemaVersion { get; set; }

    public string Name { get; set; } = string.Empty;

    public bool Enabled { get; set; } = true;

    /// <summary>
    /// Required suite property. Null means the JSON omitted <c>testCases</c>
    /// (or set it to null). An empty array is valid and means no cases yet.
    /// </summary>
    public List<SuiteTestCase>? TestCases { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class SuiteTestCase
{
    public string JiraKey { get; set; } = string.Empty;

    public bool Enabled { get; set; } = true;

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
