using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Suites;

public sealed class SuiteManifest
{
    public int SchemaVersion { get; set; }

    public string Name { get; set; } = string.Empty;

    public bool Enabled { get; set; } = true;

    public List<SuiteTestCase> TestCases { get; set; } = [];

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
