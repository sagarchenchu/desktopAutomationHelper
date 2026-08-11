using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Tests;

public class SuiteManifestReaderTests
{
    [Fact]
    public void ValidSuiteManifest_Succeeds()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "smoke",
              "enabled": true,
              "testCases": [
                { "jiraKey": "SAMPLE-1", "enabled": true },
                { "jiraKey": "SAMPLE-2", "enabled": false }
              ]
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");

        Assert.True(result.IsValid);
        Assert.Equal(2, result.TotalCount);
        Assert.Equal(1, result.EnabledCount);
        Assert.Equal(1, result.DisabledCount);
        Assert.Equal(["SAMPLE-1"], result.EnabledJiraKeys);
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void InvalidJiraKey_FailsWithEntryLocation()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "smoke",
              "testCases": [ { "jiraKey": "not-a-key" } ]
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e =>
            e.Contains("testCases[0]", StringComparison.Ordinal)
            && e.Contains("invalid jiraKey", StringComparison.Ordinal)
            && e.Contains("canonical Jira syntax", StringComparison.Ordinal));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void DuplicateJiraKey_Fails()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "smoke",
              "testCases": [
                { "jiraKey": "SAMPLE-1" },
                { "jiraKey": "SAMPLE-1" }
              ]
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");

        Assert.False(result.IsValid);
        Assert.Equal(1, result.DuplicateCount);
        Assert.Contains(result.Errors, e => e.Contains("duplicate jiraKey", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void DisabledEntries_AreExcludedFromEffectiveSelection()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "regression",
              "enabled": true,
              "testCases": [
                { "jiraKey": "SAMPLE-1", "enabled": false },
                { "jiraKey": "SAMPLE-2", "enabled": true }
              ]
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");

        Assert.True(result.IsValid);
        Assert.Equal(["SAMPLE-2"], result.EnabledJiraKeys);
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void MissingSuiteFile_Fails()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/missing.json");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("not found", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void UnsupportedSuiteSchema_Fails()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 99,
              "name": "smoke",
              "testCases": []
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("unsupported schemaVersion", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void OmittedTestCasesProperty_Fails()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "smoke",
              "enabled": true
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("'testCases' is required", StringComparison.Ordinal));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void ExplicitEmptyTestCasesArray_Succeeds()
    {
        var (options, root) = CreateWorkspaceWithSuite(
            """
            {
              "schemaVersion": 1,
              "name": "smoke",
              "enabled": true,
              "testCases": []
            }
            """);

        var result = TestSupport.CreateSuiteReader(options).ValidateFile("suites/smoke.json");
        Assert.True(result.IsValid);
        Assert.Equal(0, result.TotalCount);
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void ValidateKeys_EmptyInput_Fails()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var result = TestSupport.CreateSuiteReader(options).ValidateKeys(Array.Empty<string>());
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("At least one", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void ValidateKeys_AcceptsValidAndRejectsInvalid()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var reader = TestSupport.CreateSuiteReader(options);

        var ok = reader.ValidateKeys(["SAMPLE-1", "SAMPLE-2"]);
        Assert.True(ok.IsValid);

        var bad = reader.ValidateKeys(["SAMPLE-1", "bad", "SAMPLE-1"]);
        Assert.False(bad.IsValid);
        Assert.Contains(bad.Errors, e => e.Contains("canonical Jira syntax", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(bad.Errors, e => e.Contains("duplicate jiraKey", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void ProjectSpecificPattern_RejectsOtherwiseCanonicalKey()
    {
        var options = TestSupport.CreateOptions();
        options.Suites.JiraKeyPattern = @"^SAMPLE-[1-9][0-9]{0,15}$";
        Directory.CreateDirectory(options.Workspace.Root);
        var reader = TestSupport.CreateSuiteReader(options);

        var result = reader.ValidateKeys(["SAMPLE-1", "OTHER-1"]);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e =>
            e.Contains("OTHER-1", StringComparison.Ordinal)
            && e.Contains("project-specific pattern", StringComparison.Ordinal));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    private static (Configuration.AgentOptions Options, string Root) CreateWorkspaceWithSuite(string json)
    {
        var options = TestSupport.CreateOptions();
        var suites = Path.Combine(options.Workspace.Root, "suites");
        Directory.CreateDirectory(suites);
        File.WriteAllText(Path.Combine(suites, "smoke.json"), json);
        return (options, options.Workspace.Root);
    }
}
