using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Tests;

public class JiraKeyContractTests
{
    public static readonly object[][] ValidKeys =
    [
        ["A-1"],
        ["ABC-1234"],
        ["ABC_1-999"],
        // Project portion at supported boundary: 1 letter + 31 alnum/underscore = 32
        ["ABCDEFGHIJKLMNOPQRSTUVWXYZ012345-1"],
        // Issue number at supported boundary: 16 digits, no leading zero
        ["A-1234567890123456"],
    ];

    public static readonly object[][] InvalidKeys =
    [
        ["ABC-0"],
        ["ABC--1"],
        ["ABC-"],
        ["-123"],
        ["abc-1234"], // lowercase suite key
        ["ABC 1234"], // whitespace inside
        ["ABC/1234"], // path separator
        ["ABC\\1234"],
        ["ABC.1234"], // dots
        // Project portion exceeding 32 characters (1 + 32 rest)
        ["ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456-1"],
        // Issue portion exceeding 16 digits
        ["A-12345678901234567"],
        ["ABC-01"], // leading-zero issue numbers
        ["ABC-0123"],
    ];

    [Theory]
    [MemberData(nameof(ValidKeys))]
    public void Canonical_AcceptsValidKeys(string key)
    {
        Assert.True(JiraKeyContract.IsCanonical(key));
        Assert.True(
            JiraKeyContract.TryValidate(key, JiraKeyContract.CanonicalPattern, out var normalized, out var error),
            error);
        Assert.Equal(key, normalized);
    }

    [Theory]
    [MemberData(nameof(InvalidKeys))]
    public void Canonical_RejectsInvalidKeys(string key)
    {
        Assert.False(JiraKeyContract.IsCanonical(key));
        Assert.False(JiraKeyContract.TryValidate(key, JiraKeyContract.CanonicalPattern, out _, out var error));
        Assert.Contains("canonical Jira syntax", error, StringComparison.Ordinal);
    }

    [Fact]
    public void TryValidate_DistinguishesProjectSpecificRejection()
    {
        const string projectOnly = @"^ABC-[1-9][0-9]{0,15}$";
        Assert.True(JiraKeyContract.IsCanonical("XYZ-1"));
        Assert.False(JiraKeyContract.TryValidate("XYZ-1", projectOnly, out _, out var error));
        Assert.Contains("project-specific pattern", error, StringComparison.Ordinal);
        Assert.DoesNotContain("canonical Jira syntax", error, StringComparison.Ordinal);
    }

    [Fact]
    public void TryValidate_RejectsDuplicateKeysAfterNormalValidation()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var reader = TestSupport.CreateSuiteReader(options);

        var result = reader.ValidateKeys(["ABC-1", "ABC-1"]);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("duplicate jiraKey", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public void SuitesJiraKeyPattern_CannotBroadenBeyondCanonical()
    {
        // A permissive project pattern still cannot accept ABC-0 because canonical runs first.
        const string permissive = @"^[A-Z][A-Z0-9_]*-[0-9]+$";
        Assert.False(JiraKeyContract.TryValidate("ABC-0", permissive, out _, out var error));
        Assert.Contains("canonical Jira syntax", error, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalPattern_IsCultureInvariantCompiledWithTimeout()
    {
        var regex = JiraKeyContract.CompileProjectPattern(JiraKeyContract.CanonicalPattern);
        Assert.Equal(RegexOptions.Compiled | RegexOptions.CultureInvariant, regex.Options & (RegexOptions.Compiled | RegexOptions.CultureInvariant));
        Assert.Equal(JiraKeyContract.MatchTimeout, regex.MatchTimeout);
        Assert.Equal(JiraKeyContract.CanonicalPattern, regex.ToString());
        Assert.Matches(JiraKeyContract.CanonicalPattern, "A-1");
        Assert.DoesNotMatch(JiraKeyContract.CanonicalPattern, "a-1");
    }

    [Fact]
    public void CanonicalJiraPattern_DoesNotDriftAcrossSchemasAndDefaults()
    {
        var expected = JiraKeyContract.CanonicalPattern;

        var suiteSchema = ReadRepoJson("automation/schemas/suite.schema.json");
        var bddSchema = ReadRepoJson("automation/schemas/bdd-action-map.schema.json");
        var appsettings = ReadRepoJson("src/DesktopAutomationAgent/appsettings.json");
        var agentExample = ReadRepoJson("automation/config/agentsettings.example.json");

        Assert.Equal(expected, ExtractJiraPattern(suiteSchema, suiteSchemaPath: true));
        Assert.Equal(expected, ExtractJiraPattern(bddSchema, suiteSchemaPath: false));
        Assert.Equal(expected, appsettings.GetProperty("Suites").GetProperty("JiraKeyPattern").GetString());
        Assert.Equal(expected, agentExample.GetProperty("Suites").GetProperty("JiraKeyPattern").GetString());
        Assert.Equal(expected, new SuiteOptions().JiraKeyPattern);
    }

    private static JsonElement ReadRepoJson(string relativePath)
    {
        var full = Path.GetFullPath(Path.Combine(FindRepoRoot(), relativePath));
        using var doc = JsonDocument.Parse(File.ReadAllText(full));
        return doc.RootElement.Clone();
    }

    private static string ExtractJiraPattern(JsonElement root, bool suiteSchemaPath)
    {
        if (suiteSchemaPath)
        {
            return root
                .GetProperty("properties")
                .GetProperty("testCases")
                .GetProperty("items")
                .GetProperty("properties")
                .GetProperty("jiraKey")
                .GetProperty("pattern")
                .GetString()
                ?? string.Empty;
        }

        return root
            .GetProperty("properties")
            .GetProperty("jiraKey")
            .GetProperty("pattern")
            .GetString()
            ?? string.Empty;
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.slnx")))
                return dir.FullName;
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root from test base directory.");
    }
}
