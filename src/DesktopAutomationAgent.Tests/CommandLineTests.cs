using DesktopAutomationAgent.Cli;

namespace DesktopAutomationAgent.Tests;

public class CommandLineTests
{
    [Fact]
    public void Parse_DoctorJson()
    {
        var parsed = CommandLine.Parse(["doctor", "--json"]);
        Assert.Equal(AgentCommandKind.Doctor, parsed.Kind);
        Assert.True(parsed.Json);
        Assert.Null(parsed.Error);
    }

    [Fact]
    public void Parse_ValidateSuiteRequiresFile()
    {
        var parsed = CommandLine.Parse(["validate-suite"]);
        Assert.NotNull(parsed.Error);
        Assert.Contains("--file", parsed.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Parse_ValidateKeysSplitsCommaList()
    {
        var parsed = CommandLine.Parse(["validate-keys", "--keys", "SAMPLE-1, SAMPLE-2"]);
        Assert.Equal(["SAMPLE-1", "SAMPLE-2"], parsed.Keys);
    }

    [Fact]
    public void Parse_ValidateKeysRejectsEmptyCommaList()
    {
        var parsed = CommandLine.Parse(["validate-keys", "--keys", ",,,"]);
        Assert.NotNull(parsed.Error);
        Assert.Contains("at least one", parsed.Error, StringComparison.OrdinalIgnoreCase);
    }
}
