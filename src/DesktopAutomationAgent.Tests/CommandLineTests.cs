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

    [Fact]
    public void Parse_ValidatePlanAndRunPlan()
    {
        var validate = CommandLine.Parse(["validate-plan", "--file", "p.json", "--json"]);
        Assert.Equal(AgentCommandKind.ValidatePlan, validate.Kind);
        Assert.True(validate.Json);
        Assert.Equal("p.json", validate.PlanFile);

        var run = CommandLine.Parse(["run-plan", "--file", "p.json", "--dry-run", "--json"]);
        Assert.Equal(AgentCommandKind.RunPlan, run.Kind);
        Assert.True(run.DryRun);
        Assert.True(run.Json);
    }

    [Fact]
    public void Parse_RunPlanRequiresFile()
    {
        var parsed = CommandLine.Parse(["run-plan"]);
        Assert.NotNull(parsed.Error);
        Assert.Contains("--file", parsed.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("run-plan", "--json")]
    [InlineData("run-plan", "--json", "--file")]
    [InlineData("validate-plan", "--json", "--bogus")]
    public void Parse_PreservesJsonFlagOnUsageErrors(params string[] args)
    {
        var parsed = CommandLine.Parse(args);
        Assert.NotNull(parsed.Error);
        Assert.True(parsed.Json, $"Expected Json=true for args: {string.Join(' ', args)}");
    }
}
