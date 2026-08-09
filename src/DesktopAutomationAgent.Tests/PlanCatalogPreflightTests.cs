using System.Text.Json;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Tests;

public class PlanCatalogPreflightTests
{
    private readonly PlanCatalogPreflight _preflight = new(CatalogFixtures.Phase2Catalog());

    [Fact]
    public void AcceptsCanonicalOperation()
    {
        var manifest = Manifest(("s1", "listwindows", "{}"));
        Assert.Empty(_preflight.Validate(manifest, "p.json"));
    }

    [Fact]
    public void AcceptsCanonicalOperationCaseInsensitivelyAndNormalizes()
    {
        var manifest = Manifest(("s1", "ListWindows", "{}"));
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Empty(errors);
        Assert.Equal("listwindows", manifest.Steps![0].Operation);
    }

    [Fact]
    public void RejectsAliasWithCanonicalSuggestion()
    {
        var manifest = LaunchPlan(operation: "start");
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("alias", StringComparison.OrdinalIgnoreCase)
                                     && e.Contains("launch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsDeprecatedAlias()
    {
        var manifest = LaunchPlan(operation: "runapp");
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("alias", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsUnknownOperation()
    {
        var manifest = Manifest(("s1", "not-a-real-op", "{}"));
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("unknown operation", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsMissingFixedRequiredInput()
    {
        var manifest = Manifest(("s1", "launch", "{}"), ("s2", "quit", "{}"));
        manifest.OnFailureSteps =
        [
            Step("cleanup", "quit", "{}")
        ];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("missing required argument 'value'", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AcceptsValidInputAlternative()
    {
        var manifest = Manifest(
            ("launch", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("set", "setvalue", """{"locator":{"automationId":"x"},"value":"1"}"""),
            ("quit", "quit", "{}"));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        Assert.Empty(_preflight.Validate(manifest, "p.json"));
    }

    [Fact]
    public void RejectsWhenEveryInputAlternativeMissing()
    {
        var manifest = Manifest(
            ("launch", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("set", "setvalue", """{"value":"1"}"""),
            ("quit", "quit", "{}"));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("complete argument alternative", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsNullRequiredInput()
    {
        var manifest = Manifest(
            ("launch", "launch", """{"value":null}"""),
            ("quit", "quit", "{}"));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("must not be null", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsEmptyRequiredLocator()
    {
        var manifest = Manifest(
            ("launch", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("click", "click", """{"locator":{}}"""),
            ("quit", "quit", "{}"));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("empty locator", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsSessionRequiredBeforeLaunch()
    {
        var manifest = Manifest(("click", "click", """{"locator":{"automationId":"x"}}"""));
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("requires an active session", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsDoubleLaunch()
    {
        var manifest = Manifest(
            ("l1", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("l2", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("quit", "quit", "{}"));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("duplicate launch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsMainPlanEndingWithActiveSession()
    {
        var manifest = Manifest(("l1", "launch", """{"value":"C:\\\\App.exe"}"""));
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("end with close or quit", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RejectsLaunchWithoutFailureCleanup()
    {
        var manifest = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "SAMPLE-1",
            Name = "test",
            Steps =
            [
                Step("l1", "launch", """{"value":"C:\\\\App.exe"}"""),
                Step("quit", "quit", "{}")
            ]
        };
        var errors = _preflight.Validate(manifest, "p.json");
        Assert.Contains(errors, e => e.Contains("onFailureSteps", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AcceptsValidLaunchActionAssertionQuitPlan()
    {
        var manifest = Manifest(
            ("l1", "launch", """{"value":"C:\\\\App.exe"}"""),
            ("g1", "gettext", """{"locator":{"automationId":"DashboardTitle"}}"""),
            ("q1", "quit", "{}"));
        manifest.Steps![1].Assertions =
        [
            new PlanAssertion
            {
                Path = "",
                Operator = "equals",
                Expected = JsonSerializer.SerializeToElement("Dashboard")
            }
        ];
        manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        Assert.Empty(_preflight.Validate(manifest, "p.json"));
    }

    private static PlanManifest LaunchPlan(string operation) =>
        Manifest(
            ("l1", operation, """{"value":"C:\\\\App.exe"}"""),
            ("q1", "quit", "{}"));

    private static PlanManifest Manifest(params (string Id, string Operation, string ArgsJson)[] steps)
    {
        var manifest = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "SAMPLE-1",
            Name = "test",
            Steps = steps.Select(s => Step(s.Id, s.Operation, s.ArgsJson)).ToList()
        };

        if (steps.Any(s => string.Equals(s.Operation, "launch", StringComparison.OrdinalIgnoreCase)
                           || string.Equals(s.Operation, "start", StringComparison.OrdinalIgnoreCase)
                           || string.Equals(s.Operation, "runapp", StringComparison.OrdinalIgnoreCase)))
        {
            manifest.OnFailureSteps = [Step("cleanup", "quit", "{}")];
        }

        return manifest;
    }

    private static PlanStep Step(string id, string operation, string argsJson) =>
        new()
        {
            Id = id,
            Operation = operation,
            Arguments = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(argsJson)
        };
}
