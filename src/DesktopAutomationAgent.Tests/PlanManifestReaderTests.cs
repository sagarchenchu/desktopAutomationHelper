using System.Security.Cryptography;
using System.Text;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Tests;

public class PlanManifestReaderTests
{
    [Fact]
    public void Read_ValidMinimalPlan()
    {
        var options = TestSupport.CreateOptions();
        var path = TestSupport.WritePlan(options, "ok.plan.json", TestSupport.MinimalPlanJson());
        var result = TestSupport.CreatePlanReader(options).Read(path);

        Assert.True(result.IsValid, string.Join("; ", result.Errors));
        Assert.Equal("SAMPLE-1", result.PlanId);
        Assert.Equal(1, result.StepCount);
        Assert.False(string.IsNullOrWhiteSpace(result.Sha256));
        Assert.NotNull(result.Plan);
    }

    [Fact]
    public void Read_MissingSteps()
    {
        var json = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Missing steps"
            }
            """;
        AssertInvalid(json, "steps is required");
    }

    [Fact]
    public void Read_EmptySteps()
    {
        var json = TestSupport.MinimalPlanJson(stepsOverride: "[]");
        AssertInvalid(json, "at least one step");
    }

    [Fact]
    public void Read_UnsupportedSchemaVersion()
    {
        var json = """
            {
              "schemaVersion": 99,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Bad",
              "steps": [ { "id": "a", "operation": "listwindows", "arguments": {} } ]
            }
            """;
        AssertInvalid(json, "schemaVersion");
    }

    [Fact]
    public void Read_WrongCatalogSchemaVersion()
    {
        var json = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 1,
              "planId": "SAMPLE-1",
              "name": "Bad",
              "steps": [ { "id": "a", "operation": "listwindows", "arguments": {} } ]
            }
            """;
        AssertInvalid(json, "catalogSchemaVersion");
    }

    [Fact]
    public void Read_MissingPlanId()
    {
        var json = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "name": "Bad",
              "steps": [ { "id": "a", "operation": "listwindows", "arguments": {} } ]
            }
            """;
        AssertInvalid(json, "planId");
    }

    [Fact]
    public void Read_InvalidPlanId()
    {
        AssertInvalid(TestSupport.MinimalPlanJson(planId: "bad id!"), "planId");
    }

    [Fact]
    public void Read_DuplicateStepIds()
    {
        var steps = """
            [
              { "id": "same", "operation": "listwindows", "arguments": {} },
              { "id": "same", "operation": "listwindows", "arguments": {} }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "duplicates");
    }

    [Fact]
    public void Read_DuplicateIdsDifferingOnlyByCase()
    {
        var steps = """
            [
              { "id": "StepA", "operation": "listwindows", "arguments": {} },
              { "id": "stepa", "operation": "listwindows", "arguments": {} }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "duplicates");
    }

    [Fact]
    public void Read_DuplicateJsonProperties()
    {
        var json = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Dup",
              "name": "Dup2",
              "steps": [ { "id": "a", "operation": "listwindows", "arguments": {} } ]
            }
            """;
        AssertInvalid(json, "Duplicate");
    }

    [Fact]
    public void Read_ArgumentsNotObject()
    {
        var json = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Bad args",
              "steps": [ { "id": "a", "operation": "listwindows", "arguments": [] } ]
            }
            """;
        AssertInvalid(json, "invalid JSON");
    }

    [Fact]
    public void Read_ReservedArgumentProperty()
    {
        var steps = """
            [
              {
                "id": "a",
                "operation": "listwindows",
                "arguments": { "bearerToken": "secret" }
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "reserved");
    }

    [Fact]
    public void Read_InvalidAssertionOperator()
    {
        var steps = """
            [
              {
                "id": "a",
                "operation": "listwindows",
                "arguments": {},
                "assertions": [ { "path": "", "operator": "kindaEquals", "expected": 1 } ]
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "operator");
    }

    [Fact]
    public void Read_MissingAssertionExpected()
    {
        var steps = """
            [
              {
                "id": "a",
                "operation": "listwindows",
                "arguments": {},
                "assertions": [ { "path": "", "operator": "equals" } ]
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "expected is required");
    }

    [Fact]
    public void Read_InvalidJsonPointer()
    {
        var steps = """
            [
              {
                "id": "a",
                "operation": "listwindows",
                "arguments": {},
                "assertions": [ { "path": "title", "operator": "isNotNull" } ]
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "JSON pointer");
    }

    [Fact]
    public void Read_PlanOutsideWorkspace()
    {
        var options = TestSupport.CreateOptions();
        Directory.CreateDirectory(options.Workspace.Root);
        var outside = Path.Combine(Path.GetTempPath(), "outside-" + Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(outside, TestSupport.MinimalPlanJson());

        var result = TestSupport.CreatePlanReader(options).Read(outside);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("outside", StringComparison.OrdinalIgnoreCase));
        File.Delete(outside);
    }

    [Fact]
    public void Read_OversizedPlan()
    {
        var options = TestSupport.CreateOptions(maxPlanBytes: 64);
        var big = TestSupport.MinimalPlanJson(extraTopLevel: $"\"description\": \"{new string('x', 200)}\"");
        var path = TestSupport.WritePlan(options, "big.plan.json", big);
        var result = TestSupport.CreatePlanReader(options).Read(path);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("maximum size", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Read_UnknownTopLevelProperty()
    {
        AssertInvalid(
            TestSupport.MinimalPlanJson(extraTopLevel: "\"cleanupSteps\": []"),
            "unknown top-level property");
    }

    [Fact]
    public void Read_UnknownStepProperty()
    {
        var steps = """
            [
              {
                "id": "a",
                "operation": "listwindows",
                "arguments": {},
                "retry": true
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(stepsOverride: steps), "unknown property");
    }

    [Fact]
    public void Read_CleanupRejectsAssertionsAndCaptureResponse()
    {
        var onFailure = """
            [
              {
                "id": "cleanup",
                "operation": "quit",
                "arguments": {},
                "captureResponse": true,
                "assertions": [ { "operator": "isTrue" } ]
              }
            ]
            """;
        AssertInvalid(TestSupport.MinimalPlanJson(onFailure: onFailure), "not allowed on cleanup");
    }

    [Fact]
    public void Read_ComputesSha256OfExactBytes()
    {
        var options = TestSupport.CreateOptions();
        var json = TestSupport.MinimalPlanJson();
        var path = TestSupport.WritePlan(options, "hash.plan.json", json);
        var expected = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
        var result = TestSupport.CreatePlanReader(options).Read(path);
        Assert.Equal(expected, result.Sha256);
    }

    [Fact]
    public void SchemaAndValidator_ShareCoreRules()
    {
        var schemaPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "automation", "schemas", "plan.schema.json"));
        if (!File.Exists(schemaPath))
        {
            schemaPath = Path.GetFullPath(Path.Combine(
                Directory.GetCurrentDirectory(),
                "automation", "schemas", "plan.schema.json"));
        }

        Assert.True(File.Exists(schemaPath), $"Missing schema at {schemaPath}");
        var schema = File.ReadAllText(schemaPath);
        Assert.Contains("\"const\": 1", schema);
        Assert.Contains("\"const\": 2", schema);
        Assert.Contains("^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$", schema);
        Assert.Contains("matchesRegex", schema);
        Assert.Contains("cleanupStep", schema);
        Assert.DoesNotContain("\"cleanupSteps\"", schema);

        Assert.True(PlanValidator.IsValidJsonPointer(""));
        Assert.True(PlanValidator.IsValidJsonPointer("/a/0"));
        Assert.False(PlanValidator.IsValidJsonPointer("a"));
    }

    private static void AssertInvalid(string json, string expectedFragment)
    {
        var options = TestSupport.CreateOptions();
        var path = TestSupport.WritePlan(options, "bad.plan.json", json);
        var result = TestSupport.CreatePlanReader(options).Read(path);
        Assert.False(result.IsValid);
        Assert.True(
            result.Errors.Any(e => e.Contains(expectedFragment, StringComparison.OrdinalIgnoreCase)),
            $"Expected error containing '{expectedFragment}'. Errors: {string.Join(" | ", result.Errors)}");
    }
}
