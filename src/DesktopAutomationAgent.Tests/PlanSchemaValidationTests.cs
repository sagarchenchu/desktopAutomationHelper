using System.Text.Json.Nodes;
using Json.Schema;

namespace DesktopAutomationAgent.Tests;

public class PlanSchemaValidationTests
{
    [Fact]
    public void PlanSchema_IsValidDraft202012_AndUsesPortablePatterns()
    {
        var schemaText = ReadRepoFile("automation/schemas/plan.schema.json");
        Assert.DoesNotContain("(?i)", schemaText);

        var schema = JsonSchema.FromText(schemaText);
        Assert.NotNull(schema);

        // Valid minimal plan
        var valid = JsonNode.Parse("""
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Sample",
              "steps": [
                { "id": "a", "operation": "listwindows", "arguments": {} }
              ]
            }
            """);
        Assert.True(schema.Evaluate(valid, EvaluationOptions()).IsValid);

        // Reserved argument name rejected (mixed case via portable pattern)
        var reserved = JsonNode.Parse("""
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Sample",
              "steps": [
                { "id": "a", "operation": "listwindows", "arguments": { "BearerToken": "x" } }
              ]
            }
            """);
        Assert.False(schema.Evaluate(reserved, EvaluationOptions()).IsValid);

        // Cleanup captureResponse rejected by schema
        var cleanupCapture = JsonNode.Parse("""
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SAMPLE-1",
              "name": "Sample",
              "steps": [
                { "id": "a", "operation": "listwindows", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "c", "operation": "quit", "arguments": {}, "captureResponse": true }
              ]
            }
            """);
        Assert.False(schema.Evaluate(cleanupCapture, EvaluationOptions()).IsValid);
    }

    [Fact]
    public void PlanSchema_DocumentsCsharpAsAuthoritativeForCombinedStepLimit()
    {
        var schemaText = ReadRepoFile("automation/schemas/plan.schema.json");
        Assert.Contains("AUTHORITATIVE combined step limit", schemaText);
        Assert.Contains("PlanValidator", schemaText);
        Assert.Contains("(steps + onFailureSteps) <= 1000", schemaText);
    }

    [Fact]
    public void ExamplePlan_ValidatesAgainstDraft202012Schema()
    {
        var schema = JsonSchema.FromText(ReadRepoFile("automation/schemas/plan.schema.json"));
        var example = JsonNode.Parse(ReadRepoFile("automation/plans/example.plan.json"));
        var result = schema.Evaluate(example, EvaluationOptions());
        Assert.True(result.IsValid, "example.plan.json must satisfy plan.schema.json");
    }

    private static EvaluationOptions EvaluationOptions() => new()
    {
        OutputFormat = OutputFormat.List,
        RequireFormatValidation = true
    };

    private static string ReadRepoFile(string relativePath)
    {
        var candidates = new[]
        {
            Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), relativePath)),
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", relativePath)),
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", relativePath))
        };
        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
                return File.ReadAllText(candidate);
        }

        throw new FileNotFoundException($"Unable to locate {relativePath}");
    }
}
