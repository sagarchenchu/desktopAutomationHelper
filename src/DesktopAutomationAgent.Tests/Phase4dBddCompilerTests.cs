using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Json.Schema;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class Phase4dBddCompilerTests
{
    [Fact]
    public void Schemas_ValidateDemoFixtureAndRegistry()
    {
        var testCaseSchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/jira-test-case.schema.json"));
        var registrySchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/step-registry.schema.json"));
        var options = new EvaluationOptions { OutputFormat = OutputFormat.List, RequireFormatValidation = true };

        var demo = JsonNode.Parse(ReadRepoFile("automation/test-cases/examples/DEMO-1.json"));
        Assert.True(testCaseSchema.Evaluate(demo, options).IsValid);

        var registry = JsonNode.Parse(ReadRepoFile("automation/step-registry/registry.json"));
        Assert.True(registrySchema.Evaluate(registry, options).IsValid);
    }

    [Fact]
    public void DemoFixture_CompilesReproducibly_ToValidPhase2Plan_WithoutUi()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);

        var service = CreateService(temp.Path);
        var first = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(first.Success, string.Join("; ", first.Errors));
        Assert.False(first.CacheHit);
        Assert.False(string.IsNullOrWhiteSpace(first.Fingerprint));
        Assert.True(File.Exists(Path.Combine(temp.Path, first.PlanPath!.Replace('/', Path.DirectorySeparatorChar))));

        var planBytes1 = File.ReadAllBytes(Path.Combine(temp.Path, first.PlanPath!.Replace('/', Path.DirectorySeparatorChar)));
        var validation = new PlanValidator().Validate(
            JsonSerializer.Deserialize<PlanManifest>(planBytes1, CompilationJsonOptions())!,
            first.PlanPath!);
        Assert.True(validation.IsValid, string.Join("; ", validation.Errors));

        var second = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T01:00:00Z")
        });
        Assert.True(second.Success, string.Join("; ", second.Errors));
        Assert.True(second.CacheHit);
        Assert.Equal(first.Fingerprint, second.Fingerprint);
        Assert.Equal(first.PlanSha256, second.PlanSha256);

        var planBytes2 = File.ReadAllBytes(Path.Combine(temp.Path, second.PlanPath!.Replace('/', Path.DirectorySeparatorChar)));
        Assert.Equal(planBytes1, planBytes2);

        // Fresh compile into another workspace yields identical plan bytes.
        using var temp2 = TempWorkspace();
        SeedWorkspace(temp2.Path);
        var other = CreateService(temp2.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T02:00:00Z")
        });
        Assert.True(other.Success, string.Join("; ", other.Errors));
        var otherBytes = File.ReadAllBytes(Path.Combine(temp2.Path, other.PlanPath!.Replace('/', Path.DirectorySeparatorChar)));
        Assert.Equal(planBytes1, otherBytes);
    }

    [Fact]
    public void Unsupported_ScenarioOutline_FailsClosed()
    {
        var parser = new GherkinBddParser();
        var result = parser.ParseFeatureText(
            "Feature: X\n  Scenario Outline: Y\n    Given <n>\n    Examples:\n      | n |\n      | 1 |\n");
        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, e => e.Contains("unsupported-feature", StringComparison.Ordinal));
        Assert.Contains(result.Errors, e => e.Contains("Scenario Outline/Examples", StringComparison.Ordinal));
    }

    [Fact]
    public void Unsupported_DataTable_FailsClosed()
    {
        var parser = new GherkinBddParser();
        var result = parser.ParseFeatureText(
            "Feature: X\n  Scenario: Y\n    Given the following:\n      | a |\n      | 1 |\n");
        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, e => e.Contains("data tables", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void UnmappedAndAmbiguous_FailClosed()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-2.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-2",
              "revision": "1",
              "summary": "unmapped",
              "feature": "Feature: U\n  Scenario: S\n    Given something unknown\n"
            }
            """);

        var unmapped = CreateService(temp.Path).CompileKey(new CompileKeyRequest { JiraKey = "DEMO-2" });
        Assert.False(unmapped.Success);
        Assert.Contains(unmapped.Errors, e => e.Contains("unmapped-step", StringComparison.Ordinal));

        // Ambiguous: duplicate exact mappings for same keyword+text.
        var registryPath = Path.Combine(temp.Path, "step-registry", "registry.json");
        var registry = JsonNode.Parse(File.ReadAllText(registryPath))!.AsObject();
        var mappings = registry["mappings"]!.AsArray();
        var clone = JsonNode.Parse(mappings[0]!.ToJsonString())!.AsObject();
        clone["mappingId"] = "desktop-ready-dup";
        mappings.Add(clone);
        File.WriteAllText(registryPath, registry.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-3.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-3",
              "revision": "1",
              "summary": "ambiguous",
              "feature": "Feature: A\n  Scenario: S\n    Given the desktop is ready\n"
            }
            """);

        var ambiguous = CreateService(temp.Path).CompileKey(new CompileKeyRequest { JiraKey = "DEMO-3" });
        Assert.False(ambiguous.Success);
        Assert.Contains(ambiguous.Errors, e => e.Contains("ambiguous-step", StringComparison.Ordinal));
    }

    [Fact]
    public void ParameterEscapingAndTypes_AreValidated()
    {
        var mapping = new StepMapping
        {
            MappingId = "int-param",
            Pattern = new StepPattern { Kind = "exact", Value = "count is {n}" },
            Keywords = ["Given"],
            Parameters = [new StepParameter { Name = "n", Type = "int" }],
            PlanSteps =
            [
                new PlanStepTemplate
                {
                    IdTemplate = "count-{n}",
                    Operation = "listwindows",
                    Arguments = new Dictionary<string, JsonElement>
                    {
                        ["includeDesktopDescendants"] = JsonSerializer.SerializeToElement(false)
                    }
                }
            ],
            Risk = "low",
            Source = "test",
            Version = "1",
            Approval = new StepMappingApproval
            {
                ApprovedBy = "t",
                ApprovedAt = DateTimeOffset.Parse("2026-01-01T00:00:00Z")
            }
        };
        var errors = new List<string>();
        StepRegistryValidator.CompilePattern(mapping, "m", errors);
        Assert.Empty(errors);

        var matcher = new StepMatcher(new StepRegistryDocument
        {
            SchemaVersion = 1,
            RegistryId = "t",
            Version = "1",
            Mappings = [mapping]
        });

        var bad = matcher.Match(new BddIrStep
        {
            RawKeyword = "Given",
            Keyword = "Given",
            Text = "count is abc",
            Line = 1,
            Ordinal = 0
        });
        Assert.False(bad.IsSuccess);

        var good = matcher.Match(new BddIrStep
        {
            RawKeyword = "Given",
            Keyword = "Given",
            Text = "count is 42",
            Line = 1,
            Ordinal = 0
        });
        Assert.True(good.IsSuccess);
        Assert.Equal("42", good.Match!.Parameters["n"]);
    }

    [Fact]
    public void MaliciousAndTimeoutRegex_AreRejectedOrBounded()
    {
        // Unanchored regex is rejected at registry validation time.
        var unanchored = new StepMapping
        {
            MappingId = "evil-unanchored",
            Pattern = new StepPattern { Kind = "regex", Value = "(a+)+" },
            Keywords = ["Given"],
            Parameters = [],
            PlanSteps =
            [
                new PlanStepTemplate
                {
                    IdTemplate = "x",
                    Operation = "listwindows",
                    Arguments = new Dictionary<string, JsonElement>()
                }
            ],
            Risk = "low",
            Source = "test",
            Version = "1",
            Approval = new StepMappingApproval
            {
                ApprovedBy = "t",
                ApprovedAt = DateTimeOffset.Parse("2026-01-01T00:00:00Z")
            }
        };
        var anchorErrors = new List<string>();
        StepRegistryValidator.CompilePattern(unanchored, "m", anchorErrors);
        Assert.Contains(anchorErrors, e => e.Contains("anchored", StringComparison.OrdinalIgnoreCase));

        // Anchored catastrophic pattern: match timed out or fails closed (never hangs / never maps).
        var mapping = new StepMapping
        {
            MappingId = "evil",
            Pattern = new StepPattern
            {
                Kind = "regex",
                Value = "^(a+)+$"
            },
            Keywords = ["Given"],
            Parameters = [],
            PlanSteps =
            [
                new PlanStepTemplate
                {
                    IdTemplate = "x",
                    Operation = "listwindows",
                    Arguments = new Dictionary<string, JsonElement>()
                }
            ],
            Risk = "low",
            Source = "test",
            Version = "1",
            Approval = new StepMappingApproval
            {
                ApprovedBy = "t",
                ApprovedAt = DateTimeOffset.Parse("2026-01-01T00:00:00Z")
            }
        };

        var errors = new List<string>();
        StepRegistryValidator.CompilePattern(mapping, "m", errors);
        Assert.Empty(errors);

        var matcher = new StepMatcher(new StepRegistryDocument
        {
            SchemaVersion = 1,
            RegistryId = "t",
            Version = "1",
            Mappings = [mapping]
        });

        var attack = new string('a', 40) + "!";
        var result = matcher.Match(new BddIrStep
        {
            RawKeyword = "Given",
            Keyword = "Given",
            Text = attack,
            Line = 1,
            Ordinal = 0
        });
        Assert.False(result.IsSuccess);
        Assert.True(
            result.Errors.Any(e =>
                e.Contains("timeout", StringComparison.OrdinalIgnoreCase)
                || e.Contains("unmapped-step", StringComparison.Ordinal)));
    }

    [Fact]
    public void ObjectAndKnowledgeResolution_FailClosed_AndSucceedWhenApproved()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        WriteLoginPage(temp.Path);

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-4.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-4",
              "revision": "1",
              "summary": "click object",
              "feature": "Feature: C\n  Scenario: S\n    When I click object login.submit\n"
            }
            """);

        var ok = CreateService(temp.Path).CompileKey(new CompileKeyRequest { JiraKey = "DEMO-4" });
        Assert.True(ok.Success, string.Join("; ", ok.Errors));

        // Missing object
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-5.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-5",
              "revision": "1",
              "summary": "missing object",
              "feature": "Feature: C\n  Scenario: S\n    When I click object login.missing\n"
            }
            """);
        var missing = CreateService(temp.Path).CompileKey(new CompileKeyRequest { JiraKey = "DEMO-5" });
        Assert.False(missing.Success);
        Assert.Contains(missing.Errors, e => e.Contains("object-resolution", StringComparison.Ordinal));

        // Knowledge expectedState resolution (navigation role is rejected separately).
        var knowledgeDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "require state",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_home\"\n"
            }
            """);
        var require = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = knowledgeDir
        });
        Assert.True(require.Success, string.Join("; ", require.Errors));
        Assert.False(string.IsNullOrWhiteSpace(require.Metadata!.ApprovedKnowledgeSha256));

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-7.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-7",
              "revision": "1",
              "summary": "state miss",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_missing\"\n"
            }
            """);
        var miss = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-7",
            KnowledgePath = knowledgeDir
        });
        Assert.False(miss.Success);
        Assert.Contains(miss.Errors, e => e.Contains("knowledge-resolution", StringComparison.Ordinal));

        // approved-rogue must not pass containment (StartsWith would falsely allow it).
        var rogue = Path.Combine(temp.Path, "knowledge", "approved-rogue", "demo-app", "fp1");
        Directory.CreateDirectory(rogue);
        foreach (var file in Directory.GetFiles(knowledgeDir))
            File.Copy(file, Path.Combine(rogue, Path.GetFileName(file)), overwrite: true);
        var rogueResult = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = rogue
        });
        Assert.False(rogueResult.Success);
        Assert.Contains(rogueResult.Errors, e =>
            e.Contains("knowledge/approved", StringComparison.OrdinalIgnoreCase)
            || e.Contains("containment", StringComparison.OrdinalIgnoreCase)
            || e.Contains("outside the allowed directory", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void NavigationRole_IsRejected_WithoutStub()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var knowledgeDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-9.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-9",
              "revision": "1",
              "summary": "nav rejected",
              "feature": "Feature: N\n  Scenario: S\n    When I navigate using approved knowledge state \"st_home\"\n"
            }
            """);
        var result = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-9",
            KnowledgePath = knowledgeDir
        });
        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("unsupported-feature", StringComparison.Ordinal)
                                            && e.Contains("navigation", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ProhibitedRisk_IsRejected()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-8.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-8",
              "revision": "1",
              "summary": "risky",
              "feature": "Feature: R\n  Scenario: S\n    When I delete all records\n"
            }
            """);

        var result = CreateService(temp.Path).CompileKey(new CompileKeyRequest { JiraKey = "DEMO-8" });
        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("prohibited-risk", StringComparison.Ordinal));
    }

    [Fact]
    public void CacheInvalidation_PerFingerprintComponent()
    {
        var baseline = BaseFingerprintInput();
        var baseFp = CompilationFingerprint.Compute(baseline);

        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { InputSha256 = new string('1', 64) }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { JiraRevision = "2" }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { CompilerVersion = "phase4d-2" }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { ParserVersion = "gherkin-99" }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { StepRegistrySha256 = new string('2', 64) }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { ObjectRepositorySha256 = new string('3', 64) }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { ObjectRepositoryPath = "object-repository/other.json" }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { ApprovedKnowledgeSha256 = new string('4', 64) }));
        Assert.NotEqual(baseFp, CompilationFingerprint.Compute(baseline with { CatalogFingerprint = new string('5', 64) }));
        Assert.NotEqual(
            baseFp,
            CompilationFingerprint.Compute(baseline with
            {
                ParameterBindingsSha256 = CompilationFingerprint.ComputeParameterBindingsSha256(
                    new Dictionary<string, string> { ["x"] = "1" })
            }));

        // Integration: mutable workspace inputs also invalidate.
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var service = CreateService(temp.Path);

        var first = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(first.Success, string.Join("; ", first.Errors));

        var demoPath = Path.Combine(temp.Path, "test-cases", "examples", "DEMO-1.json");
        var demo = JsonNode.Parse(File.ReadAllText(demoPath))!.AsObject();
        demo["revision"] = "2";
        File.WriteAllText(demoPath, demo.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var rev2 = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(rev2.Success, string.Join("; ", rev2.Errors));
        Assert.NotEqual(first.Fingerprint, rev2.Fingerprint);

        var registryPath = Path.Combine(temp.Path, "step-registry", "registry.json");
        var registry = JsonNode.Parse(File.ReadAllText(registryPath))!.AsObject();
        registry["version"] = "1.0.1";
        File.WriteAllText(registryPath, registry.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var reg2 = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(reg2.Success, string.Join("; ", reg2.Errors));
        Assert.NotEqual(rev2.Fingerprint, reg2.Fingerprint);

        var withBindings = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            ParameterBindings = new Dictionary<string, string> { ["extra"] = "value" },
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(withBindings.Success, string.Join("; ", withBindings.Errors));
        Assert.NotEqual(reg2.Fingerprint, withBindings.Fingerprint);
    }

    [Fact]
    public void ValidateCompiledKey_RejectsBadFingerprintAndDrift()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));

        var ok = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.True(ok.IsValid, string.Join("; ", ok.Errors));

        var traversal = service.ValidateCompiledKey("DEMO-1", "../" + compiled.Fingerprint);
        Assert.False(traversal.IsValid);
        Assert.Contains(traversal.Errors, e => e.Contains("hex", StringComparison.OrdinalIgnoreCase)
                                               || e.Contains("fingerprint", StringComparison.OrdinalIgnoreCase));

        // Dependency drift after registry edit.
        var registryPath = Path.Combine(temp.Path, "step-registry", "registry.json");
        var registry = JsonNode.Parse(File.ReadAllText(registryPath))!.AsObject();
        registry["version"] = "9.9.9";
        File.WriteAllText(registryPath, registry.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var drift = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(drift.IsValid);
        Assert.Contains(drift.Errors, e => e.Contains("dependency-drift", StringComparison.Ordinal));
    }

    [Fact]
    public void CustomRegistry_IsValidated_Recorded_AndDriftChecked()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);

        var customRelative = "step-registry/custom-registry.json";
        var customFull = Path.Combine(temp.Path, "step-registry", "custom-registry.json");
        File.Copy(Path.Combine(temp.Path, "step-registry", "registry.json"), customFull, overwrite: true);

        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            RegistryPath = customRelative,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));
        Assert.Equal(customRelative, compiled.Metadata!.StepRegistryPath);

        // Editing the default registry must not false-drift a custom-registry compile.
        var defaultRegistry = Path.Combine(temp.Path, "step-registry", "registry.json");
        var defaultNode = JsonNode.Parse(File.ReadAllText(defaultRegistry))!.AsObject();
        defaultNode["version"] = "0.0.0-default-changed";
        File.WriteAllText(defaultRegistry, defaultNode.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var stillOk = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.True(stillOk.IsValid, string.Join("; ", stillOk.Errors));

        // Editing the recorded custom registry must drift.
        var customNode = JsonNode.Parse(File.ReadAllText(customFull))!.AsObject();
        customNode["version"] = "9.9.9-custom";
        File.WriteAllText(customFull, customNode.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var drift = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(drift.IsValid);
        Assert.Contains(drift.Errors, e => e.Contains("dependency-drift", StringComparison.Ordinal)
                                           && e.Contains("step-registry", StringComparison.Ordinal));
    }

    [Fact]
    public void KnowledgeAndCatalogPaths_AreRecorded_AndDriftChecked()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var knowledgeDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        var catalogRelative = "catalogs/ops.json";
        var catalogFull = Path.Combine(temp.Path, "catalogs", "ops.json");
        Directory.CreateDirectory(Path.GetDirectoryName(catalogFull)!);
        var catalog = MinimalCatalog();
        File.WriteAllText(catalogFull, JsonSerializer.Serialize(catalog, CompilationJsonOptions()));

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "require state",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_home\"\n"
            }
            """);

        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = knowledgeDir,
            CatalogPath = catalogRelative,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));
        Assert.False(string.IsNullOrWhiteSpace(compiled.Metadata!.ApprovedKnowledgePath));
        Assert.Equal(catalogRelative, compiled.Metadata.CatalogPath);
        Assert.False(string.IsNullOrWhiteSpace(compiled.Metadata.CatalogFingerprint));

        var ok = service.ValidateCompiledKey("DEMO-6", compiled.Fingerprint);
        Assert.True(ok.IsValid, string.Join("; ", ok.Errors));

        File.WriteAllText(
            Path.Combine(knowledgeDir, "application-knowledge.json"),
            File.ReadAllText(Path.Combine(knowledgeDir, "application-knowledge.json")) + "\n");
        var knowledgeDrift = service.ValidateCompiledKey("DEMO-6", compiled.Fingerprint);
        Assert.False(knowledgeDrift.IsValid);
        Assert.Contains(knowledgeDrift.Errors, e => e.Contains("approved knowledge", StringComparison.OrdinalIgnoreCase));

        // Restore knowledge bytes for catalog drift isolation.
        // Recompile fresh for catalog-only drift.
        using var temp2 = TempWorkspace();
        SeedWorkspace(temp2.Path);
        var knowledgeDir2 = WriteApprovedKnowledge(temp2.Path, stateId: "st_home");
        var catalogFull2 = Path.Combine(temp2.Path, "catalogs", "ops.json");
        Directory.CreateDirectory(Path.GetDirectoryName(catalogFull2)!);
        File.WriteAllText(catalogFull2, JsonSerializer.Serialize(MinimalCatalog(), CompilationJsonOptions()));
        File.WriteAllText(Path.Combine(temp2.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "require state",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_home\"\n"
            }
            """);
        var service2 = CreateService(temp2.Path);
        var compiled2 = service2.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = knowledgeDir2,
            CatalogPath = catalogRelative,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled2.Success, string.Join("; ", compiled2.Errors));

        catalog.Operations[0].Aliases.Add("listwindows-alias");
        File.WriteAllText(catalogFull2, JsonSerializer.Serialize(catalog, CompilationJsonOptions()));
        var catalogDrift = service2.ValidateCompiledKey("DEMO-6", compiled2.Fingerprint);
        Assert.False(catalogDrift.IsValid);
        Assert.Contains(catalogDrift.Errors, e => e.Contains("catalog fingerprint", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void CatalogFingerprint_IncludesAliasesDeprecatedInputsAndSession()
    {
        var baseline = MinimalCatalog();
        var baseFp = CompilationFingerprint.ComputeCatalogFingerprint(baseline)!;

        var aliases = MinimalCatalog();
        aliases.Operations[0].Aliases.Add("extra-alias");
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(aliases));

        var deprecated = MinimalCatalog();
        deprecated.Operations[0].Deprecated = true;
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(deprecated));

        var deprecatedAliases = MinimalCatalog();
        deprecatedAliases.Operations[0].DeprecatedAliases.Add("old-name");
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(deprecatedAliases));

        var required = MinimalCatalog();
        required.Operations[0].RequiredInputs.Add("timeoutMs");
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(required));

        var alternatives = MinimalCatalog();
        alternatives.Operations[0].RequiredInputAlternatives.Add(["automationId", "name"]);
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(alternatives));

        var session = MinimalCatalog();
        session.Operations[0].RequiresSession = !session.Operations[0].RequiresSession;
        Assert.NotEqual(baseFp, CompilationFingerprint.ComputeCatalogFingerprint(session));
    }

    [Fact]
    public void MissingOrNullStepRegistryPath_FailsClosed_WithoutDefaultFallback()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));
        Assert.False(string.IsNullOrWhiteSpace(compiled.Metadata!.StepRegistryPath));

        var artifactDir = Path.Combine(temp.Path, "compiled", "DEMO-1", compiled.Fingerprint!);
        var metaPath = Path.Combine(artifactDir, "compilation.json");
        var planBytes = File.ReadAllBytes(Path.Combine(artifactDir, "plan.json"));
        var meta = JsonNode.Parse(File.ReadAllText(metaPath))!.AsObject();

        meta["stepRegistryPath"] = null;
        File.WriteAllText(metaPath, meta.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllBytes(Path.Combine(artifactDir, "plan.json"), planBytes);
        var nullPath = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(nullPath.IsValid);
        Assert.Contains(nullPath.Errors, e => e.Contains("stepRegistryPath", StringComparison.OrdinalIgnoreCase));

        meta.Remove("stepRegistryPath");
        File.WriteAllText(metaPath, meta.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var missingPath = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(missingPath.IsValid);
        Assert.Contains(missingPath.Errors, e => e.Contains("stepRegistryPath", StringComparison.OrdinalIgnoreCase));

        meta["stepRegistryPath"] = "   ";
        File.WriteAllText(metaPath, meta.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        var blankPath = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(blankPath.IsValid);
        Assert.Contains(blankPath.Errors, e => e.Contains("stepRegistryPath", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AbsoluteRegistryCatalogAndObjectRepoPaths_ResolveThenSealRelative()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);

        var catalogRelative = "catalogs/ops.json";
        var catalogFull = Path.Combine(temp.Path, "catalogs", "ops.json");
        Directory.CreateDirectory(Path.GetDirectoryName(catalogFull)!);
        File.WriteAllText(catalogFull, JsonSerializer.Serialize(MinimalCatalog(), CompilationJsonOptions()));

        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            RegistryPath = Path.Combine(temp.Path, "step-registry", "registry.json"),
            ObjectRepositoryPath = Path.Combine(temp.Path, "object-repository", "repository.json"),
            CatalogPath = catalogFull,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));
        Assert.Equal("step-registry/registry.json", compiled.Metadata!.StepRegistryPath);
        Assert.Equal("object-repository/repository.json", compiled.Metadata.ObjectRepositoryPath);
        Assert.Equal(catalogRelative, compiled.Metadata.CatalogPath);

        var ok = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.True(ok.IsValid, string.Join("; ", ok.Errors));
    }

    [Fact]
    public void KnowledgeFile_IsSymlinkChecked_BeforeHashing()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var knowledgeDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "require state",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_home\"\n"
            }
            """);

        var approvedRoot = KnowledgePathSafety.ApprovedRoot(temp.Path);
        var knowledgeFile = Path.Combine(knowledgeDir, "application-knowledge.json");

        // ValidateDirectory + file containment must succeed before a hash is usable.
        var validation = new KnowledgeValidator(Path.Combine(temp.Path, "schemas"))
            .ValidateDirectory(knowledgeDir, approvedRoot);
        Assert.True(validation.IsValid, string.Join("; ", validation.Errors));
        KnowledgePathSafety.EnsureReadablePathInside(knowledgeFile, approvedRoot);

        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = knowledgeDir,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));
        Assert.False(string.IsNullOrWhiteSpace(compiled.Metadata!.ApprovedKnowledgeSha256));

        // Drift path also ValidateDirectory's before hashing: break a sibling so validate fails closed.
        File.WriteAllText(Path.Combine(knowledgeDir, "state-graph.json"), "{ \"tampered\": true }");
        var drift = service.ValidateCompiledKey("DEMO-6", compiled.Fingerprint);
        Assert.False(drift.IsValid);
        Assert.Contains(drift.Errors, e =>
            e.Contains("ValidateDirectory", StringComparison.OrdinalIgnoreCase)
            || e.Contains("Sibling", StringComparison.OrdinalIgnoreCase)
            || e.Contains("state-graph", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void MalformedCompilationJson_FailsClosed_BeforeFingerprintRecompute()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var service = CreateService(temp.Path);
        var compiled = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(compiled.Success, string.Join("; ", compiled.Errors));

        var artifactDir = Path.Combine(temp.Path, "compiled", "DEMO-1", compiled.Fingerprint!);
        var metaPath = Path.Combine(artifactDir, "compilation.json");
        var planPath = Path.Combine(artifactDir, "plan.json");

        // Corrupt JSON
        File.WriteAllText(metaPath, "{ not-json");
        var corrupt = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(corrupt.IsValid);
        Assert.Contains(corrupt.Errors, e => e.Contains("failed to read", StringComparison.OrdinalIgnoreCase)
                                             || e.Contains("compilation", StringComparison.OrdinalIgnoreCase));

        // Null required field — must not throw NRE via ToLowerInvariant.
        var planBytes = File.ReadAllBytes(planPath);
        File.WriteAllText(metaPath, """
            {
              "schemaVersion": "1",
              "jiraKey": "DEMO-1",
              "fingerprint": "REPLACE_FP",
              "compilerVersion": "phase4d-1",
              "parserVersion": "gherkin-30.1.2",
              "inputSha256": null,
              "jiraRevision": "1",
              "stepRegistrySha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
              "objectRepositorySha256": "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
              "objectRepositoryPath": "object-repository/repository.json",
              "parameterBindingsSha256": "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc",
              "planSha256": "dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd"
            }
            """.Replace("REPLACE_FP", compiled.Fingerprint!, StringComparison.Ordinal));
        // Keep plan present so TryReadExisting reaches metadata integrity.
        File.WriteAllBytes(planPath, planBytes);
        var nullField = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(nullField.IsValid);
        Assert.Contains(nullField.Errors, e => e.Contains("inputSha256", StringComparison.OrdinalIgnoreCase));

        // Missing required field
        File.WriteAllText(metaPath, $$"""
            {
              "schemaVersion": "1",
              "jiraKey": "DEMO-1",
              "fingerprint": "{{compiled.Fingerprint}}",
              "compilerVersion": "phase4d-1",
              "parserVersion": "gherkin-30.1.2",
              "jiraRevision": "1",
              "stepRegistrySha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
              "objectRepositorySha256": "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
              "objectRepositoryPath": "object-repository/repository.json",
              "parameterBindingsSha256": "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc",
              "planSha256": "dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd"
            }
            """);
        var missing = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(missing.IsValid);
        Assert.Contains(missing.Errors, e => e.Contains("inputSha256", StringComparison.OrdinalIgnoreCase));

        // Partial artifact (plan without compilation)
        File.Delete(metaPath);
        var partial = service.ValidateCompiledKey("DEMO-1", compiled.Fingerprint);
        Assert.False(partial.IsValid);
        Assert.Contains(partial.Errors, e => e.Contains("incomplete", StringComparison.OrdinalIgnoreCase)
                                             || e.Contains("partial", StringComparison.OrdinalIgnoreCase)
                                             || e.Contains("corrupt", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void KnowledgeDirectory_JunctionEscape_IsRejected_SafeInsideJunction_Allowed()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var realDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        var approvedRoot = Path.Combine(temp.Path, "knowledge", "approved");
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "require state",
              "feature": "Feature: N\n  Scenario: S\n    Given I require approved knowledge state \"st_home\"\n"
            }
            """);

        // Escape: junction under approved that points outside approved.
        var outside = Path.Combine(temp.Path, "outside-knowledge");
        Directory.CreateDirectory(outside);
        foreach (var file in Directory.GetFiles(realDir))
            File.Copy(file, Path.Combine(outside, Path.GetFileName(file)), overwrite: true);

        var escapeLink = Path.Combine(approvedRoot, "escape-junc");
        if (!TryCreateDirectoryJunction(escapeLink, outside))
            return; // environment cannot create junctions

        var escapeResult = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = escapeLink
        });
        Assert.False(escapeResult.Success);
        Assert.Contains(escapeResult.Errors, e =>
            e.Contains("outside", StringComparison.OrdinalIgnoreCase)
            || e.Contains("containment", StringComparison.OrdinalIgnoreCase)
            || e.Contains("knowledge/approved", StringComparison.OrdinalIgnoreCase)
            || e.Contains("directory link", StringComparison.OrdinalIgnoreCase));

        // Safe-inside: junction under approved pointing at another approved directory.
        var safeLink = Path.Combine(approvedRoot, "safe-junc");
        Assert.True(TryCreateDirectoryJunction(safeLink, realDir));
        var safe = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = safeLink,
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(safe.Success, string.Join("; ", safe.Errors));
    }

    [Fact]
    public void CompileKey_Cli_NeverPostsUi_AndValidateCompiledKeyPasses()
    {
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);

        var exit = AgentCli.RunAsync(
            [
                "compile-key",
                "--key", "DEMO-1",
                "--json",
                "--Workspace:Root", temp.Path
            ],
            (configArgs, json) => AgentCli.BuildHost(configArgs, json)).GetAwaiter().GetResult();
        Assert.Equal(0, exit);

        var validateExit = AgentCli.RunAsync(
            [
                "validate-compiled-key",
                "--key", "DEMO-1",
                "--json",
                "--Workspace:Root", temp.Path
            ],
            (configArgs, json) => AgentCli.BuildHost(configArgs, json)).GetAwaiter().GetResult();
        Assert.Equal(0, validateExit);
    }

    [Fact]
    public void AndBut_Normalize_ToPreviousPrimaryKeyword()
    {
        var parser = new GherkinBddParser();
        var result = parser.ParseFeatureText(
            "Feature: X\n  Scenario: Y\n    Given the desktop is ready\n    And the desktop is ready\n    When I list visible windows\n    But I list visible windows\n");
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors));
        Assert.Equal(["Given", "Given", "When", "When"], result.Document!.Scenarios[0].Steps.Select(s => s.Keyword));
    }

    [Fact]
    public void WorkspaceInit_CreatesPhase4dLayout()
    {
        using var temp = TempWorkspace();
        var workspace = new WorkspaceManager(
            Options.Create(TestSupport.CreateOptions(workspaceRoot: temp.Path)),
            NullLogger<WorkspaceManager>.Instance);
        var result = workspace.Initialize();
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "compiled/");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "step-registry/");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "test-cases/");
        Assert.True(File.Exists(Path.Combine(temp.Path, "schemas", "jira-test-case.schema.json")));
        Assert.True(File.Exists(Path.Combine(temp.Path, "schemas", "step-registry.schema.json")));
    }

    private static BddCompilationService CreateService(string root)
    {
        var options = TestSupport.CreateOptions(workspaceRoot: root);
        var workspace = new WorkspaceManager(Options.Create(options), NullLogger<WorkspaceManager>.Instance);
        var reader = new ObjectRepositoryReader(Options.Create(options), workspace);
        return new BddCompilationService(workspace, reader, Options.Create(options));
    }

    private static void SeedWorkspace(string root)
    {
        var workspace = new WorkspaceManager(
            Options.Create(TestSupport.CreateOptions(workspaceRoot: root)),
            NullLogger<WorkspaceManager>.Instance);
        workspace.Initialize();

        Directory.CreateDirectory(Path.Combine(root, "step-registry"));
        Directory.CreateDirectory(Path.Combine(root, "test-cases", "examples"));
        File.Copy(
            Path.Combine(RepoRoot(), "automation", "step-registry", "registry.json"),
            Path.Combine(root, "step-registry", "registry.json"),
            overwrite: true);
        File.Copy(
            Path.Combine(RepoRoot(), "automation", "test-cases", "examples", "DEMO-1.json"),
            Path.Combine(root, "test-cases", "examples", "DEMO-1.json"),
            overwrite: true);
    }

    private static void WriteLoginPage(string root)
    {
        var pages = Path.Combine(root, "object-repository", "pages");
        Directory.CreateDirectory(pages);
        File.WriteAllText(Path.Combine(root, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Default object repository",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(pages, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "description": "Submit",
                  "locator": { "automationId": "SubmitButton" }
                }
              }
            }
            """);
    }

    private static CompilationFingerprintInput BaseFingerprintInput() =>
        new()
        {
            InputSha256 = new string('a', 64),
            JiraRevision = "1",
            CompilerVersion = BddCompilerVersions.CompilerVersion,
            ParserVersion = BddCompilerVersions.ParserVersion,
            StepRegistrySha256 = new string('b', 64),
            ObjectRepositorySha256 = new string('c', 64),
            ObjectRepositoryPath = "object-repository/repository.json",
            ApprovedKnowledgeSha256 = null,
            CatalogFingerprint = null,
            ParameterBindingsSha256 = CompilationFingerprint.ComputeParameterBindingsSha256(null)
        };

    private static OperationsCatalogDto MinimalCatalog() =>
        new()
        {
            SchemaVersion = 1,
            DriverVersion = "test-driver",
            Operations =
            [
                new OperationDescriptorDto
                {
                    Name = "listwindows",
                    Category = "query",
                    RequiresSession = false,
                    Deprecated = false,
                    Aliases = [],
                    DeprecatedAliases = [],
                    RequiredInputs = [],
                    RequiredInputAlternatives = []
                }
            ]
        };

    private static bool TryCreateDirectoryJunction(string linkPath, string targetPath)
    {
        try
        {
            // Non-recursive: remove the junction node without deleting the target tree.
            if (Directory.Exists(linkPath))
                Directory.Delete(linkPath, recursive: false);
            else if (File.Exists(linkPath))
                File.Delete(linkPath);
        }
        catch
        {
            // best-effort
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c mklink /J \"{linkPath}\" \"{targetPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var process = Process.Start(psi);
            if (process is null)
                return false;
            process.WaitForExit(10_000);
            return process.ExitCode == 0 && Directory.Exists(linkPath);
        }
        catch
        {
            return false;
        }
    }

    private static string WriteApprovedKnowledge(string root, string stateId)
    {
        var buildFingerprint = new string('b', 32);
        var dir = Path.Combine(root, "knowledge", "approved", "demo-app", buildFingerprint);
        Directory.CreateDirectory(dir);
        var knowledge = new ApplicationKnowledgeDocument
        {
            KnowledgeSchemaVersion = 1,
            KnowledgeId = "kn_demo",
            ApplicationId = "demo-app",
            BuildFingerprint = buildFingerprint,
            Status = "candidate",
            CreatedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            SourceExplorationRuns =
            [
                new SourceExplorationRun
                {
                    ExplorationId = "exp1",
                    ArtifactHash = new string('e', 64),
                    PolicySha256 = new string('b', 64)
                }
            ],
            Fingerprint = new ApplicationFingerprint
            {
                ApplicationId = "demo-app",
                ExecutablePath = "demo.exe",
                ExecutableSha256 = new string('a', 64),
                DriverVersion = "1.0.0",
                SnapshotSchemaVersion = 1,
                ExplorerVersion = "phase4b-1",
                ExplorationPolicyHash = new string('b', 64),
                ConfigIdentity = new Dictionary<string, string> { ["processName"] = "demo" }
            },
            Graph = new ApplicationStateGraph
            {
                GraphSchemaVersion = 1,
                Windows = [new KnowledgeWindowNode { WindowId = "win_1", Title = "Demo", ClassName = "Form", Role = "main" }],
                States =
                [
                    new KnowledgeStateNode
                    {
                        StateId = stateId,
                        Label = "Home",
                        WindowId = "win_1",
                        SemanticFingerprint = new string('c', 64)
                    }
                ],
                Elements = [],
                Transitions = []
            },
            Workflows = [],
            Coverage = new KnowledgeCoverageReport()
        };

        var contentHash = KnowledgeArtifactWriter.ComputeContentHash(knowledge);
        knowledge.Status = "approved";
        knowledge.Approval = new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "tester",
            ReviewedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            SourceKnowledgeHash = contentHash,
            SourceKnowledgeId = "kn_demo",
            Decision = "approve",
            PromotesPageObjects = false
        };

        new KnowledgeArtifactWriter().WriteApproved(dir, knowledge, overwrite: true);
        return dir;
    }

    private static TempDir TempWorkspace() => new();

    private static string ReadRepoFile(string relative) =>
        File.ReadAllText(Path.Combine(RepoRoot(), relative.Replace('/', Path.DirectorySeparatorChar)));

    private static string RepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.sln"))
                || File.Exists(Path.Combine(dir.FullName, "desktopAutomationHelper.sln")))
                return dir.FullName;
            if (Directory.Exists(Path.Combine(dir.FullName, "automation"))
                && Directory.Exists(Path.Combine(dir.FullName, "src")))
                return dir.FullName;
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root.");
    }

    private static JsonSerializerOptions CompilationJsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private sealed class TempDir : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "da-phase4d-" + Guid.NewGuid().ToString("N"));

        public TempDir() => Directory.CreateDirectory(Path);

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Path))
                    Directory.Delete(Path, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }
}
