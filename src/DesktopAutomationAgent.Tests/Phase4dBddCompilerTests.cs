using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
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

        // Knowledge navigation
        var knowledgeDir = WriteApprovedKnowledge(temp.Path, stateId: "st_home");
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-6.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-6",
              "revision": "1",
              "summary": "nav",
              "feature": "Feature: N\n  Scenario: S\n    When I navigate using approved knowledge state \"st_home\"\n"
            }
            """);
        var nav = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-6",
            KnowledgePath = knowledgeDir
        });
        Assert.True(nav.Success, string.Join("; ", nav.Errors));
        Assert.False(string.IsNullOrWhiteSpace(nav.Metadata!.ApprovedKnowledgeSha256));

        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-7.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-7",
              "revision": "1",
              "summary": "nav miss",
              "feature": "Feature: N\n  Scenario: S\n    When I navigate using approved knowledge state \"st_missing\"\n"
            }
            """);
        var navMiss = CreateService(temp.Path).CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-7",
            KnowledgePath = knowledgeDir
        });
        Assert.False(navMiss.Success);
        Assert.Contains(navMiss.Errors, e => e.Contains("knowledge-resolution", StringComparison.Ordinal));
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
        using var temp = TempWorkspace();
        SeedWorkspace(temp.Path);
        var service = CreateService(temp.Path);

        var baseline = service.CompileKey(new CompileKeyRequest
        {
            JiraKey = "DEMO-1",
            CompiledAtUtc = DateTimeOffset.Parse("2026-11-12T00:00:00Z")
        });
        Assert.True(baseline.Success, string.Join("; ", baseline.Errors));
        var fp1 = baseline.Fingerprint!;

        // Change revision => new fingerprint
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
        Assert.NotEqual(fp1, rev2.Fingerprint);

        // Change step registry => new fingerprint
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

    private static string WriteApprovedKnowledge(string root, string stateId)
    {
        var dir = Path.Combine(root, "knowledge", "approved", "demo-app", "fp1");
        Directory.CreateDirectory(dir);
        var knowledge = new ApplicationKnowledgeDocument
        {
            KnowledgeSchemaVersion = 1,
            KnowledgeId = "kn_demo",
            ApplicationId = "demo-app",
            BuildFingerprint = "fp1",
            Status = "approved",
            CreatedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            Fingerprint = new ApplicationFingerprint
            {
                ApplicationId = "demo-app",
                ExecutablePath = "C:/demo.exe",
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
                Windows = [new KnowledgeWindowNode { WindowId = "win_1", Title = "Demo", ClassName = "Form", Role = "window" }],
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
            Coverage = new KnowledgeCoverageReport(),
            Approval = new KnowledgeApproval
            {
                ApprovalSchemaVersion = 1,
                Reviewer = "tester",
                ReviewedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
                SourceKnowledgeHash = new string('d', 64),
                SourceKnowledgeId = "kn_demo",
                Decision = "approve",
                PromotesPageObjects = false
            }
        };

        var json = JsonSerializer.Serialize(knowledge, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(Path.Combine(dir, "application-knowledge.json"), json);
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
