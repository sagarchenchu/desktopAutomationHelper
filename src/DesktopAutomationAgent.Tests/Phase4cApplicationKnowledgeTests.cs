using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;
using Json.Schema;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class Phase4cApplicationKnowledgeTests
{
    [Fact]
    public void ApplicationKnowledgeSchemas_ValidateMinimalDocument()
    {
        var schemaText = ReadRepoFile("automation/schemas/application-knowledge.schema.json");
        Assert.Contains("draft/2020-12/schema", schemaText, StringComparison.Ordinal);

        // Graph/workflow/approval are referenced; validate graph + approval independently too.
        var graphSchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/application-state-graph.schema.json"));
        var approvalSchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/knowledge-approval.schema.json"));
        var workflowSchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/application-workflow.schema.json"));

        var graph = MinimalGraph();
        var graphNode = JsonNode.Parse(JsonSerializer.Serialize(graph, KnowledgeJsonOptions()));
        Assert.True(graphSchema.Evaluate(graphNode, EvalOptions()).IsValid);

        var approval = new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            SourceKnowledgeHash = new string('a', 64),
            SourceKnowledgeId = "kn_test",
            Decision = "approve",
            PromotesPageObjects = false
        };
        Assert.True(approvalSchema.Evaluate(JsonNode.Parse(JsonSerializer.Serialize(approval, KnowledgeJsonOptions())), EvalOptions()).IsValid);

        var workflow = new ApplicationWorkflow
        {
            WorkflowSchemaVersion = 1,
            WorkflowId = "wf_1",
            Name = "demo",
            EntryStateId = "st_1",
            Steps =
            [
                new KnowledgeWorkflowStep
                {
                    Ordinal = 0,
                    TransitionId = "tr_1",
                    Action = "invoke:Button|x",
                    SourceStateId = "st_1",
                    TargetStateId = "st_2",
                    Outcome = "success"
                }
            ],
            ObservationCount = 1
        };
        Assert.True(workflowSchema.Evaluate(JsonNode.Parse(JsonSerializer.Serialize(workflow, KnowledgeJsonOptions())), EvalOptions()).IsValid);
    }

    [Fact]
    public void WorkspaceInit_CreatesKnowledgeLayout_AndSchemaParity()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        var result = workspace.Initialize();
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "knowledge/");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "knowledge/README.md");
        Assert.True(File.Exists(Path.Combine(temp.Path, "schemas", "application-knowledge.schema.json")));

        foreach (var name in new[]
                 {
                     "application-knowledge.schema.json",
                     "application-state-graph.schema.json",
                     "application-workflow.schema.json",
                     "knowledge-approval.schema.json"
                 })
        {
            var generated = NormalizeNewlines(File.ReadAllText(Path.Combine(temp.Path, "schemas", name)));
            var expected = NormalizeNewlines(ReadRepoFile($"automation/schemas/{name}"));
            Assert.Equal(expected, generated);
        }
    }

    [Fact]
    public void StableIds_AreDeterministic_AndPreferAutomationIdOverName()
    {
        var left = new KnowledgeStableIdFactory();
        var right = new KnowledgeStableIdFactory();

        var a = left.ElementId(new KnowledgeElementIdentity
        {
            AutomationId = "OpenSecond",
            Name = "Open Second",
            ControlType = "Button"
        });
        var b = right.ElementId(new KnowledgeElementIdentity
        {
            AutomationId = "OpenSecond",
            Name = "Totally Different Label",
            ControlType = "Button"
        });
        Assert.Equal(a, b);

        var nameOnly = left.ElementId(new KnowledgeElementIdentity
        {
            Name = "Open Second",
            ControlType = "Button"
        });
        Assert.NotEqual(a, nameOnly);
    }

    [Fact]
    public void StableIds_ExcludeHwndAndOrdinalSuffixes_FromActions()
    {
        var normalized = KnowledgeStableIdFactory.NormalizeActionKey(
            "Invoke:Button|ExplorerOpenSecondButton|Open Second|o3|d1");
        Assert.DoesNotContain("|o3|d1", normalized, StringComparison.Ordinal);

        var stripped = KnowledgeStableIdFactory.NormalizeActionKey("backtrack:close:0xD02C8");
        Assert.Contains("{runtime}", stripped, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentHash_ExcludesVolatileTimestamps()
    {
        var doc = MinimalKnowledge();
        doc.Graph.States[0].FirstObservedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z");
        doc.Graph.States[0].LastObservedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z");
        var hash1 = KnowledgeArtifactWriter.ComputeContentHash(doc);

        doc.Graph.States[0].FirstObservedAtUtc = DateTimeOffset.Parse("2026-06-01T12:00:00Z");
        doc.Graph.States[0].LastObservedAtUtc = DateTimeOffset.Parse("2026-06-02T12:00:00Z");
        doc.CreatedAtUtc = DateTimeOffset.Parse("2026-06-02T12:00:00Z");
        var hash2 = KnowledgeArtifactWriter.ComputeContentHash(doc);

        Assert.Equal(hash1, hash2);
    }

    [Fact]
    public void Build_OrderIndependent_ProducesIdenticalGraphBytes()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();

        var policyHash = new string('b', 64);
        var runA = WriteExplorationRun(temp.Path, "run-aaa", policyHash, states: 2, withConflict: false);
        var runB = WriteExplorationRun(temp.Path, "run-bbb", policyHash, states: 2, withConflict: false);

        var service = CreateKnowledgeService(workspace);
        var first = service.Build("demo-app", [runA, runB], suggestPageObjects: false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(first.Success, first.Error);

        // Rebuild with reversed order into a fresh output path.
        var secondOut = Path.Combine(temp.Path, "knowledge", "candidates", "demo-app", "alt");
        Directory.CreateDirectory(secondOut);
        var second = service.Build(
            "demo-app",
            [runB, runA],
            suggestPageObjects: false,
            outputDirectory: secondOut,
            createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(second.Success, second.Error);

        var leftBytes = File.ReadAllBytes(Path.Combine(first.OutputDirectory!, "state-graph.json"));
        var rightBytes = File.ReadAllBytes(Path.Combine(second.OutputDirectory!, "state-graph.json"));
        Assert.Equal(leftBytes, rightBytes);

        Assert.Equal(
            KnowledgeArtifactWriter.ComputeContentHash(first.Knowledge!),
            KnowledgeArtifactWriter.ComputeContentHash(second.Knowledge!));
    }

    [Fact]
    public void Build_DuplicateEvidence_IncreasesObservationCount()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('c', 64);
        var runA = WriteExplorationRun(temp.Path, "dup-a", policyHash, states: 2, withConflict: false);
        var runB = WriteExplorationRun(temp.Path, "dup-b", policyHash, states: 2, withConflict: false);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [runA, runB], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.Contains(result.Knowledge!.Graph.Transitions, t => t.ObservationCount >= 2);
    }

    [Fact]
    public void Build_PreservesConflictingTransitions()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('d', 64);
        var run = WriteExplorationRun(temp.Path, "conflict-run", policyHash, states: 3, withConflict: true);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.True(result.Knowledge!.Coverage.Conflicts > 0);
        Assert.Contains(result.Knowledge.Graph.Transitions, t => !string.IsNullOrWhiteSpace(t.ConflictGroupId));
    }

    [Fact]
    public void Build_RejectsIncompatibleFingerprints()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var runA = WriteExplorationRun(temp.Path, "fp-a", new string('e', 64), states: 1, withConflict: false);
        var runB = WriteExplorationRun(temp.Path, "fp-b", new string('f', 64), states: 1, withConflict: false);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [runA, runB], false);
        Assert.False(result.Success);
        Assert.Contains("Incompatible application fingerprints", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public void MissingControls_DoNotDeletePriorKnowledge()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('1', 64);
        var rich = WriteExplorationRun(temp.Path, "rich-run", policyHash, states: 2, withConflict: false, includeExtraElement: true);
        var sparse = WriteExplorationRun(temp.Path, "sparse-run", policyHash, states: 2, withConflict: false, includeExtraElement: false);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [rich, sparse], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.Contains(
            result.Knowledge!.Graph.Elements,
            e => string.Equals(e.Identity.AutomationId, "ExtraButton", StringComparison.Ordinal));
    }

    [Fact]
    public void SuggestPageObjects_NeverTouchesActivePagesOrRepositoryJson()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();

        var repoPath = Path.Combine(temp.Path, "object-repository", "repository.json");
        var pagesDir = Path.Combine(temp.Path, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        var beforeRepo = File.ReadAllText(repoPath);
        var sentinel = Path.Combine(pagesDir, "must-not-touch.page.json");
        File.WriteAllText(sentinel, "{ \"state\": \"active\" }\n");

        var policyHash = new string('2', 64);
        var run = WriteExplorationRun(temp.Path, "po-run", policyHash, states: 2, withConflict: false);
        var result = CreateKnowledgeService(workspace).Build(
            "demo-app",
            [run],
            suggestPageObjects: true,
            createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.NotEmpty(result.SuggestedPageObjectPaths);
        Assert.All(result.SuggestedPageObjectPaths, p => Assert.StartsWith("object-repository/candidates/", p, StringComparison.Ordinal));

        Assert.Equal(beforeRepo, File.ReadAllText(repoPath));
        Assert.Equal("{ \"state\": \"active\" }\n", File.ReadAllText(sentinel));
        Assert.False(Directory.EnumerateFiles(pagesDir, "*.page.json")
            .Any(f => !string.Equals(f, sentinel, StringComparison.OrdinalIgnoreCase)
                      && File.GetLastWriteTimeUtc(f) > DateTime.UtcNow.AddMinutes(-5)
                      && !f.EndsWith("must-not-touch.page.json", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void ApproveKnowledge_RequiresHashAndReviewer_AndNeverPromotesPageObjects()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('3', 64);
        var run = WriteExplorationRun(temp.Path, "approve-run", policyHash, states: 2, withConflict: false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);

        var hash = KnowledgeArtifactWriter.ComputeContentHash(built.Knowledge!);
        var reviewed = Path.Combine(temp.Path, "reviewed-approve.json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            SourceKnowledgeHash = hash,
            SourceKnowledgeId = built.Knowledge!.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false,
            Notes = "sanitized review"
        }, KnowledgeJsonOptions()));

        var approved = service.Approve(built.OutputDirectory!, "alice", reviewed);
        Assert.True(approved.Success, approved.Error);
        Assert.Equal("approved", approved.Knowledge!.Status);
        Assert.False(approved.Knowledge.Approval!.PromotesPageObjects);

        // Wrong hash rejected.
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            SourceKnowledgeHash = new string('0', 64),
            SourceKnowledgeId = built.Knowledge.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false
        }, KnowledgeJsonOptions()));
        // Need a fresh candidate because approved path already written; rebuild.
        var built2 = service.Build(
            "demo-app",
            [run],
            false,
            outputDirectory: Path.Combine(temp.Path, "knowledge", "candidates", "demo-app", "retry"),
            createdAtUtc: DateTimeOffset.UnixEpoch);
        var bad = service.Approve(built2.OutputDirectory!, "alice", reviewed);
        Assert.False(bad.Success);
        Assert.Contains("sourceKnowledgeHash mismatch", bad.Error, StringComparison.Ordinal);
    }

    [Fact]
    public void ApproveKnowledge_RejectsPromotesPageObjectsTrue()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('4', 64);
        var run = WriteExplorationRun(temp.Path, "promote-run", policyHash, states: 1, withConflict: false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);
        var hash = KnowledgeArtifactWriter.ComputeContentHash(built.Knowledge!);
        var reviewed = Path.Combine(temp.Path, "bad-promote.json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "bob",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = hash,
            SourceKnowledgeId = built.Knowledge!.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = true
        }, KnowledgeJsonOptions()));

        var result = service.Approve(built.OutputDirectory!, "bob", reviewed);
        Assert.False(result.Success);
        Assert.Contains("promotesPageObjects", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SearchKnowledge_IsLexicalAndDeterministic()
    {
        var doc = MinimalKnowledge();
        doc.Graph.Elements.Add(new KnowledgeElementNode
        {
            ElementId = "el_open",
            Identity = new KnowledgeElementIdentity
            {
                AutomationId = "ExplorerOpenSecondButton",
                Name = "Open Second",
                ControlType = "Button"
            },
            Capabilities = ["Invoke"]
        });

        var searcher = new KnowledgeSearcher();
        var a = searcher.Search(doc, "Open Second");
        var b = searcher.Search(doc, "Open Second");
        Assert.Equal(
            JsonSerializer.Serialize(a),
            JsonSerializer.Serialize(b));
        Assert.Contains(a, h => h.Kind == "element" && h.Id == "el_open");
    }

    [Fact]
    public void CompareKnowledge_ReportsIncompatibleFingerprints()
    {
        var left = MinimalKnowledge();
        var right = MinimalKnowledge();
        right.Fingerprint.ExplorationPolicyHash = new string('9', 64);
        right.BuildFingerprint = "different";

        var compare = new KnowledgeComparer().Compare(left, right);
        Assert.False(compare.CompatibleFingerprints);
        Assert.Contains(compare.Differences, d => d.Contains("incompatible-fingerprints", StringComparison.Ordinal));
    }

    [Fact]
    public void PathSafety_RejectsTraversalOutsideWorkspace()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-knowledge-escape-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        try
        {
            var ex = Assert.Throws<KnowledgePathException>(() =>
                KnowledgePathSafety.EnsureWritablePathInside(
                    Path.Combine(outside, "application-knowledge.json"),
                    temp.Path));
            Assert.Contains("outside the allowed directory", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            try { Directory.Delete(outside, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void Cli_KnowledgeCommands_ParseAndStayOffline()
    {
        var build = CommandLine.Parse([
            "build-knowledge",
            "--application-id", "demo-app",
            "--exploration-run", "run-1",
            "--exploration-run", "run-2",
            "--suggest-page-objects",
            "--json"
        ]);
        Assert.Equal(AgentCommandKind.BuildKnowledge, build.Kind);
        Assert.Equal("demo-app", build.ApplicationId);
        Assert.Equal(["run-1", "run-2"], build.ExplorationRuns);
        Assert.True(build.SuggestPageObjects);
        Assert.True(build.Json);

        var approve = CommandLine.Parse([
            "approve-knowledge",
            "--source", "candidates/x",
            "--reviewer", "alice",
            "--reviewed-input", "review.json",
            "--json"
        ]);
        Assert.Equal(AgentCommandKind.ApproveKnowledge, approve.Kind);
        Assert.Contains("approve-knowledge", CommandLine.HelpText, StringComparison.Ordinal);
        Assert.Contains("never promotes Page Objects", CommandLine.HelpText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Cli_BuildKnowledge_EndToEnd_Offline()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('5', 64);
        var run = WriteExplorationRun(temp.Path, "cli-run", policyHash, states: 2, withConflict: false);

        var exit = AgentCli.RunAsync(
            [
                "build-knowledge",
                "--application-id", "demo-app",
                "--exploration-run", run,
                "--json",
                "--Workspace:Root", temp.Path
            ],
            (configArgs, json) => AgentCli.BuildHost(configArgs, json)).GetAwaiter().GetResult();

        Assert.Equal(0, exit);
    }

    [Fact]
    public void SizeLimit_RejectsOversizedExplorationArtifact()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var runDir = Path.Combine(temp.Path, "exploration", "runs", "huge-run");
        Directory.CreateDirectory(runDir);
        File.WriteAllText(Path.Combine(runDir, "manifest.json"), """{"manifestSchemaVersion":1,"explorationId":"huge-run","mode":"navigate","dryRun":false,"startedAtUtc":"2026-01-01T00:00:00Z","policySha256":"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"}""");
        File.WriteAllText(Path.Combine(runDir, "summary.json"), """{"summarySchemaVersion":1,"explorationId":"huge-run","status":"ok","exitCode":0,"mode":"navigate","dryRun":false,"startedAtUtc":"2026-01-01T00:00:00Z","completedAtUtc":"2026-01-01T00:00:01Z","statesVisited":0,"actionsAttempted":0,"transitionsRecorded":0,"maxDepthReached":0}""");
        File.WriteAllText(Path.Combine(runDir, "states.jsonl"), "");
        File.WriteAllText(Path.Combine(runDir, "transitions.jsonl"), new string('x', 2000));

        var options = Options.Create(new AgentOptions
        {
            Workspace = new WorkspaceOptions { Root = temp.Path },
            Knowledge = new KnowledgeOptions { MaxFileBytes = 100 }
        });
        var service = new KnowledgeService(workspace, new ObjectArtifactWriter(), options);
        var result = service.Build("demo-app", ["huge-run"], false);
        Assert.False(result.Success);
        Assert.Contains("maximum size", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    private static ApplicationKnowledgeDocument MinimalKnowledge()
    {
        var fp = new ApplicationFingerprint
        {
            ApplicationId = "demo-app",
            SnapshotSchemaVersion = 1,
            ExplorerVersion = KnowledgeConstants.ExplorerVersion,
            ExplorationPolicyHash = new string('a', 64)
        };
        return new ApplicationKnowledgeDocument
        {
            KnowledgeSchemaVersion = 1,
            KnowledgeId = "kn_test",
            ApplicationId = "demo-app",
            BuildFingerprint = new string('b', 32),
            Fingerprint = fp,
            Status = "candidate",
            CreatedAtUtc = DateTimeOffset.UnixEpoch,
            SourceExplorationRuns =
            [
                new SourceExplorationRun
                {
                    ExplorationId = "run-1",
                    ArtifactHash = new string('c', 64),
                    PolicySha256 = new string('a', 64)
                }
            ],
            Graph = MinimalGraph(),
            Coverage = new KnowledgeCoverageReport()
        };
    }

    private static ApplicationStateGraph MinimalGraph() =>
        new()
        {
            GraphSchemaVersion = 1,
            Windows =
            [
                new KnowledgeWindowNode
                {
                    WindowId = "win_1",
                    Title = "Demo",
                    IsModal = false,
                    Role = "main",
                    ObservationCount = 1
                }
            ],
            States =
            [
                new KnowledgeStateNode
                {
                    StateId = "st_1",
                    SemanticFingerprint = new string('1', 64),
                    WindowId = "win_1",
                    ObservationCount = 1
                }
            ],
            Elements = [],
            Transitions =
            [
                new KnowledgeTransitionEdge
                {
                    TransitionId = "tr_1",
                    Action = "invoke:Button|x",
                    ActionCategory = "invoke",
                    SourceStateId = "st_1",
                    TargetStateId = "st_1",
                    Outcome = "noOp",
                    ObservationCount = 1,
                    Evidence = [new KnowledgeTransitionEvidence { ExplorationId = "run-1" }]
                }
            ],
            EntryStateIds = ["st_1"],
            TerminalStateIds = ["st_1"],
            UnresolvedElements = [],
            AmbiguousElements = []
        };

    private static string WriteExplorationRun(
        string workspaceRoot,
        string explorationId,
        string policySha256,
        int states,
        bool withConflict,
        bool includeExtraElement = false,
        bool sharedStateFingerprints = true)
    {
        var runDir = Path.Combine(workspaceRoot, "exploration", "runs", explorationId);
        Directory.CreateDirectory(runDir);

        // Shared fingerprints across runs enable duplicate-evidence merge tests.
        // Conflict tests can still opt into unique fingerprints via sharedStateFingerprints=false.
        var suffix = sharedStateFingerprints ? "shared" : explorationId;
        var fp1 = Sha("state-root-" + suffix);
        var fp2 = Sha("state-child-" + suffix);
        var fp3 = Sha("state-alt-" + suffix);

        var stateRecords = new List<object>
        {
            new
            {
                stateId = "s0001",
                fingerprint = fp1,
                windowTitle = "DA Explorer Fixture",
                windowClassName = "WindowsForms10.Window.8.app.0.abc",
                isModal = false,
                depth = 0,
                capturedAtUtc = "2026-01-01T00:00:00Z",
                nodeCount = 10
            }
        };
        if (states >= 2)
        {
            stateRecords.Add(new
            {
                stateId = "s0002",
                fingerprint = fp2,
                windowTitle = "DA Explorer Second",
                windowClassName = "WindowsForms10.Window.8.app.0.abc",
                isModal = false,
                depth = 1,
                capturedAtUtc = "2026-01-01T00:00:01Z",
                nodeCount = 5
            });
        }

        if (states >= 3)
        {
            stateRecords.Add(new
            {
                stateId = "s0003",
                fingerprint = fp3,
                windowTitle = "DA Explorer Fixture",
                windowClassName = "WindowsForms10.Window.8.app.0.abc",
                isModal = false,
                depth = 1,
                capturedAtUtc = "2026-01-01T00:00:02Z",
                nodeCount = 10
            });
        }

        var transitions = new List<object>
        {
            new
            {
                transitionId = "t0001",
                fromStateId = "s0001",
                toStateId = states >= 2 ? "s0002" : "s0001",
                fromFingerprint = fp1,
                toFingerprint = states >= 2 ? fp2 : fp1,
                actionKey = "Invoke:Button|ExplorerOpenSecondButton|Open Second|o3|d1",
                actionKind = "Invoke",
                category = "invoke",
                driverOperation = "clickuia",
                elementKey = "Button|ExplorerOpenSecondButton|Open Second|o3|d1",
                automationId = "ExplorerOpenSecondButton",
                name = "Open Second",
                outcome = states >= 2 ? "newWindow" : "noOp",
                startedAtUtc = "2026-01-01T00:00:00Z",
                completedAtUtc = "2026-01-01T00:00:01Z",
                depth = 0
            }
        };

        if (includeExtraElement)
        {
            transitions.Add(new
            {
                transitionId = "t0002",
                fromStateId = "s0001",
                toStateId = "s0001",
                fromFingerprint = fp1,
                toFingerprint = fp1,
                actionKey = "Focus:Button|ExtraButton|Extra|o9|d1",
                actionKind = "Focus",
                category = "focus",
                driverOperation = "focusuia",
                elementKey = "Button|ExtraButton|Extra|o9|d1",
                automationId = "ExtraButton",
                name = "Extra",
                outcome = "noOp",
                startedAtUtc = "2026-01-01T00:00:01Z",
                completedAtUtc = "2026-01-01T00:00:01Z",
                depth = 0
            });
        }

        if (withConflict && states >= 3)
        {
            // Same source+action as t0001, different target/outcome 뿯↽ conflict.
            transitions.Add(new
            {
                transitionId = "t0003",
                fromStateId = "s0001",
                toStateId = "s0003",
                fromFingerprint = fp1,
                toFingerprint = fp3,
                actionKey = "Invoke:Button|ExplorerOpenSecondButton|Open Second|o3|d1",
                actionKind = "Invoke",
                category = "invoke",
                driverOperation = "clickuia",
                elementKey = "Button|ExplorerOpenSecondButton|Open Second|o3|d1",
                automationId = "ExplorerOpenSecondButton",
                name = "Open Second",
                outcome = "success",
                startedAtUtc = "2026-01-01T00:00:02Z",
                completedAtUtc = "2026-01-01T00:00:03Z",
                depth = 0
            });
        }

        WriteJson(Path.Combine(runDir, "manifest.json"), new
        {
            manifestSchemaVersion = 1,
            explorationId,
            mode = "navigate",
            dryRun = false,
            startedAtUtc = "2026-01-01T00:00:00Z",
            policySha256
        });
        WriteJson(Path.Combine(runDir, "summary.json"), new
        {
            summarySchemaVersion = 1,
            explorationId,
            status = "completed",
            exitCode = 0,
            mode = "navigate",
            dryRun = false,
            startedAtUtc = "2026-01-01T00:00:00Z",
            completedAtUtc = "2026-01-01T00:00:05Z",
            statesVisited = stateRecords.Count,
            actionsAttempted = transitions.Count,
            transitionsRecorded = transitions.Count,
            maxDepthReached = 1,
            stopReason = "completed",
            deniedActionSamples = new[] { "invoke:Button|Delete:forbidden-vocabulary" },
            warnings = Array.Empty<string>()
        });
        WriteJsonl(Path.Combine(runDir, "states.jsonl"), stateRecords);
        WriteJsonl(Path.Combine(runDir, "transitions.jsonl"), transitions);
        return explorationId;
    }

    private static KnowledgeService CreateKnowledgeService(IWorkspaceManager workspace) =>
        new(
            workspace,
            new ObjectArtifactWriter(),
            Options.Create(new AgentOptions
            {
                Workspace = new WorkspaceOptions { Root = workspace.RootPath },
                Knowledge = new KnowledgeOptions()
            }));

    private static IWorkspaceManager CreateWorkspace(string root) =>
        new WorkspaceManager(
            Options.Create(new AgentOptions { Workspace = new WorkspaceOptions { Root = root } }),
            NullLogger<WorkspaceManager>.Instance);

    private static TempDir TempWorkspace() => new();

    private static void WriteJson(string path, object value) =>
        File.WriteAllText(path, JsonSerializer.Serialize(value, KnowledgeJsonOptions(writeIndented: true)) + "\n");

    private static void WriteJsonl(string path, IEnumerable<object> values)
    {
        var sb = new StringBuilder();
        var options = KnowledgeJsonOptions(writeIndented: false);
        foreach (var value in values)
        {
            sb.Append(JsonSerializer.Serialize(value, options));
            sb.Append('\n');
        }

        File.WriteAllText(path, sb.ToString());
    }

    private static JsonSerializerOptions KnowledgeJsonOptions(bool writeIndented = true) =>
        new()
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = writeIndented,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

    private static EvaluationOptions EvalOptions() => new() { OutputFormat = OutputFormat.List };

    private static string Sha(string text)
    {
        var bytes = System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static string ReadRepoFile(string relativePath)
    {
        var root = FindRepoRoot();
        return File.ReadAllText(Path.Combine(root, relativePath));
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomation.sln"))
                || File.Exists(Path.Combine(dir.FullName, "desktopAutomationHelper.sln")))
            {
                return dir.FullName;
            }

            if (Directory.Exists(Path.Combine(dir.FullName, "automation"))
                && Directory.Exists(Path.Combine(dir.FullName, "src")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException("Unable to locate repository root.");
    }

    private static string NormalizeNewlines(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd() + "\n";

    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "da-phase4c-" + Guid.NewGuid().ToString("N"));

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
