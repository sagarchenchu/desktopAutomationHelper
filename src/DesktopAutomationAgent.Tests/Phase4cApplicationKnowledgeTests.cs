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
        var policyHash = new string('a', 64);
        var runDir = Path.Combine(temp.Path, "exploration", "runs", "huge-run");
        Directory.CreateDirectory(runDir);
        File.WriteAllText(
            Path.Combine(runDir, "manifest.json"),
            "{\"manifestSchemaVersion\":1,\"explorationId\":\"huge-run\",\"mode\":\"navigate\",\"dryRun\":false,\"startedAtUtc\":\"2026-01-01T00:00:00Z\",\"policySha256\":\"" + policyHash + "\"}");
        File.WriteAllText(Path.Combine(runDir, "summary.json"), """{"summarySchemaVersion":1,"explorationId":"huge-run","status":"ok","exitCode":0,"mode":"navigate","dryRun":false,"startedAtUtc":"2026-01-01T00:00:00Z","completedAtUtc":"2026-01-01T00:00:01Z","statesVisited":0,"actionsAttempted":0,"transitionsRecorded":0,"maxDepthReached":0}""");
        File.WriteAllText(Path.Combine(runDir, "states.jsonl"), "");
        File.WriteAllText(Path.Combine(runDir, "transitions.jsonl"), new string('x', 2000));
        File.WriteAllText(
            Path.Combine(runDir, "build-identity.json"),
            "{\"buildIdentitySchemaVersion\":1,\"processName\":\"Demo\",\"executableSha256\":\"" + new string('1', 64) + "\",\"driverVersion\":\"1.0.0\",\"snapshotSchemaVersion\":1,\"explorerVersion\":\"phase4b-1\",\"policySha256\":\"" + policyHash + "\",\"configIdentity\":{\"processName\":\"Demo\"}}");

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

    [Fact]
    public void Build_FailsClosed_WhenBuildIdentityMissing()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "no-identity", new string('7', 64), 1, false);
        File.Delete(Path.Combine(temp.Path, "exploration", "runs", run, "build-identity.json"));

        var result = CreateKnowledgeService(workspace).Build("demo-app", [run], false);
        Assert.False(result.Success);
        Assert.Contains("build-identity.json", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_RejectsDuplicateExplorationRuns()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "dup-id", new string('8', 64), 1, false);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [run, run], false);
        Assert.False(result.Success);
        Assert.Contains("Duplicate exploration", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_RejectsOutputOutsideCandidates()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "out-path", new string('9', 64), 1, false);
        var outside = Path.Combine(temp.Path, "knowledge", "approved", "demo-app", "fp");

        var result = CreateKnowledgeService(workspace).Build(
            "demo-app",
            [run],
            false,
            outputDirectory: outside,
            createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.False(result.Success);
        Assert.Contains("candidates", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Approve_RejectsSourceOutsideCandidates()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('a', 64);
        var run = WriteExplorationRun(temp.Path, "appr-src", policyHash, 1, false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);

        // Copy candidate into approved path and attempt to approve from there.
        var rogue = Path.Combine(temp.Path, "knowledge", "approved", "demo-app", "rogue");
        Directory.CreateDirectory(rogue);
        foreach (var file in Directory.GetFiles(built.OutputDirectory!))
            File.Copy(file, Path.Combine(rogue, Path.GetFileName(file)));

        var reviewed = Path.Combine(temp.Path, "review.json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = built.ContentHash!,
            SourceKnowledgeId = built.Knowledge!.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false
        }, KnowledgeJsonOptions()));

        var result = service.Approve(rogue, "alice", reviewed);
        Assert.False(result.Success);
        Assert.Contains("candidates", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateApproved_DetectsTampering()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "tamper-run", new string('b', 64), 2, false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);

        var reviewed = Path.Combine(temp.Path, "tamper-review.json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = built.ContentHash!,
            SourceKnowledgeId = built.Knowledge!.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false
        }, KnowledgeJsonOptions()));

        var approved = service.Approve(built.OutputDirectory!, "alice", reviewed);
        Assert.True(approved.Success, approved.Error);

        var knowledgePath = Path.Combine(approved.OutputDirectory!, "application-knowledge.json");
        var doc = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(
            File.ReadAllText(knowledgePath),
            KnowledgeJsonOptions())!;
        doc.Graph.States.Add(new KnowledgeStateNode
        {
            StateId = "st_tampered",
            SemanticFingerprint = new string('f', 64),
            WindowId = doc.Graph.Windows[0].WindowId,
            ObservationCount = 1
        });
        File.WriteAllText(knowledgePath, JsonSerializer.Serialize(doc, KnowledgeJsonOptions()) + "\n");

        var validation = service.ValidatePath(approved.OutputDirectory!);
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Errors, e => e.Contains("sourceKnowledgeHash mismatch", StringComparison.Ordinal));
    }

    [Fact]
    public void Workflows_SplitOnDisconnectedTransitions()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var policyHash = new string('c', 64);
        var runDir = Path.Combine(temp.Path, "exploration", "runs", "disconnected");
        Directory.CreateDirectory(runDir);

        var fpA = Sha("disc-a");
        var fpB = Sha("disc-b");
        var fpC = Sha("disc-c");
        var fpD = Sha("disc-d");

        WriteJsonl(Path.Combine(runDir, "states.jsonl"),
        [
            new { stateId = "s1", fingerprint = fpA, windowTitle = "A", windowClassName = "W", isModal = false, depth = 0, capturedAtUtc = "2026-01-01T00:00:00Z", nodeCount = 1 },
            new { stateId = "s2", fingerprint = fpB, windowTitle = "B", windowClassName = "W", isModal = false, depth = 1, capturedAtUtc = "2026-01-01T00:00:01Z", nodeCount = 1 },
            new { stateId = "s3", fingerprint = fpC, windowTitle = "C", windowClassName = "W", isModal = false, depth = 1, capturedAtUtc = "2026-01-01T00:00:02Z", nodeCount = 1 },
            new { stateId = "s4", fingerprint = fpD, windowTitle = "D", windowClassName = "W", isModal = false, depth = 1, capturedAtUtc = "2026-01-01T00:00:03Z", nodeCount = 1 }
        ]);
        // Continuous A->B, then disconnected C->D (no B->C link).
        WriteJsonl(Path.Combine(runDir, "transitions.jsonl"),
        [
            new { transitionId = "t1", fromStateId = "s1", toStateId = "s2", fromFingerprint = fpA, toFingerprint = fpB, actionKey = "Invoke:Button|One|One|o0|d0", actionKind = "Invoke", category = "invoke", driverOperation = "clickuia", elementKey = "Button|One|One|o0|d0", automationId = "One", name = "One", outcome = "success", startedAtUtc = "2026-01-01T00:00:00Z", completedAtUtc = "2026-01-01T00:00:01Z", depth = 0 },
            new { transitionId = "t2", fromStateId = "s3", toStateId = "s4", fromFingerprint = fpC, toFingerprint = fpD, actionKey = "Invoke:Button|Two|Two|o0|d0", actionKind = "Invoke", category = "invoke", driverOperation = "clickuia", elementKey = "Button|Two|Two|o0|d0", automationId = "Two", name = "Two", outcome = "success", startedAtUtc = "2026-01-01T00:00:02Z", completedAtUtc = "2026-01-01T00:00:03Z", depth = 0 }
        ]);
        WriteJson(Path.Combine(runDir, "manifest.json"), new { manifestSchemaVersion = 1, explorationId = "disconnected", mode = "navigate", dryRun = false, startedAtUtc = "2026-01-01T00:00:00Z", policySha256 = policyHash });
        WriteJson(Path.Combine(runDir, "summary.json"), new { summarySchemaVersion = 1, explorationId = "disconnected", status = "completed", exitCode = 0, mode = "navigate", dryRun = false, startedAtUtc = "2026-01-01T00:00:00Z", completedAtUtc = "2026-01-01T00:00:03Z", statesVisited = 4, actionsAttempted = 2, transitionsRecorded = 2, maxDepthReached = 1, stopReason = "completed", deniedActionSamples = Array.Empty<string>(), warnings = Array.Empty<string>() });
        WriteJson(Path.Combine(runDir, "build-identity.json"), new
        {
            buildIdentitySchemaVersion = 1,
            processName = "Demo",
            executableSha256 = new string('2', 64),
            driverVersion = "1.0.0-test",
            snapshotSchemaVersion = 1,
            explorerVersion = "phase4b-1",
            policySha256 = policyHash,
            configIdentity = new Dictionary<string, string> { ["processName"] = "Demo" }
        });

        var result = CreateKnowledgeService(workspace).Build("demo-app", ["disconnected"], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.True(result.Knowledge!.Workflows.Count >= 2, "Disconnected transitions must split into separate workflows.");
        Assert.All(result.Knowledge.Workflows, wf =>
        {
            for (var i = 1; i < wf.Steps.Count; i++)
            {
                var prevTarget = wf.Steps[i - 1].TargetStateId ?? wf.Steps[i - 1].SourceStateId;
                Assert.Equal(prevTarget, wf.Steps[i].SourceStateId);
            }
        });
    }

    [Fact]
    public void ApprovedKnowledge_StripsAbsolutePaths()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "abs-path", new string('d', 64), 1, false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);
        Assert.All(built.Knowledge!.SourceExplorationRuns, s =>
            Assert.False(Path.IsPathRooted(s.ManifestPath ?? string.Empty)));

        var reviewed = Path.Combine(temp.Path, "abs-review.json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = built.ContentHash!,
            SourceKnowledgeId = built.Knowledge.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false
        }, KnowledgeJsonOptions()));

        var approved = service.Approve(built.OutputDirectory!, "alice", reviewed);
        Assert.True(approved.Success, approved.Error);
        Assert.Equal("abs-review.json", approved.Knowledge!.Approval!.ReviewedInputPath);
        Assert.All(approved.Knowledge.SourceExplorationRuns, s =>
            Assert.StartsWith("exploration/runs/", s.ManifestPath, StringComparison.Ordinal));
    }

    [Fact]
    public void Build_EmitsContentHashAndApprovalTemplate()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRun(temp.Path, "hash-emit", new string('e', 64), 1, false);
        var result = CreateKnowledgeService(workspace).Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(result.Success, result.Error);
        Assert.False(string.IsNullOrWhiteSpace(result.ContentHash));
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, "content-hash.txt")));
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, "approval-template.json")));
        Assert.Equal(result.ContentHash, File.ReadAllText(Path.Combine(result.OutputDirectory!, "content-hash.txt")).Trim());
    }

    [Fact]
    public void Fingerprint_SamePolicyDifferentExecutableHash_IsIncompatible()
    {
        var builder = new ApplicationFingerprintBuilder();
        var policyHash = new string('a', 64);
        var leftIdentity = new ExplorationBuildIdentity
        {
            BuildIdentitySchemaVersion = 1,
            ProcessName = "Demo",
            ExecutableSha256 = new string('1', 64),
            DriverVersion = "1.0.0-test",
            SnapshotSchemaVersion = 1,
            ExplorerVersion = KnowledgeConstants.ExplorerVersion,
            PolicySha256 = policyHash,
            ConfigIdentity = new Dictionary<string, string> { ["processName"] = "Demo" }
        };
        var rightIdentity = new ExplorationBuildIdentity
        {
            BuildIdentitySchemaVersion = 1,
            ProcessName = "Demo",
            ExecutableSha256 = new string('2', 64),
            DriverVersion = "1.0.0-test",
            SnapshotSchemaVersion = 1,
            ExplorerVersion = KnowledgeConstants.ExplorerVersion,
            PolicySha256 = policyHash,
            ConfigIdentity = new Dictionary<string, string> { ["processName"] = "Demo" }
        };

        var left = builder.BuildFromIdentity("demo-app", leftIdentity);
        var right = builder.BuildFromIdentity("demo-app", rightIdentity);

        Assert.False(builder.AreCompatible(left, right, out var mismatches));
        Assert.Contains("ExecutableSha256", mismatches);
        Assert.NotEqual(builder.ComputeBuildFingerprint(left), builder.ComputeBuildFingerprint(right));
    }

    [Fact]
    public void Build_FailsClosed_WhenImmutableBuildIdMissing()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var run = WriteExplorationRunWithIdentity(
            temp.Path,
            "no-immutable",
            new string('a', 64),
            executableSha256: null,
            productVersion: null,
            fileVersion: null);

        var result = CreateKnowledgeService(workspace).Build("demo-app", [run], false);
        Assert.False(result.Success);
        Assert.Contains("immutable build id", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SystemProcessIdentityProbe_CapturesShaAndVersions()
    {
        var probe = new SystemProcessIdentityProbe();
        Assert.True(
            probe.TryGetIdentity(Environment.ProcessId, out var info, out var error),
            error ?? "probe failed");
        Assert.False(string.IsNullOrWhiteSpace(info.ExecutablePath));
        Assert.False(string.IsNullOrWhiteSpace(info.ExecutableSha256));
        Assert.Equal(64, info.ExecutableSha256!.Length);
        Assert.True(
            !string.IsNullOrWhiteSpace(info.ProductVersion)
            || !string.IsNullOrWhiteSpace(info.FileVersion)
            || !string.IsNullOrWhiteSpace(info.ExecutableSha256),
            "Probe must capture at least one immutable build id.");
    }

    [Fact]
    public void ValidateApproved_RejectsMissingContentHash()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var approvedDir = ApproveMinimal(temp.Path, workspace, "missing-hash");
        File.Delete(Path.Combine(approvedDir, "content-hash.txt"));

        var validation = CreateKnowledgeService(workspace).ValidatePath(approvedDir);
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Errors, e => e.Contains("content-hash.txt", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ValidateApproved_RejectsSourceKnowledgeIdMismatch()
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var approvedDir = ApproveMinimal(temp.Path, workspace, "id-mismatch");
        var knowledgePath = Path.Combine(approvedDir, "application-knowledge.json");
        var doc = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(
            File.ReadAllText(knowledgePath),
            KnowledgeJsonOptions())!;
        doc.Approval!.SourceKnowledgeId = "kn_wrong_id";
        File.WriteAllText(knowledgePath, JsonSerializer.Serialize(doc, KnowledgeJsonOptions()) + "\n");
        File.WriteAllText(
            Path.Combine(approvedDir, "knowledge-approval.json"),
            JsonSerializer.Serialize(doc.Approval, KnowledgeJsonOptions()) + "\n");

        var validation = CreateKnowledgeService(workspace).ValidatePath(approvedDir);
        Assert.False(validation.IsValid);
        Assert.Contains(
            validation.Errors,
            e => e.Contains("sourceKnowledgeId", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("state-graph.json")]
    [InlineData("workflows.json")]
    [InlineData("coverage.json")]
    [InlineData("knowledge-approval.json")]
    public void ValidateApproved_DetectsSiblingTampering(string siblingFileName)
    {
        using var temp = TempWorkspace();
        var workspace = CreateWorkspace(temp.Path);
        workspace.Initialize();
        var approvedDir = ApproveMinimal(temp.Path, workspace, "sibling-" + Path.GetFileNameWithoutExtension(siblingFileName));
        var siblingPath = Path.Combine(approvedDir, siblingFileName);

        if (siblingFileName == "state-graph.json")
        {
            var graph = JsonSerializer.Deserialize<ApplicationStateGraph>(
                File.ReadAllText(siblingPath),
                KnowledgeJsonOptions())!;
            graph.States.Add(new KnowledgeStateNode
            {
                StateId = "st_sibling_tamper",
                SemanticFingerprint = new string('e', 64),
                WindowId = graph.Windows[0].WindowId,
                ObservationCount = 1
            });
            File.WriteAllText(siblingPath, JsonSerializer.Serialize(graph, KnowledgeJsonOptions()) + "\n");
        }
        else if (siblingFileName == "workflows.json")
        {
            var workflows = JsonSerializer.Deserialize<List<ApplicationWorkflow>>(
                File.ReadAllText(siblingPath),
                KnowledgeJsonOptions()) ?? [];
            workflows.Add(new ApplicationWorkflow
            {
                WorkflowSchemaVersion = 1,
                WorkflowId = "wf_tampered",
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
            });
            File.WriteAllText(siblingPath, JsonSerializer.Serialize(workflows, KnowledgeJsonOptions()) + "\n");
        }
        else if (siblingFileName == "coverage.json")
        {
            var coverage = JsonSerializer.Deserialize<KnowledgeCoverageReport>(
                File.ReadAllText(siblingPath),
                KnowledgeJsonOptions())!;
            coverage.Conflicts = coverage.Conflicts + 99;
            File.WriteAllText(siblingPath, JsonSerializer.Serialize(coverage, KnowledgeJsonOptions()) + "\n");
        }
        else
        {
            var approval = JsonSerializer.Deserialize<KnowledgeApproval>(
                File.ReadAllText(siblingPath),
                KnowledgeJsonOptions())!;
            approval.Notes = "tampered sibling notes";
            File.WriteAllText(siblingPath, JsonSerializer.Serialize(approval, KnowledgeJsonOptions()) + "\n");
        }

        var validation = CreateKnowledgeService(workspace).ValidatePath(approvedDir);
        Assert.False(validation.IsValid);
        Assert.Contains(
            validation.Errors,
            e => e.Contains(siblingFileName, StringComparison.OrdinalIgnoreCase)
                 && e.Contains("match", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ContentHash_BindsKnowledgeId()
    {
        var doc = MinimalKnowledge();
        var original = KnowledgeArtifactWriter.ComputeContentHash(doc);
        doc.KnowledgeId = "kn_changed";
        var changed = KnowledgeArtifactWriter.ComputeContentHash(doc);
        Assert.NotEqual(original, changed);
    }

    private static string ApproveMinimal(string root, IWorkspaceManager workspace, string runSuffix)
    {
        var run = WriteExplorationRun(root, "approve-" + runSuffix, new string('b', 64), 2, false);
        var service = CreateKnowledgeService(workspace);
        var built = service.Build("demo-app", [run], false, createdAtUtc: DateTimeOffset.UnixEpoch);
        Assert.True(built.Success, built.Error);

        var reviewed = Path.Combine(root, "review-" + runSuffix + ".json");
        File.WriteAllText(reviewed, JsonSerializer.Serialize(new KnowledgeApproval
        {
            ApprovalSchemaVersion = 1,
            Reviewer = "alice",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = built.ContentHash!,
            SourceKnowledgeId = built.Knowledge!.KnowledgeId,
            Decision = "approve",
            PromotesPageObjects = false
        }, KnowledgeJsonOptions()));

        var approved = service.Approve(built.OutputDirectory!, "alice", reviewed);
        Assert.True(approved.Success, approved.Error);
        return approved.OutputDirectory!;
    }

    private static ApplicationKnowledgeDocument MinimalKnowledge()
    {
        var fp = new ApplicationFingerprint
        {
            ApplicationId = "demo-app",
            ExecutableSha256 = new string('a', 64),
            SnapshotSchemaVersion = 1,
            ExplorerVersion = KnowledgeConstants.ExplorerVersion,
            ExplorationPolicyHash = new string('a', 64),
            DriverVersion = "1.0.0-test",
            ConfigIdentity = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["processName"] = "DemoApp"
            }
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
                    PolicySha256 = new string('a', 64),
                    ManifestPath = "exploration/runs/run-1/manifest.json"
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
        WriteJson(Path.Combine(runDir, "build-identity.json"), new
        {
            buildIdentitySchemaVersion = 1,
            processName = "DesktopAutomationExplorerFixture",
            executablePath = (string?)null,
            executableSha256 = new string('f', 64),
            productVersion = (string?)null,
            fileVersion = (string?)null,
            driverVersion = "1.0.0-test",
            snapshotSchemaVersion = 1,
            explorerVersion = "phase4b-1",
            policySha256,
            configIdentity = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["processName"] = "DesktopAutomationExplorerFixture"
            }
        });
        WriteJsonl(Path.Combine(runDir, "states.jsonl"), stateRecords);
        WriteJsonl(Path.Combine(runDir, "transitions.jsonl"), transitions);
        return explorationId;
    }

    private static string WriteExplorationRunWithIdentity(
        string root,
        string explorationId,
        string policySha256,
        string? executableSha256,
        string? productVersion = null,
        string? fileVersion = null,
        int states = 1)
    {
        var id = WriteExplorationRun(root, explorationId, policySha256, states, withConflict: false);
        WriteJson(Path.Combine(root, "exploration", "runs", id, "build-identity.json"), new
        {
            buildIdentitySchemaVersion = 1,
            processName = "DesktopAutomationExplorerFixture",
            executablePath = (string?)null,
            executableSha256,
            productVersion,
            fileVersion,
            driverVersion = "1.0.0-test",
            snapshotSchemaVersion = 1,
            explorerVersion = "phase4b-1",
            policySha256,
            configIdentity = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["processName"] = "DesktopAutomationExplorerFixture"
            }
        });
        return id;
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
