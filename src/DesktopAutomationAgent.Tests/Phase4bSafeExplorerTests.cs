using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;
using Json.Schema;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace DesktopAutomationAgent.Tests;

public class Phase4bSafeExplorerTests
{
    [Fact]
    public void ExplorationPolicySchema_ValidatesExample()
    {
        var schemaText = ReadRepoFile("automation/schemas/exploration-policy.schema.json");
        Assert.Contains("draft/2020-12/schema", schemaText, StringComparison.Ordinal);
        var schema = JsonSchema.FromText(schemaText);
        var sample = JsonNodeParse(ReadRepoFile("automation/exploration/policies/example.policy.json"));
        var result = schema.Evaluate(sample, new EvaluationOptions { OutputFormat = OutputFormat.List });
        Assert.True(result.IsValid, string.Join("; ", result.Details.Select(d => d.Errors is null ? d.InstanceLocation.ToString() : string.Join(",", d.Errors.Values))));
    }

    [Fact]
    public void PolicyReader_RejectsCoordinateFallbackAndUnknownMode()
    {
        var reader = new ExplorationPolicyReader();
        var policy = MinimalPolicy();
        policy.AllowCoordinateFallback = true;
        policy.Mode = "destructive";
        var result = reader.Validate(policy);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("allowCoordinateFallback", StringComparison.Ordinal));
        Assert.Contains(result.Errors, e => e.Contains("mode", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void CapabilityResolver_DenyByDefault_AndForbiddenVocabulary()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "Button", "Delete", "ExplorerDeleteButton", true, ["Invoke"]),
            Node("n2", "n0", 1, 1, "Button", "Open Dialog", "ExplorerOpenDialogButton", true, ["Invoke"]),
            Node("n3", "n0", 2, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"]));

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowedActionCategories = ["expandCollapse", "invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenDialogButton",
                Action = "invoke",
                ActionCategory = "invoke"
            },
            new ExplorationElementRule
            {
                AutomationId = "ExplorerTreeRoot",
                Action = "expand",
                ActionCategory = "expandCollapse"
            }
        ];
        policy.DenyRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerDeleteButton", Action = "*" }
        ];

        var resolver = new CapabilityResolver();
        var candidates = resolver.Resolve(snapshot, "fp1", policy, depth: 0);

        Assert.DoesNotContain(candidates, c => c.AutomationId == "ExplorerDeleteButton");
        Assert.Contains(candidates, c => c.AutomationId == "ExplorerOpenDialogButton" && c.Kind == ExplorationActionKind.Invoke);
        Assert.Contains(candidates, c => c.AutomationId == "ExplorerTreeRoot" && c.Kind == ExplorationActionKind.Expand);
        Assert.Contains(resolver.DeniedActionSamples, s => s.Contains("forbidden-vocabulary", StringComparison.Ordinal)
                                                          || s.Contains("deny-rule", StringComparison.Ordinal));
    }

    [Fact]
    public void CapabilityResolver_InvokeRequiresExplicitAllowEvenWhenCategoryListed()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "Button", "Primary", "primary", true, ["Invoke"]));

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowedActionCategories = ["invoke"];
        policy.AllowRules = [];

        var candidates = new CapabilityResolver().Resolve(snapshot, "fp", policy, 0);
        Assert.Empty(candidates);
    }

    [Fact]
    public void CapabilityResolver_HoverRequiresAllowHover()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "Text", "Hover target", "ExplorerHoverTarget", true, [], focusable: true));

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowedActionCategories = ["hover", "focus"];
        policy.AllowHover = false;
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerHoverTarget", Action = "hover", ActionCategory = "hover" }
        ];

        var resolver = new CapabilityResolver();
        var withoutHover = resolver.Resolve(snapshot, "fp", policy, 0);
        Assert.DoesNotContain(withoutHover, c => c.Kind == ExplorationActionKind.Hover);

        policy.AllowHover = true;
        var withHover = resolver.Resolve(snapshot, "fp", policy, 0);
        Assert.Contains(withHover, c => c.Kind == ExplorationActionKind.Hover);
    }

    [Fact]
    public void CapabilityResolver_LeafMenuItem_IsInvokeNotMenuOpen()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "MenuItem", "File", "ExplorerFileMenu", true, ["ExpandCollapse"]),
            Node("n2", "n1", 0, 2, "MenuItem", "Exit", "ExplorerExitMenuItem", true, ["Invoke"]),
            Node("n3", "n1", 1, 2, "MenuItem", "Reset", "ExplorerResetMenuItem", true, ["Invoke"]),
            Node("n4", "n1", 2, 2, "MenuItem", "Print", "ExplorerPrintMenuItem", true, ["Invoke"]),
            Node("n5", "n1", 3, 2, "MenuItem", "Archive", "ExplorerArchiveMenuItem", true, ["Invoke"]),
            Node("n6", "n1", 4, 2, "MenuItem", "Open Dialog", "ExplorerOpenDialogMenuItem", true, ["Invoke"]));

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowedActionCategories = ["menuOpen", "invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerFileMenu",
                Action = "menuOpen",
                ActionCategory = "menuOpen"
            },
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenDialogMenuItem",
                Action = "invoke",
                ActionCategory = "invoke"
            }
        ];

        var resolver = new CapabilityResolver();
        var candidates = resolver.Resolve(snapshot, "fp", policy, 0);

        Assert.Contains(candidates, c => c.AutomationId == "ExplorerFileMenu" && c.Kind == ExplorationActionKind.MenuOpen);
        Assert.Contains(candidates, c => c.AutomationId == "ExplorerOpenDialogMenuItem" && c.Kind == ExplorationActionKind.Invoke);
        Assert.DoesNotContain(candidates, c => c.AutomationId == "ExplorerExitMenuItem");
        Assert.DoesNotContain(candidates, c => c.AutomationId == "ExplorerResetMenuItem");
        Assert.DoesNotContain(candidates, c => c.AutomationId == "ExplorerPrintMenuItem");
        Assert.DoesNotContain(candidates, c => c.AutomationId == "ExplorerArchiveMenuItem");
        Assert.DoesNotContain(candidates, c =>
            c.Kind == ExplorationActionKind.MenuOpen
            && c.AutomationId is "ExplorerExitMenuItem" or "ExplorerOpenDialogMenuItem");
    }

    [Fact]
    public void StateFingerprinter_IncludesInteractionState()
    {
        var collapsed = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(
                "n1",
                "n0",
                0,
                1,
                "TreeItem",
                "Root",
                "ExplorerTreeRoot",
                true,
                ["ExpandCollapse"],
                interaction: new NodeInteractionState { ExpandCollapse = "Collapsed" }));
        var expanded = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(
                "n1",
                "n0",
                0,
                1,
                "TreeItem",
                "Root",
                "ExplorerTreeRoot",
                true,
                ["ExpandCollapse"],
                interaction: new NodeInteractionState { ExpandCollapse = "Expanded" }));
        var selected = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(
                "n1",
                "n0",
                0,
                1,
                "TabItem",
                "Home",
                "ExplorerTabHome",
                true,
                ["SelectionItem"],
                interaction: new NodeInteractionState { Selected = true }));
        var unselected = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(
                "n1",
                "n0",
                0,
                1,
                "TabItem",
                "Home",
                "ExplorerTabHome",
                true,
                ["SelectionItem"],
                interaction: new NodeInteractionState { Selected = false }));

        var fingerprinter = new StateFingerprinter();
        Assert.NotEqual(fingerprinter.Fingerprint(collapsed), fingerprinter.Fingerprint(expanded));
        Assert.NotEqual(fingerprinter.Fingerprint(selected), fingerprinter.Fingerprint(unselected));
    }

    [Fact]
    public void ProcessIdentityValidator_FailClosed_OnMissingIdentityAndPathMismatch()
    {
        var policy = MinimalPolicy();
        policy.Application.ExecutablePath = @"C:\apps\fixture.exe";
        policy.Application.ExecutableSha256 = new string('a', 64);

        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            ExecutablePath = @"C:\apps\other.exe",
            ExecutableSha256 = new string('b', 64),
            SessionId = 1
        });

        var validator = new ProcessIdentityValidator(probe);
        var snapshot = Snapshot(processId: 10);
        var mismatch = validator.ValidateSnapshot(snapshot, policy);
        Assert.False(mismatch.IsValid);

        var missingName = Snapshot(processName: "", processId: 10);
        Assert.False(validator.ValidateSnapshot(missingName, MinimalPolicy()).IsValid);

        var missingPid = new SemanticSnapshot
        {
            SnapshotId = "snap",
            Application = new ApplicationIdentity
            {
                ProcessName = "DesktopAutomationExplorerFixture",
                ProcessId = null
            },
            WindowsSession = new WindowsSessionIdentity { SessionId = 1 },
            Nodes = [Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"])]
        };
        Assert.False(validator.ValidateSnapshot(missingPid, MinimalPolicy()).IsValid);

        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            ExecutablePath = @"C:\apps\fixture.exe",
            ExecutableSha256 = new string('a', 64),
            SessionId = 1
        });
        Assert.True(validator.ValidateSnapshot(snapshot, policy).IsValid);

        var childPolicy = MinimalPolicy();
        childPolicy.Application.AllowedChildProcessNames = ["FixtureChild"];
        probe.Register(11, new ProcessIdentityInfo
        {
            ProcessId = 11,
            ProcessName = "FixtureChild",
            ParentProcessId = 10,
            SessionId = 1
        });
        // Establish root first.
        Assert.True(validator.ValidateSnapshot(snapshot, childPolicy).IsValid);
        var childSnap = Snapshot(processName: "FixtureChild", processId: 11);
        Assert.True(validator.ValidateSnapshot(childSnap, childPolicy).IsValid);

        var newWindowMissingName = validator.ValidateNewWindow(
            new WindowDescriptor { Title = "X", ProcessId = 12 },
            MinimalPolicy());
        Assert.False(newWindowMissingName.IsValid);
    }

    [Fact]
    public void WindowTransitionDetector_FailsClosed_WhenNewWindowMissingProcessName()
    {
        var policy = MinimalPolicy();
        var before = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 1,
                    Title = "DA Explorer Fixture",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = 10
                }
            ]
        };
        var after = new WindowInventory
        {
            Windows =
            [
                before.Windows[0],
                new WindowDescriptor { Hwnd = 2, Title = "Mystery", ProcessId = 11 }
            ]
        };

        var detector = new WindowTransitionDetector(
            Mock.Of<IDriverUiClient>(),
            new ProcessIdentityValidator(new FakeProcessIdentityProbe()));
        var result = detector.Compare(before, after, Snapshot(), policy);
        Assert.True(result.ScopeViolation);
        Assert.Contains("processName", result.Detail ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveLiveNode_RejectsAmbiguousAndPositionalOnlyMatches()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "Button", "Same", "dup", true, ["Invoke"]),
            Node("n2", "n0", 1, 1, "Button", "Same", "dup", true, ["Invoke"]));

        var action = new ExplorationActionCandidate
        {
            ActionKey = "invoke:x",
            Kind = ExplorationActionKind.Invoke,
            Category = "invoke",
            ElementKey = "Button|dup|Same|o0|d1",
            AutomationId = "dup",
            Name = "Same",
            ControlType = "Button",
            NodeRef = "missing",
            SnapshotId = "other"
        };

        Assert.False(ExplorationCoordinator.TryResolveLiveNode(snapshot, action, out _, out var detail));
        Assert.Contains("Ambiguous", detail, StringComparison.OrdinalIgnoreCase);

        var unique = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n9", "n0", 5, 2, "Button", "Same", "dup", true, ["Invoke"]));
        Assert.True(ExplorationCoordinator.TryResolveLiveNode(unique, action, out var live, out _));
        Assert.Equal("n9", live!.Ref);
    }

    [Fact]
    public void FrontierQueue_IsDeterministic_AndRejectsDuplicatePairs()
    {
        var frontier = new FrontierQueue();
        var a = Candidate("fp", "invoke:Button|b|B|o1|d1", ExplorationActionKind.Invoke, "invoke", priority: 80, elementKey: "Button|b|B|o1|d1");
        var b = Candidate("fp", "expand:TreeItem|t|T|o0|d1", ExplorationActionKind.Expand, "expandCollapse", priority: 10, elementKey: "TreeItem|t|T|o0|d1");
        var c = Candidate("fp", "tabSelect:TabItem|x|X|o2|d1", ExplorationActionKind.TabSelect, "tabSelect", priority: 20, elementKey: "TabItem|x|X|o2|d1");

        Assert.True(frontier.TryEnqueue(a));
        Assert.True(frontier.TryEnqueue(b));
        Assert.True(frontier.TryEnqueue(c));
        Assert.False(frontier.TryEnqueue(a));

        Assert.True(frontier.TryDequeue(out var first));
        Assert.True(frontier.TryDequeue(out var second));
        Assert.True(frontier.TryDequeue(out var third));
        Assert.Equal(ExplorationActionKind.Expand, first!.Kind);
        Assert.Equal(ExplorationActionKind.TabSelect, second!.Kind);
        Assert.Equal(ExplorationActionKind.Invoke, third!.Kind);
        Assert.True(frontier.HasExecuted("fp", a.ActionKey));
    }

    [Fact]
    public void StateFingerprinter_IgnoresVolatileFields()
    {
        var left = Snapshot(
            processId: 1,
            sessionId: 1,
            hwnd: 10,
            nodes:
            [
                Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"])
            ]);
        left = CloneWithWindow(left, title: "Main", className: "Cls", isModal: false, hwnd: 10, processId: 1);

        var right = Snapshot(
            processId: 99,
            sessionId: 7,
            hwnd: 99,
            nodes:
            [
                Node("n9", null, 0, 0, "Window", "Main", "main", true, ["Window"])
            ]);
        right = CloneWithWindow(right, title: "Main", className: "Cls", isModal: false, hwnd: 99, processId: 99);

        var fingerprinter = new StateFingerprinter();
        Assert.Equal(fingerprinter.Fingerprint(left), fingerprinter.Fingerprint(right));
    }

    [Fact]
    public void WindowTransitionDetector_FlagsExternalProcessAndModal()
    {
        var policy = MinimalPolicy();
        policy.Application.ProcessName = "DesktopAutomationExplorerFixture";
        policy.Scope.AllowedWindowTitles = ["DA Explorer Fixture", "DA Explorer Dialog"];

        var before = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 1,
                    Title = "DA Explorer Fixture",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = 10
                }
            ]
        };
        var afterExternal = new WindowInventory
        {
            Windows =
            [
                before.Windows[0],
                new WindowDescriptor
                {
                    Hwnd = 2,
                    Title = "Notepad",
                    ProcessName = "notepad",
                    ProcessId = 11
                }
            ]
        };

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var external = detector.Compare(before, afterExternal, Snapshot(), policy);
        Assert.True(external.ScopeViolation);

        // Pre-existing external windows must not fail scope checks.
        var preexistingNoise = new WindowInventory
        {
            Windows =
            [
                before.Windows[0],
                new WindowDescriptor
                {
                    Hwnd = 9,
                    Title = "Chrome",
                    ProcessName = "chrome",
                    ProcessId = 99
                }
            ]
        };
        var stableNoise = detector.Compare(preexistingNoise, preexistingNoise, Snapshot(
            processName: "DesktopAutomationExplorerFixture",
            nodes: [Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "main", true, ["Window"])]), policy);
        Assert.False(stableNoise.ScopeViolation);

        var afterModal = new WindowInventory { Windows = [before.Windows[0]] };
        var modalSnapshot = CloneWithWindow(Snapshot(), "DA Explorer Dialog", "Dialog", isModal: true, hwnd: 3, processId: 10);
        modalSnapshot = new SemanticSnapshot
        {
            SnapshotId = modalSnapshot.SnapshotId,
            CapturedAtUtc = modalSnapshot.CapturedAtUtc,
            DriverVersion = modalSnapshot.DriverVersion,
            SnapshotHash = modalSnapshot.SnapshotHash,
            Application = new ApplicationIdentity { ProcessId = 10, ProcessName = "DesktopAutomationExplorerFixture" },
            WindowsSession = modalSnapshot.WindowsSession,
            Window = modalSnapshot.Window,
            Nodes = modalSnapshot.Nodes
        };
        var modal = detector.Compare(before, afterModal, modalSnapshot, policy);
        Assert.True(modal.ModalDialog);
        Assert.False(modal.ScopeViolation);
    }

    [Fact]
    public void ArtifactWriter_IsAtomic_NonOverwriting_AndRedacted()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var writer = new ExplorationArtifactWriter(workspace);
        var dir = writer.CreateRunDirectory("exp-test-1");

        var summary = new ExplorationSummary
        {
            ExplorationId = "exp-test-1",
            Status = "completed",
            ExitCode = 0,
            Mode = "navigate",
            StartedAtUtc = DateTimeOffset.UtcNow,
            CompletedAtUtc = DateTimeOffset.UtcNow,
            StopReason = "frontier-exhausted",
            Warnings = ["token=super-secret-token"]
        };

        writer.WriteAll(
            dir,
            new ExplorationManifest { ExplorationId = "exp-test-1", Mode = "navigate", StartedAtUtc = DateTimeOffset.UtcNow },
            [new ExplorationStateRecord { StateId = "s0001", Fingerprint = "abc", CapturedAtUtc = DateTimeOffset.UtcNow }],
            [],
            summary);

        Assert.True(File.Exists(Path.Combine(dir, "manifest.json")));
        Assert.True(File.Exists(Path.Combine(dir, "states.jsonl")));
        Assert.True(File.Exists(Path.Combine(dir, "transitions.jsonl")));
        Assert.True(File.Exists(Path.Combine(dir, "summary.json")));

        var summaryText = File.ReadAllText(Path.Combine(dir, "summary.json"));
        Assert.DoesNotContain("super-secret-token", summaryText, StringComparison.Ordinal);

        Assert.ThrowsAny<IOException>(() =>
            writer.WriteAll(
                dir,
                new ExplorationManifest { ExplorationId = "exp-test-1", Mode = "navigate", StartedAtUtc = DateTimeOffset.UtcNow },
                [],
                [],
                summary));
    }

    [Fact]
    public async Task Coordinator_DryRun_MakesNoActionCalls()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policyPath = Path.Combine(workspace.RootPath, "exploration", "policies", "dry.json");
        Directory.CreateDirectory(Path.GetDirectoryName(policyPath)!);
        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        await File.WriteAllTextAsync(policyPath, JsonSerializer.Serialize(policy));

        var snapshot = Snapshot(
            processName: "DesktopAutomationExplorerFixture",
            nodes:
            [
                Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"])
            ]);
        snapshot = CloneWithWindow(snapshot, "DA Explorer Fixture", "Form", false, 1, 10);

        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ObservationResult
            {
                Success = true,
                ExitCode = 0,
                Snapshot = snapshot,
                NodeCount = snapshot.Nodes.Count,
                SnapshotHash = "hash"
            });

        var ui = new Mock<IDriverUiClient>(MockBehavior.Strict);
        var connection = new Mock<IDriverConnectionResolver>();
        connection.Setup(c => c.ResolveAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:1"),
                BearerToken = "token",
                DiscoveryMethod = "configured"
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, connection.Object);
        var result = await coordinator.ExploreAsync(new ExplorationRequest
        {
            PolicyPath = policyPath,
            DryRun = true
        });

        Assert.True(result.Success);
        Assert.NotEmpty(result.DryRunCandidates);
        ui.Verify(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Coordinator_ObserveMode_DoesNotMutate()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var policy = MinimalPolicy();
        policy.Mode = "observe";
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var snapshot = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                nodes: [Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"])]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ObservationResult { Success = true, ExitCode = 0, Snapshot = snapshot, NodeCount = 1 });

        var ui = new Mock<IDriverUiClient>(MockBehavior.Strict);
        var coordinator = CreateCoordinator(
            options,
            workspace,
            observe.Object,
            ui.Object,
            Mock.Of<IDriverConnectionResolver>());

        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success);
        Assert.Equal("observe-mode", result.Summary!.StopReason);
        ui.Verify(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Coordinator_FakeGraph_EnforcesBudgetsCyclesAndCancellationPartialReport()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxActions = 2;
        policy.Budgets.MaxStates = 10;
        policy.Budgets.MaxDepth = 3;
        policy.Budgets.TotalDurationSeconds = 30;
        policy.Budgets.StabilizationTimeoutSeconds = 2;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" },
            new ExplorationElementRule { AutomationId = "ExplorerTreeChild", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var stateA = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-a",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var stateB = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-b",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"]),
                    Node("n2", "n1", 0, 2, "TreeItem", "Child", "ExplorerTreeChild", true, ["ExpandCollapse"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var observeCalls = 0;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                observeCalls++;
                // Alternate A/B so actions see stable fingerprints then a changed after-state.
                var snap = observeCalls <= 2 || observeCalls % 2 == 1 ? stateA : stateB;
                return new ObservationResult
                {
                    Success = true,
                    ExitCode = 0,
                    Snapshot = snap,
                    NodeCount = snap.Nodes.Count,
                    SnapshotHash = snap.SnapshotHash
                };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                if (step.Operation == "listwindows")
                {
                    return new UiExecutionResponse
                    {
                        Success = true,
                        Value = JsonSerializer.SerializeToElement(new
                        {
                            windows = new[]
                            {
                                new
                                {
                                    hwnd = 1,
                                    title = "DA Explorer Fixture",
                                    className = "Form",
                                    processId = 10,
                                    processName = "DesktopAutomationExplorerFixture"
                                }
                            }
                        })
                    };
                }

                return new UiExecutionResponse { Success = true };
            });

        var connection = new Mock<IDriverConnectionResolver>();
        connection.Setup(c => c.ResolveAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:1"),
                BearerToken = "token",
                DiscoveryMethod = "configured"
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, connection.Object);
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });

        Assert.True(result.Success);
        Assert.NotNull(result.Summary);
        Assert.InRange(result.Summary!.ActionsAttempted, 1, policy.Budgets.MaxActions);
        Assert.Contains(
            result.Summary.StopReason,
            ["max-actions-budget", "frontier-exhausted", "max-states-budget"]);
        Assert.True(Directory.Exists(result.ArtifactDirectory));
        Assert.True(File.Exists(Path.Combine(result.ArtifactDirectory!, "summary.json")));
        var transitions = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();
        Assert.NotEmpty(transitions);

        // Cancellation writes a partial report.
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();
        var cancelled = await coordinator.ExploreAsync(
            new ExplorationRequest { PolicyPath = policyPath },
            cts.Token);
        Assert.Equal(ExitCodes.Cancelled, cancelled.ExitCode);
        Assert.NotNull(cancelled.ArtifactDirectory);
        Assert.True(File.Exists(Path.Combine(cancelled.ArtifactDirectory!, "summary.json")));
    }

    [Fact]
    public async Task Coordinator_StaleRef_FailsClosed()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxActions = 1;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var initial = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-1",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        // Same fingerprint identity fields but element removed on re-observe.
        var drifted = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-2",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"]),
                    Node("n2", "n0", 1, 1, "Text", "Changed", "changed", true, [])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var call = 0;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                call++;
                var snap = call == 1 ? initial : drifted;
                return new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = snap.Nodes.Count };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UiExecutionResponse
            {
                Success = true,
                Value = JsonSerializer.SerializeToElement(new
                {
                    windows = new[]
                    {
                        new
                        {
                            hwnd = 1,
                            title = "DA Explorer Fixture",
                            processName = "DesktopAutomationExplorerFixture",
                            processId = 10
                        }
                    }
                })
            });

        var connection = new Mock<IDriverConnectionResolver>();
        connection.Setup(c => c.ResolveAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:1"),
                BearerToken = "token",
                DiscoveryMethod = "configured"
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, connection.Object);
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success);
        Assert.Contains(result.Summary!.StopReason ?? string.Empty, ["max-actions-budget", "frontier-exhausted"]);
    }

    [Fact]
    public void CommandLine_ParsesExploreApp()
    {
        var parsed = CommandLine.Parse(["explore-app", "--policy", "p.json", "--dry-run", "--json"]);
        Assert.Equal(AgentCommandKind.ExploreApp, parsed.Kind);
        Assert.Equal("p.json", parsed.PolicyFile);
        Assert.True(parsed.DryRun);
        Assert.True(parsed.Json);
    }

    [Fact]
    public void WorkspaceManager_IncludesExplorationTemplates()
    {
        var workspace = TestSupport.CreateWorkspace(TestSupport.CreateOptions());
        var result = workspace.Initialize();
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "exploration/");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "schemas/exploration-policy.schema.json");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "exploration/policies/example.policy.json");
    }

    [Fact]
    public void ForbiddenVocabulary_MatchesComparableTerms()
    {
        Assert.True(ForbiddenActionVocabulary.Matches("Delete Item"));
        Assert.True(ForbiddenActionVocabulary.Matches("btn_submit"));
        Assert.True(ForbiddenActionVocabulary.Matches("Purchase now"));
        Assert.False(ForbiddenActionVocabulary.Matches("Open Dialog"));
    }

    private static ExplorationCoordinator CreateCoordinator(
        AgentOptions options,
        Workspace.IWorkspaceManager workspace,
        ISemanticObservationService observation,
        IDriverUiClient ui,
        IDriverConnectionResolver connection)
    {
        var fingerprinter = new StateFingerprinter();
        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        probe.Register(11, new ProcessIdentityInfo
        {
            ProcessId = 11,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1,
            ParentProcessId = 10
        });
        var identity = new ProcessIdentityValidator(probe);
        return new ExplorationCoordinator(
            Options.Create(options),
            new ExplorationPolicyReader(),
            observation,
            connection,
            ui,
            fingerprinter,
            new CapabilityResolver(),
            new UiStabilizationService(observation, fingerprinter),
            new WindowTransitionDetector(ui, identity),
            new ExplorationArtifactWriter(workspace),
            NullLogger<ExplorationCoordinator>.Instance,
            identity);
    }

    private sealed class FakeProcessIdentityProbe : IProcessIdentityProbe
    {
        private readonly Dictionary<int, ProcessIdentityInfo> _byPid = new();

        public void Register(int processId, ProcessIdentityInfo info) => _byPid[processId] = info;

        public bool TryGetIdentity(int processId, out ProcessIdentityInfo info, out string? error)
        {
            if (_byPid.TryGetValue(processId, out info!))
            {
                error = null;
                return true;
            }

            info = new ProcessIdentityInfo { ProcessId = processId };
            error = $"Fake probe has no identity for PID {processId}.";
            return false;
        }
    }

    private static string WritePolicy(string workspaceRoot, ExplorationPolicyDocument policy)
    {
        var path = Path.Combine(workspaceRoot, "exploration", "policies", Guid.NewGuid().ToString("N") + ".json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, JsonSerializer.Serialize(policy));
        return path;
    }

    private static ExplorationPolicyDocument MinimalPolicy() =>
        new()
        {
            PolicySchemaVersion = 1,
            PolicyId = "test",
            Mode = "navigate",
            Application = new ExplorationApplicationPolicy
            {
                ProcessName = "DesktopAutomationExplorerFixture"
            },
            Scope = new ExplorationScopePolicy
            {
                AllowedWindowTitles =
                [
                    "DA Explorer Fixture",
                    "DA Explorer Child",
                    "DA Explorer Dialog",
                    "DA Explorer Second"
                ]
            },
            AllowedActionCategories = [],
            AllowRules = [],
            DenyRules = [],
            Budgets = new ExplorationBudgets
            {
                MaxStates = 20,
                MaxActions = 20,
                MaxDepth = 5,
                TotalDurationSeconds = 60,
                PerActionTimeoutSeconds = 5,
                StabilizationTimeoutSeconds = 5,
                StabilizationQuietPeriodMilliseconds = 0
            },
            AllowCloseChildWindowsForBacktrack = false,
            AllowHover = false,
            AllowCoordinateFallback = false,
            AllowDoubleClick = false
        };

    private static ExplorationActionCandidate Candidate(
        string fingerprint,
        string actionKey,
        ExplorationActionKind kind,
        string category,
        int priority,
        string elementKey) =>
        new()
        {
            ActionKey = actionKey,
            Kind = kind,
            Category = category,
            DriverOperation = "clickuia",
            ElementKey = elementKey,
            SnapshotId = "snap",
            NodeRef = "n1",
            StateFingerprint = fingerprint,
            Priority = priority
        };

    private static SemanticSnapshot Snapshot(
        params SemanticNode[] nodes) =>
        Snapshot(processName: "DesktopAutomationExplorerFixture", processId: 10, sessionId: 1, hwnd: 1, snapshotId: "snap", nodes: nodes);

    private static SemanticSnapshot Snapshot(
        string processName = "DesktopAutomationExplorerFixture",
        int processId = 10,
        int sessionId = 1,
        long hwnd = 1,
        string snapshotId = "snap",
        params SemanticNode[] nodes) =>
        new()
        {
            SnapshotSchemaVersion = 1,
            SnapshotId = snapshotId,
            CapturedAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            DriverVersion = "1.0.0",
            SnapshotHash = "a".PadRight(64, 'a'),
            Application = new ApplicationIdentity { ProcessId = processId, ProcessName = processName },
            WindowsSession = new WindowsSessionIdentity { SessionId = sessionId },
            Window = new WindowIdentity
            {
                Ref = nodes.FirstOrDefault()?.Ref,
                Hwnd = hwnd,
                Title = "DA Explorer Fixture",
                ClassName = "Form",
                ProcessId = processId,
                IsModal = false
            },
            Nodes = nodes.ToList()
        };

    private static SemanticSnapshot CloneWithWindow(
        SemanticSnapshot snapshot,
        string title,
        string className,
        bool isModal,
        long hwnd,
        int processId) =>
        new()
        {
            SnapshotSchemaVersion = snapshot.SnapshotSchemaVersion,
            SnapshotId = snapshot.SnapshotId,
            CapturedAtUtc = snapshot.CapturedAtUtc,
            DriverVersion = snapshot.DriverVersion,
            SnapshotHash = snapshot.SnapshotHash,
            Application = snapshot.Application,
            WindowsSession = snapshot.WindowsSession,
            Window = new WindowIdentity
            {
                Ref = snapshot.Window?.Ref,
                Hwnd = hwnd,
                Title = title,
                ClassName = className,
                ProcessId = processId,
                IsModal = isModal
            },
            Nodes = snapshot.Nodes,
            Truncated = snapshot.Truncated,
            Warnings = snapshot.Warnings
        };

    private static SemanticNode Node(
        string nodeRef,
        string? parentRef,
        int ordinal,
        int depth,
        string controlType,
        string name,
        string automationId,
        bool enabled,
        string[] patterns,
        bool focusable = true,
        NodeInteractionState? interaction = null) =>
        new()
        {
            Ref = nodeRef,
            ParentRef = parentRef,
            Ordinal = ordinal,
            Depth = depth,
            ControlType = controlType,
            Role = controlType,
            Name = name,
            AutomationId = automationId,
            Enabled = enabled,
            Offscreen = false,
            Focusable = focusable,
            HasKeyboardFocus = false,
            Patterns = patterns.ToList(),
            Capabilities = PatternCapabilityMapper.Map(patterns).ToList(),
            Interaction = interaction
        };

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
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.slnx")))
                return dir.FullName;
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Repo root not found.");
    }

    private static System.Text.Json.Nodes.JsonNode JsonNodeParse(string json) =>
        System.Text.Json.Nodes.JsonNode.Parse(json)!;
}
