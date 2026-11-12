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

        // Probe failure must fail closed (never Ok).
        var probeFail = validator.ValidateNewWindow(
            new WindowDescriptor
            {
                Title = "Ghost",
                ProcessName = "DesktopAutomationExplorerFixture",
                ProcessId = 999
            },
            MinimalPolicy());
        Assert.False(probeFail.IsValid);

        // Missing parent fails child lineage.
        var noParentProbe = new FakeProcessIdentityProbe();
        noParentProbe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        noParentProbe.Register(20, new ProcessIdentityInfo
        {
            ProcessId = 20,
            ProcessName = "FixtureChild",
            SessionId = 1,
            ParentProcessId = null
        });
        var lineageValidator = new ProcessIdentityValidator(noParentProbe);
        var rootOk = lineageValidator.ValidateSnapshot(Snapshot(processId: 10), childPolicy);
        Assert.True(rootOk.IsValid);
        Assert.Equal(10, lineageValidator.PinnedRootProcessId);
        Assert.False(lineageValidator.ValidateSnapshot(
            Snapshot(processName: "FixtureChild", processId: 20),
            childPolicy).IsValid);

        // Pin is sticky — second root validation must not reset.
        Assert.True(lineageValidator.ValidateSnapshot(Snapshot(processId: 10), childPolicy).IsValid);
        Assert.Equal(10, lineageValidator.PinnedRootProcessId);

        // Same processName as root but different PID is NOT accepted without ancestry to the pin.
        var sameNameProbe = new FakeProcessIdentityProbe();
        sameNameProbe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        sameNameProbe.Register(55, new ProcessIdentityInfo
        {
            ProcessId = 55,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1,
            ParentProcessId = 1
        });
        var sameNameValidator = new ProcessIdentityValidator(sameNameProbe);
        Assert.True(sameNameValidator.ValidateSnapshot(Snapshot(processId: 10), MinimalPolicy()).IsValid);
        var sameNameDifferentPid = sameNameValidator.ValidateNewWindow(
            new WindowDescriptor
            {
                Title = "Other Instance",
                ProcessName = "DesktopAutomationExplorerFixture",
                ProcessId = 55,
                Hwnd = 99
            },
            MinimalPolicy());
        Assert.False(sameNameDifferentPid.IsValid);
        Assert.Contains("Unexpected external", sameNameDifferentPid.Detail ?? string.Empty, StringComparison.OrdinalIgnoreCase);
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

        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        var identity = new ProcessIdentityValidator(probe);
        Assert.True(identity.ValidateSnapshot(Snapshot(processId: 10), policy).IsValid);

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var result = detector.Compare(before, after, Snapshot(processId: 10), policy, identity);
        Assert.True(result.ScopeViolation);
        Assert.Contains("processName", result.Detail ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveLiveNode_UsesNodeRefOnlyWhenSnapshotIdMatches()
    {
        var snapshot = Snapshot(
            snapshotId: "snap-live",
            nodes:
            [
                Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
                Node("n1", "n0", 0, 1, "Button", "Open", "open", true, ["Invoke"])
            ]);

        var sameSnapshot = new ExplorationActionCandidate
        {
            ActionKey = "invoke:open",
            Kind = ExplorationActionKind.Invoke,
            Category = "invoke",
            ElementKey = "Button|open|Open|o0|d1",
            AutomationId = "open",
            Name = "Open",
            ControlType = "Button",
            NodeRef = "n1",
            SnapshotId = "snap-live"
        };
        Assert.True(ExplorationCoordinator.TryResolveLiveNode(snapshot, sameSnapshot, out var byRef, out _));
        Assert.Equal("n1", byRef!.Ref);

        var staleSnapshotId = new ExplorationActionCandidate
        {
            ActionKey = sameSnapshot.ActionKey,
            Kind = sameSnapshot.Kind,
            Category = sameSnapshot.Category,
            ElementKey = sameSnapshot.ElementKey,
            AutomationId = sameSnapshot.AutomationId,
            Name = sameSnapshot.Name,
            ControlType = sameSnapshot.ControlType,
            NodeRef = "n1",
            SnapshotId = "snap-other"
        };
        // NodeRef must be ignored when SnapshotId differs — fall through to identity match.
        Assert.True(ExplorationCoordinator.TryResolveLiveNode(snapshot, staleSnapshotId, out var byIdentity, out _));
        Assert.Equal("n1", byIdentity!.Ref);

        var ambiguous = Snapshot(
            snapshotId: "snap-live",
            nodes:
            [
                Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
                Node("n1", "n0", 0, 1, "Button", "Same", "dup", true, ["Invoke"]),
                Node("n2", "n0", 1, 1, "Button", "Same", "dup", true, ["Invoke"])
            ]);
        var ambiguousAction = new ExplorationActionCandidate
        {
            ActionKey = "invoke:x",
            Kind = ExplorationActionKind.Invoke,
            Category = "invoke",
            ElementKey = "Button|dup|Same|o0|d1",
            AutomationId = "dup",
            Name = "Same",
            ControlType = "Button",
            NodeRef = "n1",
            SnapshotId = "snap-other"
        };
        Assert.False(ExplorationCoordinator.TryResolveLiveNode(ambiguous, ambiguousAction, out _, out var detail));
        Assert.Contains("Ambiguous", detail, StringComparison.OrdinalIgnoreCase);
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
        Assert.True(frontier.HasExecuted(a));
    }

    [Fact]
    public void FrontierQueue_IdenticalSiblingCandidates_DoNotCollapse()
    {
        var frontier = new FrontierQueue();
        // Same fingerprint + action identity; only TargetHwnd differs (identical sibling windows).
        var left = Candidate(
            "same-fp",
            "Invoke:Button|ExplorerSiblingPing|Ping|o1|d1",
            ExplorationActionKind.Invoke,
            "invoke",
            priority: 80,
            elementKey: "Button|ExplorerSiblingPing|Ping|o1|d1",
            targetHwnd: 2,
            targetProcessId: 10);
        var right = Candidate(
            "same-fp",
            "Invoke:Button|ExplorerSiblingPing|Ping|o1|d1",
            ExplorationActionKind.Invoke,
            "invoke",
            priority: 80,
            elementKey: "Button|ExplorerSiblingPing|Ping|o1|d1",
            targetHwnd: 3,
            targetProcessId: 10);

        Assert.True(frontier.TryEnqueue(left));
        Assert.True(frontier.TryEnqueue(right));
        Assert.Equal(2, frontier.Count);
        Assert.False(frontier.TryEnqueue(left));

        Assert.True(frontier.TryDequeue(out var first));
        Assert.True(frontier.TryDequeue(out var second));
        Assert.Equal(2L, first!.TargetHwnd);
        Assert.Equal(3L, second!.TargetHwnd);
        Assert.Equal(left.ActionKey, second.ActionKey);
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

        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        var identity = new ProcessIdentityValidator(probe);
        Assert.True(identity.ValidateSnapshot(Snapshot(processId: 10), policy).IsValid);

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var external = detector.Compare(before, afterExternal, Snapshot(processId: 10), policy, identity);
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
        var stableNoise = detector.Compare(
            preexistingNoise,
            preexistingNoise,
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                processId: 10,
                nodes: [Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "main", true, ["Window"])]),
            policy,
            identity);
        Assert.False(stableNoise.ScopeViolation);

        var afterModal = new WindowInventory { Windows = [before.Windows[0]] };
        var modalSnapshot = CloneWithWindow(Snapshot(processId: 10), "DA Explorer Dialog", "Dialog", isModal: true, hwnd: 3, processId: 10);
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
        var modal = detector.Compare(before, afterModal, modalSnapshot, policy, identity);
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
        await File.WriteAllTextAsync(policyPath, JsonSerializer.Serialize(policy, CamelCaseJson));

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

    [Fact]
    public void PolicyReader_SchemaRejectsMisspelledMaxActions()
    {
        var json = """
            {
              "policySchemaVersion": 1,
              "mode": "observe",
              "application": { "processName": "DesktopAutomationExplorerFixture" },
              "scope": {},
              "allowedActionCategories": [],
              "allowRules": [],
              "denyRules": [],
              "budgets": {
                "maxStates": 5,
                "maxActons": 3,
                "maxDepth": 2,
                "totalDurationSeconds": 30,
                "perActionTimeoutSeconds": 5,
                "stabilizationTimeoutSeconds": 5
              },
              "allowCloseChildWindowsForBacktrack": false,
              "allowHover": false,
              "allowCoordinateFallback": false
            }
            """;
        var result = new ExplorationPolicyReader().ValidateJsonAgainstSchema(json);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("schema", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void CapabilityResolver_MenuOpenRequiresExpandCollapse_NoAutomationIdHeuristic()
    {
        var snapshot = Snapshot(
            Node("n0", null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node("n1", "n0", 0, 1, "MenuItem", "File", "ExplorerFileMenu", true, ["Invoke"]));

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
            }
        ];

        var candidates = new CapabilityResolver().Resolve(snapshot, "fp", policy, 0);
        Assert.DoesNotContain(candidates, c => c.Kind == ExplorationActionKind.MenuOpen);
    }

    [Fact]
    public void ComputeHash_IncludesInteractionState()
    {
        var collapsed = Snapshot(
            Node(
                "n1",
                null,
                0,
                0,
                "TreeItem",
                "Root",
                "root",
                true,
                ["ExpandCollapse"],
                interaction: new NodeInteractionState { ExpandCollapse = "Collapsed" }));
        var expanded = Snapshot(
            Node(
                "n1",
                null,
                0,
                0,
                "TreeItem",
                "Root",
                "root",
                true,
                ["ExpandCollapse"],
                interaction: new NodeInteractionState { ExpandCollapse = "Expanded" }));

        Assert.NotEqual(
            SemanticSnapshotNormalizer.ComputeHash(collapsed),
            SemanticSnapshotNormalizer.ComputeHash(expanded));
    }

    [Fact]
    public async Task Coordinator_MaxDepthZero_DoesNotExecuteActions()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxDepth = 0;
        policy.Budgets.MaxActions = 10;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var snapshot = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
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

        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ObservationResult { Success = true, ExitCode = 0, Snapshot = snapshot, NodeCount = 2 });

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
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success);
        Assert.Equal(0, result.Summary!.ActionsAttempted);
        ui.Verify(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Coordinator_MaxDepthOne_ExecutesOnlyDepthZero()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxDepth = 1;
        policy.Budgets.MaxActions = 20;
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
                var snap = observeCalls <= 2 ? stateA : stateB;
                return new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = snap.Nodes.Count };
            });

        var mutateCalls = 0;
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

                mutateCalls++;
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
        Assert.True(result.Success, result.Error ?? result.Summary?.StopReason ?? "failed");
        Assert.True(mutateCalls >= 1);
        Assert.True(result.Summary!.MaxDepthReached <= 1);
        Assert.True(result.Summary.ActionsAttempted >= 1);

        // Every attempted action must have been at depth 0 (depth >= maxDepth never executes).
        var transitionLines = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l));
        foreach (var line in transitionLines)
        {
            using var doc = JsonDocument.Parse(line);
            if (doc.RootElement.TryGetProperty("depth", out var depthProp))
                Assert.True(depthProp.GetInt32() < policy.Budgets.MaxDepth);
        }
    }

    [Fact]
    public async Task Coordinator_RecordsStabilizationTimeout_WhenFingerprintNeverQuiets()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxActions = 1;
        policy.Budgets.StabilizationTimeoutSeconds = 1;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 500;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var tick = 0;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                tick++;
                // Initial + pre-action share identity; stabilization polls alternate node names so FP never quiets.
                var name = tick <= 2 ? "Root" : $"Root-{tick}";
                var snap = CloneWithWindow(
                    Snapshot(
                        processName: "DesktopAutomationExplorerFixture",
                        snapshotId: $"snap-{tick}",
                        nodes:
                        [
                            Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                            Node("n1", "n0", 0, 1, "TreeItem", name, "ExplorerTreeRoot", true, ["ExpandCollapse"])
                        ]),
                    "DA Explorer Fixture",
                    "Form",
                    false,
                    1,
                    10);
                return new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = 2 };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
                step.Operation == "listwindows"
                    ? ListWindowsResponse([(1L, "DA Explorer Fixture", 10)])
                    : new UiExecutionResponse { Success = true });

        var connection = MockConnection();
        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, connection);
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success);
        var transitions = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => JsonDocument.Parse(l).RootElement.GetProperty("outcome").GetString())
            .ToList();
        Assert.Contains("stabilizationTimeout", transitions);
    }

    [Fact]
    public async Task Coordinator_StopsAtExactMaxStatesBudget()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.MaxStates = 2;
        policy.Budgets.MaxActions = 50;
        policy.Budgets.MaxDepth = 5;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" },
            new ExplorationElementRule { AutomationId = "ExplorerTreeChild", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var observeCalls = 0;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                observeCalls++;
                var snap = CloneWithWindow(
                    Snapshot(
                        processName: "DesktopAutomationExplorerFixture",
                        snapshotId: $"snap-{observeCalls}",
                        nodes:
                        [
                            Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                            Node("n1", "n0", 0, 1, "TreeItem", $"Root-{observeCalls}", "ExplorerTreeRoot", true, ["ExpandCollapse"]),
                            Node("n2", "n1", 0, 2, "TreeItem", $"Child-{observeCalls}", "ExplorerTreeChild", true, ["ExpandCollapse"])
                        ]),
                    "DA Explorer Fixture",
                    "Form",
                    false,
                    1,
                    10);
                return new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = snap.Nodes.Count };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
                step.Operation == "listwindows"
                    ? ListWindowsResponse([(1L, "DA Explorer Fixture", 10)])
                    : new UiExecutionResponse { Success = true });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success);
        Assert.Equal("max-states-budget", result.Summary!.StopReason);
        Assert.Equal(policy.Budgets.MaxStates, result.Summary.StatesVisited);
    }

    [Fact]
    public async Task Coordinator_StopsOnTotalDurationBudget()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.Budgets.TotalDurationSeconds = 1;
        policy.Budgets.MaxActions = 50;
        policy.Budgets.PerActionTimeoutSeconds = 5;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule { AutomationId = "ExplorerTreeRoot", Action = "expand", ActionCategory = "expandCollapse" }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var snap = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
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

        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = 2 });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .Returns(async (DriverConnection _, PlanStep step, CancellationToken ct) =>
            {
                if (step.Operation != "listwindows")
                    await Task.Delay(1200, ct);
                return ListWindowsResponse([(1L, "DA Explorer Fixture", 10)]);
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success, result.Error ?? result.Summary?.StopReason ?? "failed");
        Assert.Equal("total-duration-budget", result.Summary!.StopReason);
        Assert.Equal("budget-exhausted", result.Summary.Status);
    }

    [Fact]
    public async Task Coordinator_Backtrack_UsesHwndWindowPattern_AndPreservesDepth()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowCloseChildWindowsForBacktrack = true;
        policy.Budgets.MaxActions = 20;
        policy.Budgets.MaxDepth = 3;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenSecondButton",
                Action = "invoke",
                ActionCategory = "invoke"
            }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var main = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-main",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Second", "ExplorerOpenSecondButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        // Child surface has no further permitted actions so frontier empties → backtrack.
        var child = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-child",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Second", "ExplorerSecondWindow", true, ["Window"])
                ]),
            "DA Explorer Second",
            "Form",
            false,
            2,
            10);

        var activeHwnd = 1L;
        var childOpen = false;
        var childClosed = false;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                var snap = activeHwnd == 2 && childOpen && !childClosed ? child : main;
                return new ObservationResult
                {
                    Success = true,
                    ExitCode = 0,
                    Snapshot = snap,
                    NodeCount = snap.Nodes.Count
                };
            });

        var operations = new List<string>();
        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                operations.Add(step.Operation);
                if (step.Operation == "listwindows")
                {
                    return childOpen && !childClosed
                        ? ListWindowsResponse(
                        [
                            (1L, "DA Explorer Fixture", 10),
                            (2L, "DA Explorer Second", 10)
                        ])
                        : ListWindowsResponse([(1L, "DA Explorer Fixture", 10)]);
                }

                if (step.Operation == "clickuia")
                {
                    childOpen = true;
                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "switchwindow")
                {
                    Assert.True(
                        step.Arguments is not null
                        && step.Arguments.ContainsKey("hwnd")
                        && step.Arguments.ContainsKey("processId"),
                        "Backtrack/NewWindow switch must use hwnd+processId.");
                    activeHwnd = step.Arguments!["hwnd"].GetInt64();
                    Assert.Equal(10, step.Arguments["processId"].GetInt32());
                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "closewindow")
                {
                    Assert.True(
                        step.Arguments is not null
                        && step.Arguments.ContainsKey("hwnd")
                        && step.Arguments.ContainsKey("processId"),
                        "Backtrack must use HWND+PID targeted closewindow.");
                    Assert.False(
                        step.Arguments.ContainsKey("value"),
                        "Backtrack must not use title-based closewindow.");
                    var closedHwnd = step.Arguments["hwnd"].GetInt64();
                    Assert.Equal(2, closedHwnd);
                    Assert.Equal(10, step.Arguments["processId"].GetInt32());
                    childClosed = true;
                    activeHwnd = 1;
                    return new UiExecutionResponse
                    {
                        Success = true,
                        Value = JsonSerializer.SerializeToElement(new
                        {
                            operation = "closewindow",
                            hwnd = closedHwnd,
                            processId = 10,
                            closed = true,
                            method = "windowPattern"
                        })
                    };
                }

                return new UiExecutionResponse { Success = true };
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success, result.Error ?? result.Summary?.StopReason ?? "failed");
        Assert.Contains("closewindow", operations);
        Assert.Contains("switchwindow", operations);

        var transitionLines = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => JsonDocument.Parse(l).RootElement)
            .ToList();

        Assert.Contains(transitionLines, t =>
            t.TryGetProperty("outcome", out var o)
            && o.GetString() == "newWindow");

        var backtrack = transitionLines.FirstOrDefault(t =>
            t.TryGetProperty("actionKey", out var key)
            && (key.GetString()?.StartsWith("backtrack:close:", StringComparison.Ordinal) ?? false));
        Assert.False(backtrack.ValueKind == JsonValueKind.Undefined);
        Assert.Equal("closedWindow", backtrack.GetProperty("outcome").GetString());
        Assert.Equal(1, backtrack.GetProperty("depth").GetInt32());
        Assert.Contains(
            "WindowPattern",
            backtrack.GetProperty("detail").GetString() ?? string.Empty,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Coordinator_NewWindow_IndependentOfNoOpFingerprint_SwitchesAndEnqueues()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowCloseChildWindowsForBacktrack = false;
        policy.Budgets.MaxActions = 5;
        policy.Budgets.MaxDepth = 2;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["invoke", "expandCollapse"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenSecondButton",
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
        var policyPath = WritePolicy(workspace.RootPath, policy);

        // Parent fingerprint intentionally unchanged across the invoke (NoOp FP), but a new HWND appears.
        var main = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-main",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Second", "ExplorerOpenSecondButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var child = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-child",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Second", "ExplorerSecondWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "TreeItem", "Root", "ExplorerTreeRoot", true, ["ExpandCollapse"])
                ]),
            "DA Explorer Second",
            "Form",
            false,
            2,
            10);

        var observeCalls = 0;
        var switched = false;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                observeCalls++;
                var snap = switched ? child : main;
                return new ObservationResult { Success = true, ExitCode = 0, Snapshot = snap, NodeCount = snap.Nodes.Count };
            });

        var childOpen = false;
        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                if (step.Operation == "listwindows")
                {
                    return childOpen
                        ? ListWindowsResponse(
                        [
                            (1L, "DA Explorer Fixture", 10),
                            (2L, "DA Explorer Second", 10)
                        ])
                        : ListWindowsResponse([(1L, "DA Explorer Fixture", 10)]);
                }

                if (step.Operation == "clickuia")
                {
                    childOpen = true;
                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "switchwindow")
                {
                    Assert.True(step.Arguments!.ContainsKey("hwnd"));
                    Assert.True(step.Arguments.ContainsKey("processId"));
                    Assert.Equal(2, step.Arguments["hwnd"].GetInt64());
                    Assert.Equal(10, step.Arguments["processId"].GetInt32());
                    switched = true;
                    return new UiExecutionResponse { Success = true };
                }

                return new UiExecutionResponse { Success = true };
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success, result.Error ?? result.Summary?.StopReason ?? "failed");
        Assert.True(switched);

        var outcomes = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => JsonDocument.Parse(l).RootElement.GetProperty("outcome").GetString())
            .ToList();
        Assert.Contains("newWindow", outcomes);
        // The open transition must be NewWindow even though the parent fingerprint was unchanged.
        Assert.Equal("newWindow", outcomes[0]);
    }

    [Fact]
    public void WindowTransitionDetector_HwndReuseWithDifferentPid_SurfacesAsRemoveAndAdd()
    {
        var policy = MinimalPolicy();
        policy.Application.AllowedChildProcessNames = ["FixtureChild"];
        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        probe.Register(99, new ProcessIdentityInfo
        {
            ProcessId = 99,
            ProcessName = "FixtureChild",
            SessionId = 1,
            ParentProcessId = 10
        });
        var identity = new ProcessIdentityValidator(probe);
        Assert.True(identity.ValidateSnapshot(Snapshot(processId: 10), policy).IsValid);

        var before = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 42,
                    Title = "DA Explorer Fixture",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = 10
                }
            ]
        };
        var afterReuse = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 42,
                    Title = "DA Explorer Child",
                    ProcessName = "FixtureChild",
                    ProcessId = 99
                }
            ]
        };

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var result = detector.Compare(
            before,
            afterReuse,
            Snapshot(processId: 10),
            policy,
            identity);

        Assert.False(result.ScopeViolation);
        Assert.True(result.NewWindow);
        Assert.True(result.ClosedWindow);
        Assert.Single(result.AddedWindows);
        Assert.Equal(42, result.AddedWindows[0].Hwnd);
        Assert.Equal(99, result.AddedWindows[0].ProcessId);
    }

    [Fact]
    public async Task Coordinator_SameHwndDifferentPid_RetargetsAndRejectsMismatch()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowCloseChildWindowsForBacktrack = false;
        policy.Budgets.MaxActions = 5;
        policy.Budgets.MaxDepth = 2;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenSecondButton",
                Action = "invoke",
                ActionCategory = "invoke"
            }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        // Candidate was resolved against HWND 1 / PID 10, but live observe now shows
        // the same HWND owned by PID 99 (handle reuse) — retarget must run and fail closed.
        var expected = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                processId: 10,
                snapshotId: "snap-expected",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Second", "ExplorerOpenSecondButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        // Application PID stays 10 (scope-valid); Window PID flips to 99 (HWND reuse).
        var reused = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                processId: 10,
                snapshotId: "snap-reused",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Second", "ExplorerOpenSecondButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            99);

        var observeCalls = 0;
        var switchAttempted = false;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                observeCalls++;
                // First observation seeds the frontier from the expected identity.
                // Subsequent pre-action observes simulate HWND reuse under a new PID.
                var snap = observeCalls == 1 ? expected : reused;
                return new ObservationResult
                {
                    Success = true,
                    ExitCode = 0,
                    Snapshot = snap,
                    NodeCount = snap.Nodes.Count
                };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                if (step.Operation == "listwindows")
                {
                    // After reuse the live inventory advertises HWND 1 under PID 99.
                    return ListWindowsResponse([(1L, "DA Explorer Fixture", 99)]);
                }

                if (step.Operation == "switchwindow")
                {
                    switchAttempted = true;
                    Assert.Equal(1, step.Arguments!["hwnd"].GetInt64());
                    Assert.Equal(10, step.Arguments["processId"].GetInt32());
                    // Switch "succeeds" at the driver layer; inventory PID check must fail closed.
                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "clickuia")
                    throw new InvalidOperationException("Click must not run after HWND/PID mismatch.");

                return new UiExecutionResponse { Success = true };
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(switchAttempted, "Retarget must switch when HWND matches but PID differs.");

        var transitions = File.ReadAllLines(Path.Combine(result.ArtifactDirectory!, "transitions.jsonl"))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => JsonDocument.Parse(l).RootElement)
            .ToList();
        Assert.Contains(
            transitions,
            t => string.Equals(t.GetProperty("outcome").GetString(), "staleRef", StringComparison.Ordinal));
        Assert.DoesNotContain(
            transitions,
            t => string.Equals(t.GetProperty("outcome").GetString(), "success", StringComparison.Ordinal)
                 || string.Equals(t.GetProperty("outcome").GetString(), "newWindow", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Coordinator_TwoSiblingWindows_RetargetsHwndBeforeEachUniqueAction()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowCloseChildWindowsForBacktrack = false;
        policy.Budgets.MaxActions = 12;
        policy.Budgets.MaxDepth = 3;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenBothButton",
                Action = "invoke",
                ActionCategory = "invoke"
            },
            new ExplorationElementRule
            {
                AutomationId = "ExplorerSiblingAPing",
                Action = "invoke",
                ActionCategory = "invoke"
            },
            new ExplorationElementRule
            {
                AutomationId = "ExplorerSiblingBPing",
                Action = "invoke",
                ActionCategory = "invoke"
            }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var main = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-main",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Both", "ExplorerOpenBothButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var siblingA = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-a",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Child", "ExplorerChildWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Ping A", "ExplorerSiblingAPing", true, ["Invoke"])
                ]),
            "DA Explorer Child",
            "Form",
            false,
            2,
            10);

        var siblingB = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-b",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Second", "ExplorerSecondWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Ping B", "ExplorerSiblingBPing", true, ["Invoke"])
                ]),
            "DA Explorer Second",
            "Form",
            false,
            3,
            10);

        long activeHwnd = 1;
        var bothOpen = false;
        var pinged = new HashSet<string>(StringComparer.Ordinal);
        var switchedHwnds = new HashSet<long>();

        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() =>
            {
                var snap = activeHwnd switch
                {
                    2 when bothOpen => siblingA,
                    3 when bothOpen => siblingB,
                    _ => main
                };
                return new ObservationResult
                {
                    Success = true,
                    ExitCode = 0,
                    Snapshot = snap,
                    NodeCount = snap.Nodes.Count
                };
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                if (step.Operation == "listwindows")
                {
                    return bothOpen
                        ? ListWindowsResponse(
                        [
                            (1L, "DA Explorer Fixture", 10),
                            (2L, "DA Explorer Child", 10),
                            (3L, "DA Explorer Second", 10)
                        ])
                        : ListWindowsResponse([(1L, "DA Explorer Fixture", 10)]);
                }

                if (step.Operation == "clickuia")
                {
                    var aid = TryReadLocatorAutomationId(step);
                    if (string.Equals(aid, "ExplorerOpenBothButton", StringComparison.Ordinal))
                    {
                        bothOpen = true;
                        return new UiExecutionResponse { Success = true };
                    }

                    if (string.Equals(aid, "ExplorerSiblingAPing", StringComparison.Ordinal))
                    {
                        Assert.Equal(2, activeHwnd);
                        pinged.Add("A");
                        return new UiExecutionResponse { Success = true };
                    }

                    if (string.Equals(aid, "ExplorerSiblingBPing", StringComparison.Ordinal))
                    {
                        Assert.Equal(3, activeHwnd);
                        pinged.Add("B");
                        return new UiExecutionResponse { Success = true };
                    }

                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "switchwindow")
                {
                    Assert.True(step.Arguments!.ContainsKey("hwnd"));
                    Assert.True(step.Arguments.ContainsKey("processId"));
                    Assert.Equal(10, step.Arguments["processId"].GetInt32());
                    activeHwnd = step.Arguments["hwnd"].GetInt64();
                    switchedHwnds.Add(activeHwnd);
                    return new UiExecutionResponse { Success = true };
                }

                return new UiExecutionResponse { Success = true };
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success, result.Error ?? result.Summary?.StopReason ?? "failed");
        Assert.Contains(2L, switchedHwnds);
        Assert.Contains(3L, switchedHwnds);
        Assert.Contains("A", pinged);
        Assert.Contains("B", pinged);
    }

    [Fact]
    public async Task Coordinator_BudgetExhaustedDuringNewWindowBind_ReportsBudgetNotBindFailed()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policy = MinimalPolicy();
        policy.Mode = "navigate";
        policy.AllowCloseChildWindowsForBacktrack = false;
        // Click consumes the only action slot; bind must stop with max-actions-budget.
        policy.Budgets.MaxActions = 1;
        policy.Budgets.MaxDepth = 2;
        policy.Budgets.StabilizationQuietPeriodMilliseconds = 0;
        policy.AllowedActionCategories = ["invoke"];
        policy.AllowRules =
        [
            new ExplorationElementRule
            {
                AutomationId = "ExplorerOpenSecondButton",
                Action = "invoke",
                ActionCategory = "invoke"
            }
        ];
        var policyPath = WritePolicy(workspace.RootPath, policy);

        var main = CloneWithWindow(
            Snapshot(
                processName: "DesktopAutomationExplorerFixture",
                snapshotId: "snap-main",
                nodes:
                [
                    Node("n0", null, 0, 0, "Window", "DA Explorer Fixture", "ExplorerMainWindow", true, ["Window"]),
                    Node("n1", "n0", 0, 1, "Button", "Open Second", "ExplorerOpenSecondButton", true, ["Invoke"])
                ]),
            "DA Explorer Fixture",
            "Form",
            false,
            1,
            10);

        var childOpen = false;
        var observe = new Mock<ISemanticObservationService>();
        observe.Setup(o => o.ObserveAsync(It.IsAny<ObservationOptions>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() => new ObservationResult
            {
                Success = true,
                ExitCode = 0,
                Snapshot = main,
                NodeCount = main.Nodes.Count
            });

        var ui = new Mock<IDriverUiClient>();
        ui.Setup(u => u.ExecuteStepAsync(It.IsAny<DriverConnection>(), It.IsAny<PlanStep>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DriverConnection _, PlanStep step, CancellationToken _) =>
            {
                if (step.Operation == "listwindows")
                {
                    return childOpen
                        ? ListWindowsResponse(
                        [
                            (1L, "DA Explorer Fixture", 10),
                            (2L, "DA Explorer Second", 10)
                        ])
                        : ListWindowsResponse([(1L, "DA Explorer Fixture", 10)]);
                }

                if (step.Operation == "clickuia")
                {
                    childOpen = true;
                    return new UiExecutionResponse { Success = true };
                }

                if (step.Operation == "switchwindow")
                    throw new InvalidOperationException("Bind switch should not run after budget.");

                return new UiExecutionResponse { Success = true };
            });

        var coordinator = CreateCoordinator(options, workspace, observe.Object, ui.Object, MockConnection());
        var result = await coordinator.ExploreAsync(new ExplorationRequest { PolicyPath = policyPath });
        Assert.True(result.Success, result.Error ?? "expected budget success exit");
        Assert.Equal("max-actions-budget", result.Summary?.StopReason);
        Assert.NotEqual("new-window-bind-failed", result.Summary?.StopReason);
    }

    private static string? TryReadLocatorAutomationId(PlanStep step)
    {
        if (step.Arguments is null || !step.Arguments.TryGetValue("locator", out var locator))
            return null;
        if (locator.ValueKind != JsonValueKind.Object)
            return null;
        if (locator.TryGetProperty("automationId", out var aid)
            || locator.TryGetProperty("AutomationId", out aid))
        {
            return aid.ValueKind == JsonValueKind.String ? aid.GetString() : null;
        }

        return null;
    }

    [Fact]
    public void WindowTransitionDetector_MissingProcessId_FailsClosed_AndIsNotIgnored()
    {
        var policy = MinimalPolicy();
        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        var identity = new ProcessIdentityValidator(probe);
        Assert.True(identity.ValidateSnapshot(Snapshot(processId: 10), policy).IsValid);

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
        var afterMissingPid = new WindowInventory
        {
            Windows =
            [
                before.Windows[0],
                new WindowDescriptor
                {
                    Hwnd = 99,
                    Title = "Elevated Mystery",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = null
                }
            ]
        };

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var result = detector.Compare(
            before,
            afterMissingPid,
            Snapshot(processId: 10),
            policy,
            identity);

        Assert.True(result.ScopeViolation);
        Assert.True(result.NewWindow);
        Assert.Contains("processId", result.Detail ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("fail closed", result.Detail ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WindowTransitionDetector_TitleChangeAlone_IsNeitherNewNorClosedWindow()
    {
        var policy = MinimalPolicy();
        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        var identity = new ProcessIdentityValidator(probe);
        Assert.True(identity.ValidateSnapshot(Snapshot(processId: 10), policy).IsValid);

        var before = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 42,
                    Title = "DA Explorer Fixture",
                    ClassName = "Form",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = 10
                }
            ]
        };
        var afterTitleOnly = new WindowInventory
        {
            Windows =
            [
                new WindowDescriptor
                {
                    Hwnd = 42,
                    Title = "DA Explorer Fixture — renamed",
                    ClassName = "Form",
                    ProcessName = "DesktopAutomationExplorerFixture",
                    ProcessId = 10
                }
            ]
        };

        var detector = new WindowTransitionDetector(Mock.Of<IDriverUiClient>());
        var result = detector.Compare(
            before,
            afterTitleOnly,
            Snapshot(processId: 10),
            policy,
            identity);

        Assert.False(result.ScopeViolation);
        Assert.False(result.NewWindow);
        Assert.False(result.ClosedWindow);
        Assert.Empty(result.AddedWindows);
    }

    [Fact]
    public void ProcessIdentityValidator_SessionPinning_FailClosed_OnMissingOrDisagreeingSession()
    {
        var policy = MinimalPolicy();
        var probe = new FakeProcessIdentityProbe();
        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = null
        });
        var missingLive = new ProcessIdentityValidator(probe);
        Assert.False(missingLive.ValidateSnapshot(Snapshot(processId: 10, sessionId: 1), policy).IsValid);

        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 7
        });
        var disagree = new ProcessIdentityValidator(probe);
        Assert.False(disagree.ValidateSnapshot(Snapshot(processId: 10, sessionId: 1), policy).IsValid);

        probe.Register(10, new ProcessIdentityInfo
        {
            ProcessId = 10,
            ProcessName = "DesktopAutomationExplorerFixture",
            SessionId = 1
        });
        var ok = new ProcessIdentityValidator(probe);
        Assert.True(ok.ValidateSnapshot(Snapshot(processId: 10, sessionId: 1), policy).IsValid);
        Assert.Equal(1, ok.PinnedRootSessionId);

        // Child/new-window must match pinned session.
        probe.Register(11, new ProcessIdentityInfo
        {
            ProcessId = 11,
            ProcessName = "FixtureChild",
            ParentProcessId = 10,
            SessionId = 2
        });
        policy.Application.AllowedChildProcessNames = ["FixtureChild"];
        Assert.False(ok.ValidateNewWindow(
            new WindowDescriptor
            {
                Hwnd = 99,
                Title = "Child",
                ProcessName = "FixtureChild",
                ProcessId = 11
            },
            policy).IsValid);

        probe.Register(11, new ProcessIdentityInfo
        {
            ProcessId = 11,
            ProcessName = "FixtureChild",
            ParentProcessId = 10,
            SessionId = 1
        });
        Assert.True(ok.ValidateNewWindow(
            new WindowDescriptor
            {
                Hwnd = 99,
                Title = "Child",
                ProcessName = "FixtureChild",
                ProcessId = 11
            },
            policy).IsValid);
    }

    [Fact]
    public void ExplorationWindowStack_RequiresChildHwndAndProcessId()
    {
        var stack = new ExplorationWindowStack();
        Assert.ThrowsAny<Exception>(() => stack.PushChild(new ExplorationWindowFrame
        {
            ChildHwnd = 0,
            ChildProcessId = 10,
            Depth = 1
        }));
        Assert.ThrowsAny<Exception>(() => stack.PushChild(new ExplorationWindowFrame
        {
            ChildHwnd = 1,
            ChildProcessId = 0,
            Depth = 1
        }));
        stack.PushChild(new ExplorationWindowFrame
        {
            ChildHwnd = 2,
            ChildProcessId = 10,
            ParentHwnd = 1,
            Depth = 1
        });
        Assert.True(stack.TryPeek(out var frame));
        Assert.Equal(2, frame!.ChildHwnd);
        Assert.Equal(1, frame.Depth);
    }

    private static IDriverConnectionResolver MockConnection()
    {
        var connection = new Mock<IDriverConnectionResolver>();
        connection.Setup(c => c.ResolveAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:1"),
                BearerToken = "token",
                DiscoveryMethod = "configured"
            });
        return connection.Object;
    }

    private static UiExecutionResponse ListWindowsResponse(
        (long Hwnd, string Title, int ProcessId)[] windows) =>
        new()
        {
            Success = true,
            Value = JsonSerializer.SerializeToElement(new
            {
                windows = windows.Select(w => new
                {
                    hwnd = w.Hwnd,
                    title = w.Title,
                    className = "Form",
                    processId = w.ProcessId,
                    processName = "DesktopAutomationExplorerFixture"
                }).ToArray()
            })
        };

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
            ProcessName = "FixtureChild",
            SessionId = 1,
            ParentProcessId = 10
        });
        return new ExplorationCoordinator(
            Options.Create(options),
            new ExplorationPolicyReader(),
            observation,
            connection,
            ui,
            fingerprinter,
            new CapabilityResolver(),
            new UiStabilizationService(observation, fingerprinter),
            new WindowTransitionDetector(ui),
            new ExplorationArtifactWriter(workspace),
            NullLogger<ExplorationCoordinator>.Instance,
            probe);
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

    private static readonly JsonSerializerOptions CamelCaseJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true
    };

    private static string WritePolicy(string workspaceRoot, ExplorationPolicyDocument policy)
    {
        var path = Path.Combine(workspaceRoot, "exploration", "policies", Guid.NewGuid().ToString("N") + ".json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, JsonSerializer.Serialize(policy, CamelCaseJson));
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
        string elementKey,
        long? targetHwnd = null,
        int? targetProcessId = null) =>
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
            Priority = priority,
            TargetHwnd = targetHwnd,
            TargetProcessId = targetProcessId
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
