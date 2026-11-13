using System.IO.Pipelines;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Mcp;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Moq;

namespace DesktopAutomationAgent.Tests;

public class Phase5McpTests
{
    private const string SecretToken = "secret-token-ABC12345";

    private static readonly string[] ExpectedTools =
    [
        "desktop_doctor",
        "desktop_observe",
        "desktop_explore_dry_run",
        "desktop_explore",
        "desktop_search_knowledge",
        "desktop_validate_plan",
        "desktop_compile_key",
        "desktop_run_plan",
        "desktop_run_key",
        "desktop_run_suite",
        "desktop_get_run_summary"
    ];

    [Fact]
    public void Parse_AcceptsMcpServer()
    {
        var parsed = CommandLine.Parse(["mcp-server", "--Workspace:Root=automation"]);
        Assert.Null(parsed.Error);
        Assert.Equal(AgentCommandKind.McpServer, parsed.Kind);
        Assert.Contains(parsed.ConfigurationArgs, a => a.Contains("Workspace:Root", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ToolAnnotations_ExploreDryRun_IsObservationCategory_NotReadOnly()
    {
        var method = typeof(DesktopMcpTools).GetMethod(nameof(DesktopMcpTools.DesktopExploreDryRun));
        Assert.NotNull(method);
        var attr = method!.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        Assert.Equal("desktop_explore_dry_run", attr!.Name);
        // May write workspace artifact stubs — must not advertise ReadOnly or Idempotent.
        Assert.False(attr.ReadOnly);
        Assert.False(attr.Destructive);
        Assert.False(attr.Idempotent);
        Assert.False(attr.OpenWorld);
        var description = method.GetCustomAttribute<System.ComponentModel.DescriptionAttribute>()!.Description;
        Assert.StartsWith("Observation:", description, StringComparison.Ordinal);
        Assert.DoesNotContain("Navigation:", description, StringComparison.Ordinal);
    }

    [Fact]
    public void ToolAnnotations_CompileKey_NotReadOnly_AndOpenWorld()
    {
        var attr = GetToolAttr(nameof(DesktopMcpTools.DesktopCompileKey));
        // Compile writes plan-cache artifacts — must not advertise ReadOnly.
        Assert.False(attr.ReadOnly);
        Assert.False(attr.Destructive);
        Assert.True(attr.Idempotent);
        Assert.True(attr.OpenWorld);
    }

    [Fact]
    public void ToolAnnotations_CompileAndFromJiraTools_AreOpenWorld()
    {
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopCompileKey)).OpenWorld);
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopRunKey)).OpenWorld);
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopRunSuite)).OpenWorld);
    }

    [Fact]
    public void ToolAnnotations_RunTools_DefaultDryRunTrue()
    {
        Assert.True(GetParameterDefaultBool(nameof(DesktopMcpTools.DesktopRunPlan), "dryRun"));
        Assert.True(GetParameterDefaultBool(nameof(DesktopMcpTools.DesktopRunKey), "dryRun"));
        Assert.True(GetParameterDefaultBool(nameof(DesktopMcpTools.DesktopRunSuite), "dryRun"));
    }

    [Fact]
    public void OpaqueRefStore_StaleRef_Throws()
    {
        var store = new McpOpaqueRefStore(TimeSpan.FromMilliseconds(1));
        var snapshot = new SemanticSnapshot
        {
            SnapshotId = "snap-1",
            Nodes =
            [
                new SemanticNode
                {
                    Ref = "r1",
                    ControlType = "Button",
                    Name = "Ok"
                }
            ]
        };
        store.ReplaceSnapshot(snapshot);
        Thread.Sleep(5);
        var ex = Assert.Throws<McpToolException>(() => store.Resolve("snap-1", "r1"));
        Assert.Equal(McpErrorCode.StaleRef, ex.ErrorCode);

        store.ReplaceSnapshot(snapshot);
        ex = Assert.Throws<McpToolException>(() => store.Resolve("other", "r1"));
        Assert.Equal(McpErrorCode.StaleRef, ex.ErrorCode);
    }

    [Fact]
    public void ResponseBudget_TruncatesLargePayload_ToValidJson()
    {
        var huge = new
        {
            success = true,
            token = SecretToken,
            nodes = Enumerable.Range(0, 5000).Select(i => new { id = i, name = new string('x', 80) }).ToArray()
        };
        var json = McpResponseBudget.SerializeBounded(huge, maxChars: 8_000);
        Assert.True(json.Length <= 8_000);
        Assert.Contains("truncated", json, StringComparison.OrdinalIgnoreCase);
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.ValueKind is JsonValueKind.Object);
    }

    [Fact]
    public void ResponseBudget_HugeSingleElement_ProducesValidJsonOrBudgetError()
    {
        var payload = new
        {
            success = true,
            blob = new string('A', 50_000) + "\\\"escape\\\"\n\t" + SecretToken
        };

        try
        {
            var json = McpResponseBudget.SerializeBounded(payload, maxChars: 4_000);
            Assert.True(json.Length <= 4_000);
            using var doc = JsonDocument.Parse(json);
            Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
        }
        catch (McpToolException ex)
        {
            Assert.Equal(McpErrorCode.Truncated, ex.ErrorCode);
        }
    }

    [Fact]
    public void ResponseBudget_TinyBudget_YieldsValidJsonOrTruncatedError()
    {
        var payload = new { success = true, data = Enumerable.Range(0, 100).Select(i => new string('z', 200)).ToArray() };
        try
        {
            var json = McpResponseBudget.SerializeBounded(payload, maxChars: 256);
            Assert.True(json.Length <= 256);
            using var doc = JsonDocument.Parse(json);
            Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
        }
        catch (McpToolException ex)
        {
            Assert.Equal(McpErrorCode.Truncated, ex.ErrorCode);
        }
    }

    [Fact]
    public void PathGuard_RejectsTraversal()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var ex = Assert.Throws<McpToolException>(() =>
            McpPathGuard.RequireWorkspacePath(workspace, "../outside.json", "planPath"));
        Assert.Equal(McpErrorCode.PathTraversal, ex.ErrorCode);
    }

    [Fact]
    public void PathGuard_RejectsDirectorySymlinkEscape()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), $"{{\"token\":\"{SecretToken}\"}}");
        var linkPath = Path.Combine(workspace.RootPath, "plans-link");

        try
        {
            if (!TryCreateSymbolicLink(() => Directory.CreateSymbolicLink(linkPath, outside), linkPath))
                return;
            var target = Path.Combine(linkPath, "secret.json");
            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(target, workspace.RootPath));
            Assert.Contains("outside", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(linkPath);
            TryDelete(outside);
        }
    }

    [Fact]
    public void PathGuard_RejectsChainedFileSymlinkEscape()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-chain-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        var outsideFile = Path.Combine(outside, "secret.json");
        File.WriteAllText(outsideFile, $"{{\"token\":\"{SecretToken}\"}}");
        var mid = Path.Combine(workspace.RootPath, "plans", "mid.json");
        var entry = Path.Combine(workspace.RootPath, "plans", "entry.json");

        try
        {
            if (!TryCreateSymbolicLink(() => File.CreateSymbolicLink(mid, outsideFile), mid))
                return;
            if (!TryCreateSymbolicLink(() => File.CreateSymbolicLink(entry, mid), entry))
                return;
            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(entry, workspace.RootPath));
            Assert.Contains("outside", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(entry);
            TryDelete(mid);
            TryDelete(outside);
        }
    }

    [Fact]
    public void PathGuard_RejectsWindowsJunctionEscape()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-junc-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), "{}");
        var junction = Path.Combine(workspace.RootPath, "junc-plans");

        try
        {
            RequireWindowsJunction(junction, outside);
            var target = Path.Combine(junction, "secret.json");
            Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(target, workspace.RootPath));
        }
        finally
        {
            TryDelete(junction);
            TryDelete(outside);
        }
    }

    [Fact]
    public void PathGuard_RejectsChainedDirectoryJunctionEscape()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-chain-dir-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), $"{{\"token\":\"{SecretToken}\"}}");

        var inner = Path.Combine(workspace.RootPath, "inner-link");
        var mid = Path.Combine(workspace.RootPath, "mid-link");

        try
        {
            // mid -> inner (inside) -> outside (escape). Immediate mid target looks safe;
            // ResolveLinkTarget(true) must still reject the chain.
            RequireWindowsJunction(inner, outside);
            RequireWindowsJunction(mid, inner);
            var target = Path.Combine(mid, "secret.json");
            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(target, workspace.RootPath));
            Assert.Contains("outside", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(mid);
            TryDelete(inner);
            TryDelete(outside);
        }
    }

    [Fact]
    public void PathGuard_RejectsChainedDirectorySymlinkEscape()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-chain-sl-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), $"{{\"token\":\"{SecretToken}\"}}");

        var inner = Path.Combine(workspace.RootPath, "inner-sl");
        var mid = Path.Combine(workspace.RootPath, "mid-sl");

        try
        {
            if (!TryCreateSymbolicLink(() => Directory.CreateSymbolicLink(inner, outside), inner))
                return;
            if (!TryCreateSymbolicLink(() => Directory.CreateSymbolicLink(mid, inner), mid))
                return;

            var target = Path.Combine(mid, "secret.json");
            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(target, workspace.RootPath));
            Assert.Contains("outside", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(mid);
            TryDelete(inner);
            TryDelete(outside);
        }
    }

    [Fact]
    public void PathGuard_RejectsSymlinkCycle_FailClosed()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var a = Path.Combine(workspace.RootPath, "plans", "cycle-a.json");
        var b = Path.Combine(workspace.RootPath, "plans", "cycle-b.json");
        Directory.CreateDirectory(Path.GetDirectoryName(a)!);

        try
        {
            // Create a temporary real file so the first link has a target, then retarget into a cycle.
            var temp = Path.Combine(workspace.RootPath, "plans", "cycle-temp.json");
            File.WriteAllText(temp, "{}");
            if (!TryCreateSymbolicLink(() => File.CreateSymbolicLink(a, temp), a))
                return;
            if (!TryCreateSymbolicLink(() => File.CreateSymbolicLink(b, a), b))
                return;

            File.Delete(a);
            if (!TryCreateSymbolicLink(() => File.CreateSymbolicLink(a, b), a))
                return;

            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(a, workspace.RootPath));
            Assert.True(
                thrown.Message.Contains("cycle", StringComparison.OrdinalIgnoreCase)
                || thrown.Message.Contains("cannot be validated", StringComparison.OrdinalIgnoreCase)
                || thrown.Message.Contains("outside", StringComparison.OrdinalIgnoreCase),
                thrown.Message);
        }
        finally
        {
            TryDelete(a);
            TryDelete(b);
            TryDelete(Path.Combine(workspace.RootPath, "plans", "cycle-temp.json"));
        }
    }

    [Fact]
    public void PathGuard_RejectsUnresolvableReparsePoint_FailClosed()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        // Junction to a missing target: ResolveLinkTarget should fail closed.
        var missingTarget = Path.Combine(Path.GetTempPath(), "da-mcp-missing-" + Guid.NewGuid().ToString("N"));
        var junction = Path.Combine(workspace.RootPath, "broken-junc");

        try
        {
            // Create target, make junction, then delete target so resolution fails / is dangling.
            Directory.CreateDirectory(missingTarget);
            RequireWindowsJunction(junction, missingTarget);
            Directory.Delete(missingTarget, recursive: true);

            var probe = Path.Combine(junction, "nope.json");
            var thrown = Assert.ThrowsAny<Exception>(() =>
                McpPathGuard.EnsureNotSymlinkEscape(probe, workspace.RootPath));
            Assert.True(
                thrown.Message.Contains("cannot be validated", StringComparison.OrdinalIgnoreCase)
                || thrown.Message.Contains("outside", StringComparison.OrdinalIgnoreCase)
                || thrown.Message.Contains("reparse", StringComparison.OrdinalIgnoreCase),
                thrown.Message);
        }
        finally
        {
            TryDelete(junction);
            TryDelete(missingTarget);
        }
    }

    [Fact]
    public void Facade_ValidatePlan_DelegatesAndKeepsValidation()
    {
        var options = TestSupport.CreateOptions(bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "mcp.plan.json", TestSupport.MinimalPlanJson());

        var facade = BuildFacade(options, workspace, ReadyHandler());
        var payload = facade.ValidatePlan(planPath);
        var json = JsonSerializer.Serialize(payload);
        Assert.Contains("success", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(SecretToken, json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Facade_RunPlan_RealRun_KeepsReport_NoTokenLeak()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33261",
            bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "mcp-run.plan.json", TestSupport.MinimalPlanJson());

        var facade = BuildFacade(options, workspace, ReadyHandler());
        var payload = await facade.RunPlanAsync(planPath, dryRun: false, CancellationToken.None);
        var json = McpResponseBudget.SerializeBounded(payload);
        Assert.Contains("runId", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(SecretToken, json, StringComparison.Ordinal);
        Assert.DoesNotContain("Authorization", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Facade_Observe_MapsDriverUnavailable()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33262",
            bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var handler = new FakeHttpMessageHandler(_ =>
            throw new HttpRequestException("connection refused bearer " + SecretToken));
        var facade = BuildFacade(options, workspace, handler);
        var ex = await Assert.ThrowsAsync<McpToolException>(() =>
            facade.ObserveAsync(null, null, 2, 10, 20, false, false, false, true, true, null, CancellationToken.None));
        Assert.True(
            ex.ErrorCode is McpErrorCode.DriverUnavailable or McpErrorCode.Internal or McpErrorCode.Validation,
            ex.ErrorCode);
        Assert.DoesNotContain(SecretToken, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Facade_Cancellation_Propagates()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33263",
            bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "cancel.plan.json", TestSupport.MinimalPlanJson());
        var facade = BuildFacade(options, workspace, ReadyHandler());
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => facade.RunPlanAsync(planPath, dryRun: false, cts.Token));
    }

    [Fact]
    public async Task Tools_DelegateOnce_ToFacade()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "observation" })
            .Verifiable();
        mock.Setup(f => f.ValidatePlan(It.IsAny<string>()))
            .Returns(new { success = true, category = "validation" })
            .Verifiable();
        mock.Setup(f => f.ObserveAsync(
                It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(),
                It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool?>(), It.IsAny<bool?>(),
                It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "observation", nodeCount = 0 })
            .Verifiable();

        var tools = new DesktopMcpTools(mock.Object);
        Assert.True((await tools.DesktopDoctor(CancellationToken.None)).IsError != true);
        Assert.True((await tools.DesktopValidatePlan("plans/x.json", CancellationToken.None)).IsError != true);
        Assert.True((await tools.DesktopObserve(null, null, null, null, null, false, false, false, null, null, null, CancellationToken.None)).IsError != true);
        mock.Verify(f => f.DoctorAsync(It.IsAny<CancellationToken>()), Times.Once);
        mock.Verify(f => f.ValidatePlan(It.IsAny<string>()), Times.Once);
        mock.Verify(f => f.ObserveAsync(
            It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(),
            It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool?>(), It.IsAny<bool?>(),
            It.IsAny<string?>(), It.IsAny<CancellationToken>()), Times.Once);
        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Tools_DoNotWriteToStdout()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, note = "Bearer " + SecretToken });

        var original = Console.Out;
        var captured = new StringWriter();
        Console.SetOut(captured);
        try
        {
            var tools = new DesktopMcpTools(mock.Object);
            var result = await tools.DesktopDoctor(CancellationToken.None);
            Assert.Equal(string.Empty, captured.ToString());
            AssertRawWireHasNoSecret(result);
        }
        finally
        {
            Console.SetOut(original);
        }
    }

    [Fact]
    public async Task Tools_RedactSecrets_InSuccessAndException_WirePayload()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.SetupSequence(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new
            {
                success = true,
                authorizationHeader = "Bearer " + SecretToken,
                token = SecretToken,
                detail = "Authorization: Bearer " + SecretToken
            })
            .ThrowsAsync(new InvalidOperationException("driver failed Authorization: Bearer " + SecretToken));

        var tools = new DesktopMcpTools(mock.Object);
        var ok = await tools.DesktopDoctor(CancellationToken.None);
        AssertRawWireHasNoSecret(ok);
        AssertWireIsValidJsonUnderBudget(ok);

        var err = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(err.IsError == true);
        AssertRawWireHasNoSecret(err);
        AssertWireIsValidJsonUnderBudget(err);
    }

    [Fact]
    public async Task Tools_LargeException_IsBudgetedAndRedacted()
    {
        var huge = new string('X', 60_000) + " Authorization: Bearer " + SecretToken;
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException(huge));

        var tools = new DesktopMcpTools(mock.Object);
        var err = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(err.IsError == true);
        AssertRawWireHasNoSecret(err);
        AssertWireIsValidJsonUnderBudget(err);
        Assert.True(GetRawText(err).Length <= McpResponseBudget.DefaultMaxChars);
    }

    [Fact]
    public async Task Tools_NearLimitSecretPayload_RedactedUnderBudget()
    {
        // Build a success payload that sits near the default budget with an embedded secret.
        var pad = new string('p', McpResponseBudget.DefaultMaxChars - 400);
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new
            {
                success = true,
                pad,
                token = SecretToken,
                authorizationHeader = "Bearer " + SecretToken
            });

        var tools = new DesktopMcpTools(mock.Object);
        var ok = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(ok.IsError != true, GetRawText(ok));
        AssertRawWireHasNoSecret(ok);
        AssertWireIsValidJsonUnderBudget(ok);
        Assert.True(GetRawText(ok).Length <= McpResponseBudget.DefaultMaxChars);
    }

    [Fact]
    public async Task Tools_RepeatedSecrets_RedactedThenBudgeted()
    {
        var repeats = Enumerable.Range(0, 200).Select(i => new
        {
            i,
            token = SecretToken,
            authorizationHeader = "Bearer " + SecretToken,
            note = $"Authorization: Bearer {SecretToken}"
        }).ToArray();

        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, repeats });

        var tools = new DesktopMcpTools(mock.Object);
        var ok = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(ok.IsError != true, GetRawText(ok));
        AssertRawWireHasNoSecret(ok);
        AssertWireIsValidJsonUnderBudget(ok);
        Assert.DoesNotContain(SecretToken, GetRawText(ok), StringComparison.Ordinal);
    }

    [Fact]
    public void ResponseBudget_RedactThenBudget_RemovesSecretBeforeTruncation()
    {
        var payload = new
        {
            success = true,
            token = SecretToken,
            authorizationHeader = "Bearer " + SecretToken,
            // Space-separated so each Bearer token is a complete redaction unit before budgeting.
            blob = string.Join(
                " ",
                Enumerable.Repeat($"Authorization: Bearer {SecretToken}", 5_000))
        };

        var json = McpResponseBudget.SerializeRedactedBounded(payload, maxChars: 2_000);
        Assert.True(json.Length <= 2_000);
        Assert.DoesNotContain(SecretToken, json, StringComparison.Ordinal);
        using var doc = JsonDocument.Parse(json);
        Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
    }

    [Fact]
    public async Task Protocol_InitListsExpectedToolsAndSchemas()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.ValidatePlan(It.IsAny<string>()))
            .Returns(new { success = true, category = "validation", planId = "SAMPLE-1", stepCount = 1 });
        mock.Setup(f => f.RunPlanAsync(It.IsAny<string>(), false, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new
            {
                success = true,
                category = "execution",
                runId = "run-fixture",
                status = "passed",
                exitCode = 0,
                dryRun = false,
                stepCount = 1,
                token = SecretToken
            });
        mock.Setup(f => f.ObserveAsync(
                It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(),
                It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool?>(), It.IsAny<bool?>(),
                It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "observation", nodeCount = 2, snapshotId = "snap" });

        await using var session = await McpTestSession.CreateAsync(mock.Object);
        var tools = await session.Client.ListToolsAsync();
        var names = tools.Select(t => t.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();
        Assert.Equal(ExpectedTools.OrderBy(n => n, StringComparer.Ordinal).ToArray(), names);

        foreach (var tool in tools)
        {
            Assert.False(string.IsNullOrWhiteSpace(tool.Description));
            Assert.True(tool.JsonSchema.ValueKind is JsonValueKind.Object);
            Assert.True(tool.JsonSchema.TryGetProperty("type", out var type));
            Assert.Equal("object", type.GetString());
            Assert.True(tool.JsonSchema.TryGetProperty("properties", out _));
        }

        var compile = tools.First(t => t.Name == "desktop_compile_key");
        Assert.Contains("fromJira", compile.Description, StringComparison.OrdinalIgnoreCase);

        var exploreDry = tools.First(t => t.Name == "desktop_explore_dry_run");
        Assert.StartsWith("Observation:", exploreDry.Description, StringComparison.Ordinal);

        var validate = tools.First(t => t.Name == "desktop_validate_plan");
        var validateResult = await validate.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "plans/example.plan.json"
        });
        AssertRawWireHasNoSecret(validateResult);
        Assert.Contains("SAMPLE-1", GetRawText(validateResult), StringComparison.Ordinal);

        var observe = tools.First(t => t.Name == "desktop_observe");
        var observeResult = await observe.CallAsync(new Dictionary<string, object?>
        {
            ["includeOffscreen"] = false,
            ["includeBounds"] = false,
            ["includeHelpText"] = false,
            ["maxNodes"] = 20,
            ["maxDepth"] = 3
        });
        Assert.True(observeResult.IsError != true, GetRawText(observeResult));

        var run = tools.First(t => t.Name == "desktop_run_plan");
        Assert.True(run.JsonSchema.TryGetProperty("properties", out var runProps));
        Assert.True(runProps.TryGetProperty("dryRun", out _));

        var runResult = await run.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "plans/example.plan.json",
            ["dryRun"] = false
        });
        AssertRawWireHasNoSecret(runResult);
        Assert.Contains("run-fixture", GetRawText(runResult), StringComparison.Ordinal);
        mock.Verify(f => f.ValidatePlan(It.IsAny<string>()), Times.Once);
        mock.Verify(f => f.RunPlanAsync(It.IsAny<string>(), false, It.IsAny<CancellationToken>()), Times.Once);
        mock.Verify(f => f.ObserveAsync(
            It.IsAny<string?>(), It.IsAny<string?>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<int?>(),
            It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool>(), It.IsAny<bool?>(), It.IsAny<bool?>(),
            It.IsAny<string?>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Protocol_RunPlan_OmittingDryRun_DefaultsToDryRunTrue()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.RunPlanAsync(It.IsAny<string>(), true, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "validation", dryRun = true, runId = "dry-default" })
            .Verifiable();

        await using var session = await McpTestSession.CreateAsync(mock.Object);
        var tools = await session.Client.ListToolsAsync();
        var run = tools.First(t => t.Name == "desktop_run_plan");

        // Omission must not live-run.
        var result = await run.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "plans/example.plan.json"
        });
        Assert.True(result.IsError != true, GetRawText(result));
        Assert.Contains("dry-default", GetRawText(result), StringComparison.Ordinal);
        mock.Verify(f => f.RunPlanAsync(It.IsAny<string>(), true, It.IsAny<CancellationToken>()), Times.Once);
        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Protocol_ChainedDirectoryJunction_ReturnsPathTraversal()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var options = TestSupport.CreateOptions(bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var outside = Path.Combine(Path.GetTempPath(), "da-mcp-proto-chain-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), $"{{\"token\":\"{SecretToken}\"}}");
        var inner = Path.Combine(workspace.RootPath, "proto-inner");
        var mid = Path.Combine(workspace.RootPath, "proto-mid");

        try
        {
            RequireWindowsJunction(inner, outside);
            RequireWindowsJunction(mid, inner);
            // Place a "plan" path that walks through the chained junctions.
            File.WriteAllText(Path.Combine(outside, "evil.plan.json"), TestSupport.MinimalPlanJson());

            var facade = BuildFacade(options, workspace, ReadyHandler());
            await using var session = await McpTestSession.CreateAsync(facade);
            var tools = await session.Client.ListToolsAsync();
            var validate = tools.First(t => t.Name == "desktop_validate_plan");
            var result = await validate.CallAsync(new Dictionary<string, object?>
            {
                ["planPath"] = "proto-mid/evil.plan.json"
            });
            Assert.True(result.IsError == true, GetRawText(result));
            Assert.Contains("path_traversal", GetRawText(result), StringComparison.OrdinalIgnoreCase);
            AssertRawWireHasNoSecret(result);
        }
        finally
        {
            TryDelete(mid);
            TryDelete(inner);
            TryDelete(outside);
        }
    }

    [Fact]
    public async Task Tools_InvalidInputMatrix_Validation_NoFacadeCalls()
    {
        // Strict mock with zero setups: real tool-layer validation must reject before any facade call.
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        var tools = new DesktopMcpTools(mock.Object);

        async Task AssertValidation(CallToolResult result)
        {
            Assert.True(result.IsError == true, GetRawText(result));
            AssertRawWireHasNoSecret(result);
            AssertWireIsValidJsonUnderBudget(result);
            using var doc = JsonDocument.Parse(GetRawText(result));
            Assert.Equal(McpErrorCode.Validation, doc.RootElement.GetProperty("errorCode").GetString());
        }

        await AssertValidation(await tools.DesktopObserve(
            view: "not-a-view", root: null, maxDepth: null, maxChildren: null, maxNodes: null,
            false, false, false, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopObserve(
            view: null, root: null, maxDepth: 99, maxChildren: null, maxNodes: null,
            false, false, false, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopExploreDryRun("  ", CancellationToken.None));
        await AssertValidation(await tools.DesktopExplore("", CancellationToken.None));
        await AssertValidation(await tools.DesktopSearchKnowledge("", "q", null, CancellationToken.None));
        await AssertValidation(await tools.DesktopSearchKnowledge("k.json", "q", limit: 0, CancellationToken.None));
        await AssertValidation(await tools.DesktopValidatePlan("", CancellationToken.None));
        await AssertValidation(await tools.DesktopCompileKey("", file: null, fromJira: false, null, null, null, null, false, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunPlan("", dryRun: false, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunKey("", dryRun: false, fromJira: false, null, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunSuite("", dryRun: false, fromJira: false, false, null, null, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopGetRunSummary(" ", CancellationToken.None));

        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Tools_MaxLength_EnforcedBeforeFacade_VerifyNoOtherCalls()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        var tools = new DesktopMcpTools(mock.Object);

        async Task AssertValidation(CallToolResult result)
        {
            Assert.True(result.IsError == true, GetRawText(result));
            AssertRawWireHasNoSecret(result);
            AssertWireIsValidJsonUnderBudget(result);
            using var doc = JsonDocument.Parse(GetRawText(result));
            Assert.Equal(McpErrorCode.Validation, doc.RootElement.GetProperty("errorCode").GetString());
            Assert.Contains("maximum length", doc.RootElement.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
        }

        var key65 = new string('K', 65);
        var runId129 = new string('R', 129);
        var query257 = new string('Q', 257);
        var path513 = new string('P', 513);

        await AssertValidation(await tools.DesktopCompileKey(
            key65, file: null, fromJira: false, null, null, null, null, false, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunKey(
            key65, dryRun: true, fromJira: false, null, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopGetRunSummary(runId129, CancellationToken.None));
        await AssertValidation(await tools.DesktopSearchKnowledge(
            "knowledge/app.json", query257, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopValidatePlan(path513, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunPlan(path513, dryRun: true, CancellationToken.None));
        await AssertValidation(await tools.DesktopExploreDryRun(path513, CancellationToken.None));
        await AssertValidation(await tools.DesktopExplore(path513, CancellationToken.None));
        await AssertValidation(await tools.DesktopSearchKnowledge(path513, "ok", null, CancellationToken.None));
        await AssertValidation(await tools.DesktopRunSuite(path513, dryRun: true, fromJira: false, false, null, null, null, null, null, CancellationToken.None));
        await AssertValidation(await tools.DesktopCompileKey(
            "SAMPLE-1", file: path513, fromJira: false, null, null, null, null, false, CancellationToken.None));
        await AssertValidation(await tools.DesktopObserve(
            null, null, null, null, null, false, false, false, null, null, path513, CancellationToken.None));

        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Protocol_InvalidInput_RequiredOmissions_NoFacadeCalls()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        await using var session = await McpTestSession.CreateAsync(mock.Object);
        var tools = (await session.Client.ListToolsAsync()).ToDictionary(t => t.Name, StringComparer.Ordinal);

        async Task AssertRejected(string toolName, Dictionary<string, object?> args)
        {
            var result = await tools[toolName].CallAsync(args);
            Assert.True(result.IsError == true, $"{toolName}: {GetRawText(result)}");
            AssertRawWireHasNoSecret(result);
            var text = GetRawText(result).TrimStart();
            if (text.StartsWith("{", StringComparison.Ordinal))
            {
                AssertWireIsValidJsonUnderBudget(result);
                using var doc = JsonDocument.Parse(text);
                if (doc.RootElement.TryGetProperty("errorCode", out var code))
                    Assert.Equal(McpErrorCode.Validation, code.GetString());
            }
        }

        await AssertRejected("desktop_explore_dry_run", new Dictionary<string, object?>());
        await AssertRejected("desktop_explore", new Dictionary<string, object?>());
        await AssertRejected("desktop_search_knowledge", new Dictionary<string, object?> { ["query"] = "x" });
        await AssertRejected("desktop_validate_plan", new Dictionary<string, object?>());
        await AssertRejected("desktop_compile_key", new Dictionary<string, object?>());
        await AssertRejected("desktop_run_plan", new Dictionary<string, object?>());
        await AssertRejected("desktop_run_key", new Dictionary<string, object?>());
        await AssertRejected("desktop_run_suite", new Dictionary<string, object?>());
        await AssertRejected("desktop_get_run_summary", new Dictionary<string, object?>());
        await AssertRejected("desktop_run_plan", new Dictionary<string, object?> { ["dryRun"] = false });
        await AssertRejected("desktop_observe", new Dictionary<string, object?> { ["maxDepth"] = 99 });

        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Protocol_OptionalArgs_MayBeOmitted_AndRequiredArraysMatch()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.SearchKnowledge("knowledge/app.json", "login", null))
            .Returns(new { success = true, category = "observation", hits = Array.Empty<object>() });
        mock.Setup(f => f.CompileKey("SAMPLE-1", null, false, null, null, null, null, false))
            .Returns(new { success = true, category = "validation", key = "SAMPLE-1" });
        mock.Setup(f => f.ObserveAsync(
                null, null, null, null, null,
                false, false, false, null, null, null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "observation", nodeCount = 0 });

        await using var session = await McpTestSession.CreateAsync(mock.Object);
        var tools = (await session.Client.ListToolsAsync()).ToDictionary(t => t.Name, StringComparer.Ordinal);

        static string[] RequiredOf(JsonElement schema)
        {
            if (!schema.TryGetProperty("required", out var required) || required.ValueKind != JsonValueKind.Array)
                return [];
            return required.EnumerateArray().Select(e => e.GetString()!).Where(s => s is not null).ToArray()!;
        }

        Assert.Equal(["policyPath"], RequiredOf(tools["desktop_explore_dry_run"].JsonSchema));
        Assert.Equal(["policyPath"], RequiredOf(tools["desktop_explore"].JsonSchema));
        Assert.Contains("knowledgePath", RequiredOf(tools["desktop_search_knowledge"].JsonSchema));
        Assert.Contains("query", RequiredOf(tools["desktop_search_knowledge"].JsonSchema));
        Assert.DoesNotContain("limit", RequiredOf(tools["desktop_search_knowledge"].JsonSchema));
        Assert.Equal(["planPath"], RequiredOf(tools["desktop_validate_plan"].JsonSchema));
        Assert.Contains("key", RequiredOf(tools["desktop_compile_key"].JsonSchema));
        Assert.DoesNotContain("file", RequiredOf(tools["desktop_compile_key"].JsonSchema));
        Assert.Equal(["planPath"], RequiredOf(tools["desktop_run_plan"].JsonSchema));
        Assert.DoesNotContain("dryRun", RequiredOf(tools["desktop_run_plan"].JsonSchema));
        Assert.Contains("key", RequiredOf(tools["desktop_run_key"].JsonSchema));
        Assert.DoesNotContain("dryRun", RequiredOf(tools["desktop_run_key"].JsonSchema));
        Assert.Equal(["suitePath"], RequiredOf(tools["desktop_run_suite"].JsonSchema));
        Assert.Equal(["runId"], RequiredOf(tools["desktop_get_run_summary"].JsonSchema));
        Assert.Empty(RequiredOf(tools["desktop_doctor"].JsonSchema));
        Assert.Empty(RequiredOf(tools["desktop_observe"].JsonSchema));

        // Minimal omit: only required args; optional file/limit/view omitted.
        var search = await tools["desktop_search_knowledge"].CallAsync(new Dictionary<string, object?>
        {
            ["knowledgePath"] = "knowledge/app.json",
            ["query"] = "login"
        });
        Assert.True(search.IsError != true, GetRawText(search));

        var compile = await tools["desktop_compile_key"].CallAsync(new Dictionary<string, object?>
        {
            ["key"] = "SAMPLE-1"
        });
        Assert.True(compile.IsError != true, GetRawText(compile));

        var observe = await tools["desktop_observe"].CallAsync(new Dictionary<string, object?>());
        Assert.True(observe.IsError != true, GetRawText(observe));

        mock.VerifyAll();
        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Protocol_PathTraversal_ReturnsStructuredError()
    {
        var options = TestSupport.CreateOptions(bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var facade = BuildFacade(options, workspace, ReadyHandler());

        await using var session = await McpTestSession.CreateAsync(facade);
        var tools = await session.Client.ListToolsAsync();
        var validate = tools.First(t => t.Name == "desktop_validate_plan");
        var result = await validate.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "../Secrets/token.txt"
        });
        Assert.True(result.IsError == true);
        Assert.Contains("path_traversal", GetRawText(result), StringComparison.OrdinalIgnoreCase);
        AssertRawWireHasNoSecret(result);
    }

    [Fact]
    public async Task Protocol_PolicyDenial_ReturnsDenied()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33264",
            bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var policyPath = Path.Combine(workspace.RootPath, "exploration", "policies", "deny-all.json");
        Directory.CreateDirectory(Path.GetDirectoryName(policyPath)!);
        await File.WriteAllTextAsync(policyPath, """
            {
              "policySchemaVersion": 1,
              "policyId": "deny-all",
              "mode": "navigate",
              "application": { "processName": "MissingProcessThatDoesNotExist" },
              "scope": { "allowedWindowTitles": [], "deniedWindowTitles": ["*"] },
              "allowedActionCategories": [],
              "allowRules": [],
              "denyRules": [ { "name": "*", "action": "*" } ],
              "budgets": { "maxActions": 1, "maxActions": 0, "maxDepth": 0 }
            }
            """);

        var facade = BuildFacade(options, workspace, ReadyHandler());
        await using var session = await McpTestSession.CreateAsync(facade);
        var tools = await session.Client.ListToolsAsync();
        var explore = tools.First(t => t.Name == "desktop_explore_dry_run");
        var result = await explore.CallAsync(new Dictionary<string, object?>
        {
            ["policyPath"] = "exploration/policies/deny-all.json"
        });
        AssertRawWireHasNoSecret(result);
        Assert.True(result.IsError == true
                    || GetRawText(result).Contains("denied", StringComparison.OrdinalIgnoreCase)
                    || GetRawText(result).Contains("policy", StringComparison.OrdinalIgnoreCase)
                    || GetRawText(result).Contains("validation", StringComparison.OrdinalIgnoreCase)
                    || GetRawText(result).Contains("driver", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task ExitGate_McpClient_ObserveAndRunDeterministicPlan()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33265",
            bearerToken: SecretToken);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "exit-gate.plan.json", TestSupport.MinimalPlanJson(planId: "exit.gate"));
        var relativePlan = Path.GetRelativePath(workspace.RootPath, planPath).Replace('\\', '/');

        var facade = BuildFacade(options, workspace, ObserveAndRunHandler());
        await using var session = await McpTestSession.CreateAsync(facade);
        var tools = await session.Client.ListToolsAsync();

        var observe = tools.First(t => t.Name == "desktop_observe");
        var observeResult = await observe.CallAsync(new Dictionary<string, object?>
        {
            ["includeOffscreen"] = false,
            ["includeBounds"] = false,
            ["includeHelpText"] = false,
            ["maxNodes"] = 20,
            ["maxDepth"] = 3
        });
        Assert.True(observeResult.IsError != true, GetRawText(observeResult));
        AssertRawWireHasNoSecret(observeResult);
        Assert.Contains("observation", GetRawText(observeResult), StringComparison.OrdinalIgnoreCase);

        var validate = tools.First(t => t.Name == "desktop_validate_plan");
        var validateResult = await validate.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = relativePlan
        });
        Assert.True(validateResult.IsError != true, GetRawText(validateResult));
        Assert.Contains("exit.gate", GetRawText(validateResult), StringComparison.Ordinal);

        var run = tools.First(t => t.Name == "desktop_run_plan");
        var runResult = await run.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = relativePlan,
            ["dryRun"] = false
        });
        Assert.True(runResult.IsError != true, GetRawText(runResult));
        var runText = GetRawText(runResult);
        Assert.Contains("runId", runText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\"dryRun\":true", runText.Replace(" ", ""), StringComparison.OrdinalIgnoreCase);
        AssertRawWireHasNoSecret(runResult);
    }

    private static McpServerToolAttribute GetToolAttr(string methodName)
    {
        var method = typeof(DesktopMcpTools).GetMethod(methodName);
        Assert.NotNull(method);
        var attr = method!.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        return attr!;
    }

    private static bool GetParameterDefaultBool(string methodName, string parameterName)
    {
        var method = typeof(DesktopMcpTools).GetMethod(methodName);
        Assert.NotNull(method);
        var parameter = method!.GetParameters().First(p => p.Name == parameterName);
        Assert.True(parameter.HasDefaultValue, $"{methodName}.{parameterName} must have a default");
        Assert.IsType<bool>(parameter.DefaultValue);
        return (bool)parameter.DefaultValue!;
    }

    private static string GetRawText(CallToolResult result) =>
        string.Join("", result.Content.OfType<TextContentBlock>().Select(c => c.Text));

    private static void AssertRawWireHasNoSecret(CallToolResult result)
    {
        var text = GetRawText(result);
        Assert.DoesNotContain(SecretToken, text, StringComparison.Ordinal);
        Assert.DoesNotContain("Bearer " + SecretToken, text, StringComparison.Ordinal);

        if (result.StructuredContent is { } structured
            && structured.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null)
        {
            var structuredJson = structured.GetRawText();
            Assert.DoesNotContain(SecretToken, structuredJson, StringComparison.Ordinal);
            Assert.DoesNotContain("Bearer " + SecretToken, structuredJson, StringComparison.Ordinal);
        }
    }

    private static void AssertWireIsValidJsonUnderBudget(CallToolResult result)
    {
        var text = GetRawText(result);
        Assert.True(text.Length <= McpResponseBudget.DefaultMaxChars, $"wire length {text.Length}");
        using var doc = JsonDocument.Parse(text);
        Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);

        if (result.StructuredContent is { } structured
            && structured.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null)
        {
            Assert.Equal(JsonValueKind.Object, structured.ValueKind);
            Assert.True(structured.GetRawText().Length <= McpResponseBudget.DefaultMaxChars);
        }
    }

    private static FakeHttpMessageHandler ReadyHandler() =>
        new(req =>
        {
            if (req.Method == HttpMethod.Post)
                return FakeHttpMessageHandler.Json(new { success = true, value = true });

            if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    status = 0,
                    value = new
                    {
                        ready = true,
                        message = "ok",
                        build = new { version = "1.0.105" }
                    }
                });
            }

            return FakeHttpMessageHandler.Json(new
            {
                success = true,
                value = CatalogFixtures.Phase2Catalog()
            });
        });

    private static FakeHttpMessageHandler ObserveAndRunHandler() =>
        new(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                var body = req.Content is null
                    ? string.Empty
                    : await req.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (body.Contains("dumpuia", StringComparison.OrdinalIgnoreCase))
                {
                    return FakeHttpMessageHandler.Json(new
                    {
                        success = true,
                        value = new
                        {
                            operation = "dumpuia",
                            success = true,
                            semanticDiagnosticVersion = 1,
                            truncated = false,
                            rootDiagnostics = new
                            {
                                activeWindowHwnd = 42,
                                processId = 77,
                                processName = "SampleApp",
                                windowsSessionId = 1,
                                activeWindow = new
                                {
                                    name = "Sample",
                                    className = "Wnd",
                                    processId = 77,
                                    nativeWindowHandle = 42
                                }
                            },
                            nodes = new object[]
                            {
                                new
                                {
                                    nodeIndex = 0,
                                    parentNodeIndex = (int?)null,
                                    ordinal = 0,
                                    depth = 0,
                                    controlType = "Window",
                                    name = "Sample",
                                    automationId = "main",
                                    className = "Wnd",
                                    isEnabled = true,
                                    isOffscreen = false,
                                    supportedPatterns = new[] { "Window" },
                                    capabilities = new[] { "window" }
                                },
                                new
                                {
                                    nodeIndex = 1,
                                    parentNodeIndex = 0,
                                    ordinal = 0,
                                    depth = 1,
                                    controlType = "Button",
                                    name = "Go",
                                    automationId = "go",
                                    isEnabled = true,
                                    isOffscreen = false,
                                    supportedPatterns = new[] { "Invoke" },
                                    capabilities = new[] { "invoke" }
                                }
                            }
                        }
                    });
                }

                return FakeHttpMessageHandler.Json(new { success = true, value = true });
            }

            if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    status = 0,
                    value = new { ready = true, message = "ok", build = new { version = "1.0.105" } }
                });
            }

            return FakeHttpMessageHandler.Json(new { success = true, value = CatalogFixtures.Phase2Catalog() });
        });

    private static IDesktopMcpFacade BuildFacade(
        AgentOptions options,
        IWorkspaceManager workspace,
        FakeHttpMessageHandler handler)
    {
        var factory = TestSupport.CreateFactory(handler);
        var services = new ServiceCollection();
        services.AddLogging(b => b.ClearProviders());
        services.AddSingleton(Options.Create(options));
        services.AddSingleton(workspace);
        services.AddSingleton(factory);
        services.AddSingleton<IDriverConnectionResolver>(_ =>
            new DriverConnectionResolver(Options.Create(options), factory, NullLogger<DriverConnectionResolver>.Instance));
        services.AddSingleton<IDriverCatalogClient>(_ =>
            new DriverCatalogClient(Options.Create(options), factory, NullLogger<DriverCatalogClient>.Instance));
        services.AddSingleton<IDriverUiClient>(_ =>
            new DriverUiClient(Options.Create(options), factory, NullLogger<DriverUiClient>.Instance));
        services.AddSingleton<AssertionEvaluator>();
        services.AddSingleton<RunArtifactWriter>();
        TestSupport.AddRecoveryServices(services);
        services.AddSingleton<IDeterministicPlanRunner, DeterministicPlanRunner>();
        services.AddSingleton<Readiness.IAgentReadinessService, Readiness.AgentReadinessService>();
        services.AddSingleton<Plans.PlanManifestReader>(_ =>
            new Plans.PlanManifestReader(Options.Create(options), workspace));
        services.AddSingleton<ObjectRepository.ObjectRepositoryReader>(_ =>
            new ObjectRepository.ObjectRepositoryReader(Options.Create(options), workspace));
        services.AddSingleton<ObjectRepository.ObjectReferenceResolver>();
        services.AddSingleton<ObjectRepository.PlanObjectReferenceExpander>();
        services.AddSingleton<ObjectRepository.PlanObjectRepositoryIntegrator>();
        services.AddSingleton<ObjectRepository.ObjectArtifactWriter>();
        services.AddSingleton<ObjectRepository.ObjectCandidateGenerator>();
        services.AddSingleton<ObjectRepository.ObjectCaptureService>();
        services.AddSingleton<ObjectRepository.ObjectVerificationService>();
        services.AddSingleton<SemanticSnapshotNormalizer>();
        services.AddSingleton<ISemanticObservationService, SemanticObservationService>();
        services.AddSingleton<ExplorationPolicyReader>();
        services.AddSingleton<StateFingerprinter>();
        services.AddSingleton<CapabilityResolver>();
        services.AddSingleton<IProcessIdentityProbe, SystemProcessIdentityProbe>();
        services.AddSingleton<UiStabilizationService>();
        services.AddSingleton<WindowTransitionDetector>();
        services.AddSingleton<ExplorationArtifactWriter>();
        services.AddSingleton<IExplorationCoordinator, ExplorationCoordinator>();
        services.AddSingleton<Knowledge.KnowledgeService>();
        services.AddSingleton<Compilation.BddCompilationService>();
        services.AddSingleton<Suites.ISuiteManifestReader>(_ =>
            new Suites.SuiteManifestReader(Options.Create(options), workspace));
        services.AddSingleton<Jira.IJiraTestCaseSource, Jira.JiraRestTestCaseSource>();
        services.AddSingleton<Jira.JiraSuiteOrchestrator>();
        services.AddSingleton<McpOpaqueRefStore>();
        services.AddSingleton<IDesktopMcpFacade, DesktopMcpFacade>();
        services.AddHttpClient("driver-unauthenticated");
        services.AddHttpClient("driver-authenticated");
        services.AddHttpClient(Jira.JiraHttpClientNames.Default);
        return services.BuildServiceProvider().GetRequiredService<IDesktopMcpFacade>();
    }

    private static bool TryCreateSymbolicLink(Action create, string path)
    {
        try
        {
            create();
        }
        catch (Exception ex) when (ex is IOException or PlatformNotSupportedException or UnauthorizedAccessException)
        {
            // Developer machines without SeCreateSymbolicLinkPrivilege cannot create symlinks.
            // CI Windows images that grant the privilege still exercise these cases.
            if (ex.Message.Contains("privilege", StringComparison.OrdinalIgnoreCase)
                || ex is PlatformNotSupportedException)
            {
                return false;
            }

            Assert.Fail(
                $"Unable to create symbolic link '{path}' required for path-safety coverage. {ex.GetType().Name}: {ex.Message}");
        }

        if (!File.Exists(path) && !Directory.Exists(path))
            Assert.Fail($"Symbolic link '{path}' was not created.");
        return true;
    }

    private static void RequireWindowsJunction(string junctionPath, string targetPath)
    {
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c mklink /J \"{junctionPath}\" \"{targetPath}\"",
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        using var process = System.Diagnostics.Process.Start(psi);
        Assert.NotNull(process);
        process!.WaitForExit(5000);
        if (process.ExitCode != 0 || !Directory.Exists(junctionPath))
        {
            Assert.Fail(
                $"Unable to create junction '{junctionPath}' -> '{targetPath}'. exit={process.ExitCode} err={process.StandardError.ReadToEnd()}");
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
            else if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // best-effort cleanup
        }
    }

    private sealed class McpTestSession : IAsyncDisposable
    {
        private readonly IHost _host;

        private McpTestSession(IHost host, McpClient client)
        {
            _host = host;
            Client = client;
        }

        public McpClient Client { get; }

        public static async Task<McpTestSession> CreateAsync(IDesktopMcpFacade facade)
        {
            var clientToServer = new Pipe();
            var serverToClient = new Pipe();

            var builder = Host.CreateApplicationBuilder();
            builder.Logging.ClearProviders();
            builder.Services.AddSingleton(facade);
            builder.Services
                .AddMcpServer(o =>
                {
                    o.ServerInfo = new() { Name = McpServerHost.ServerName, Version = McpServerHost.ServerVersion };
                })
                .WithStreamServerTransport(clientToServer.Reader.AsStream(), serverToClient.Writer.AsStream())
                .WithTools<DesktopMcpTools>();

            var host = builder.Build();
            await host.StartAsync().ConfigureAwait(false);

            var client = await McpClient.CreateAsync(new StreamClientTransport(
                clientToServer.Writer.AsStream(),
                serverToClient.Reader.AsStream())).ConfigureAwait(false);

            return new McpTestSession(host, client);
        }

        public async ValueTask DisposeAsync()
        {
            await Client.DisposeAsync().ConfigureAwait(false);
            await _host.StopAsync().ConfigureAwait(false);
            if (_host is IAsyncDisposable asyncDisposable)
                await asyncDisposable.DisposeAsync().ConfigureAwait(false);
            else
                _host.Dispose();
        }
    }
}
