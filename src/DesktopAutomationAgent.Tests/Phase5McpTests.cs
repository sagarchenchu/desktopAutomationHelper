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
    public void ToolAnnotations_ExploreDryRun_IsObservationOnly()
    {
        var method = typeof(DesktopMcpTools).GetMethod(nameof(DesktopMcpTools.DesktopExploreDryRun));
        Assert.NotNull(method);
        var attr = method!.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        Assert.Equal("desktop_explore_dry_run", attr!.Name);
        Assert.True(attr.ReadOnly);
        Assert.False(attr.Destructive);
        Assert.True(attr.Idempotent);
        Assert.False(attr.OpenWorld);
        var description = method.GetCustomAttribute<System.ComponentModel.DescriptionAttribute>()!.Description;
        Assert.StartsWith("Observation:", description, StringComparison.Ordinal);
        Assert.DoesNotContain("Navigation:", description, StringComparison.Ordinal);
    }

    [Fact]
    public void ToolAnnotations_CompileAndFromJiraTools_AreOpenWorld()
    {
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopCompileKey)).OpenWorld);
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopRunKey)).OpenWorld);
        Assert.True(GetToolAttr(nameof(DesktopMcpTools.DesktopRunSuite)).OpenWorld);
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

        var err = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(err.IsError == true);
        AssertRawWireHasNoSecret(err);
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
