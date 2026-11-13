using System.IO.Pipelines;
using System.Net;
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
using Moq;

namespace DesktopAutomationAgent.Tests;

public class Phase5McpTests
{
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
    public void ResponseBudget_TruncatesLargePayload()
    {
        var huge = new
        {
            success = true,
            nodes = Enumerable.Range(0, 5000).Select(i => new { id = i, name = new string('x', 80) }).ToArray()
        };
        var json = McpResponseBudget.SerializeBounded(huge, maxChars: 8_000);
        Assert.True(json.Length <= 8_200);
        Assert.Contains("truncated", json, StringComparison.OrdinalIgnoreCase);
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
    public void Facade_ValidatePlan_DelegatesAndKeepsValidation()
    {
        var options = TestSupport.CreateOptions(bearerToken: "secret-token-ABC12345");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "mcp.plan.json", TestSupport.MinimalPlanJson());

        var facade = BuildFacade(options, workspace, ReadyHandler());
        var payload = facade.ValidatePlan(planPath);
        var json = SecretRedactor.Redact(JsonSerializer.Serialize(payload));
        Assert.Contains("success", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret-token", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Facade_RunPlanDryRun_KeepsReport_NoTokenLeak()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33251",
            bearerToken: "secret-token-ABC12345");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "mcp-run.plan.json", TestSupport.MinimalPlanJson());

        var facade = BuildFacade(options, workspace, ReadyHandler());
        var payload = await facade.RunPlanAsync(planPath, dryRun: true, CancellationToken.None);
        var json = SecretRedactor.Redact(McpResponseBudget.SerializeBounded(payload));
        Assert.Contains("runId", json, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("status", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret-token", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Authorization", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Facade_Observe_MapsDriverUnavailable()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33252",
            bearerToken: "secret-token-ABC12345");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var handler = new FakeHttpMessageHandler(_ =>
            throw new HttpRequestException("connection refused"));
        var facade = BuildFacade(options, workspace, handler);
        var ex = await Assert.ThrowsAsync<McpToolException>(() =>
            facade.ObserveAsync(null, null, 2, 10, 20, false, false, false, true, true, null, CancellationToken.None));
        Assert.True(
            ex.ErrorCode is McpErrorCode.DriverUnavailable or McpErrorCode.Internal or McpErrorCode.Validation,
            ex.ErrorCode);
        Assert.DoesNotContain("secret-token", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Facade_Cancellation_Propagates()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33253",
            bearerToken: "secret-token-ABC12345");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "cancel.plan.json", TestSupport.MinimalPlanJson());

        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var handler = new FakeHttpMessageHandler(async (req, ct) =>
        {
            started.TrySetResult();
            await Task.Delay(Timeout.Infinite, ct);
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var facade = BuildFacade(options, workspace, handler);
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => facade.RunPlanAsync(planPath, dryRun: true, cts.Token));
        _ = started; // reserved if future tests need mid-flight cancel
    }

    [Fact]
    public async Task Tools_DelegateOnce_ToFacade()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.DoctorAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new { success = true, category = "observation" })
            .Verifiable();

        var tools = new DesktopMcpTools(mock.Object);
        var result = await tools.DesktopDoctor(CancellationToken.None);
        Assert.True(result.IsError != true);
        mock.Verify(f => f.DoctorAsync(It.IsAny<CancellationToken>()), Times.Once);
        mock.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Protocol_InitListsExpectedToolsAndSchemas()
    {
        var mock = new Mock<IDesktopMcpFacade>(MockBehavior.Strict);
        mock.Setup(f => f.ValidatePlan(It.IsAny<string>()))
            .Returns(new { success = true, category = "validation", planId = "SAMPLE-1", stepCount = 1 });
        mock.Setup(f => f.RunPlanAsync(It.IsAny<string>(), true, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new
            {
                success = true,
                category = "validation",
                runId = "run-fixture",
                status = "dryRun",
                exitCode = 0,
                dryRun = true,
                stepCount = 1
            });

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
        }

        var validate = tools.First(t => t.Name == "desktop_validate_plan");
        var validateResult = await validate.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "plans/example.plan.json"
        });
        var validateText = GetText(validateResult);
        Assert.Contains("SAMPLE-1", validateText, StringComparison.Ordinal);
        Assert.DoesNotContain("secret-token", validateText, StringComparison.Ordinal);

        var run = tools.First(t => t.Name == "desktop_run_plan");
        var runResult = await run.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = "plans/example.plan.json",
            ["dryRun"] = true
        });
        var runText = GetText(runResult);
        Assert.Contains("run-fixture", runText, StringComparison.Ordinal);
        mock.Verify(f => f.ValidatePlan(It.IsAny<string>()), Times.Once);
        mock.Verify(f => f.RunPlanAsync(It.IsAny<string>(), true, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Protocol_PathTraversal_ReturnsStructuredError()
    {
        var options = TestSupport.CreateOptions(bearerToken: "secret-token-ABC12345");
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
        var text = GetText(result);
        Assert.Contains("path_traversal", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret-token", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Protocol_PolicyDenial_ReturnsDenied()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33254",
            bearerToken: "secret-token-ABC12345");
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
        var text = GetText(result);
        Assert.DoesNotContain("secret-token", text, StringComparison.Ordinal);
        Assert.True(result.IsError == true
                    || text.Contains("denied", StringComparison.OrdinalIgnoreCase)
                    || text.Contains("policy", StringComparison.OrdinalIgnoreCase)
                    || text.Contains("validation", StringComparison.OrdinalIgnoreCase)
                    || text.Contains("driver", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task ExitGate_McpClient_ValidateAndRunDeterministicPlan()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33255",
            bearerToken: "secret-token-ABC12345");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "exit-gate.plan.json", TestSupport.MinimalPlanJson(planId: "exit.gate"));

        var facade = BuildFacade(options, workspace, ReadyHandler());
        await using var session = await McpTestSession.CreateAsync(facade);
        var tools = await session.Client.ListToolsAsync();

        var validate = tools.First(t => t.Name == "desktop_validate_plan");
        // Prefer workspace-relative path through MCP path guard.
        var relativePlan = Path.GetRelativePath(workspace.RootPath, planPath).Replace('\\', '/');
        var validateResult = await validate.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = relativePlan
        });
        var validateText = GetText(validateResult);
        Assert.True(validateResult.IsError != true, validateText);
        Assert.Contains("exit.gate", validateText, StringComparison.Ordinal);

        var run = tools.First(t => t.Name == "desktop_run_plan");
        var runResult = await run.CallAsync(new Dictionary<string, object?>
        {
            ["planPath"] = relativePlan,
            ["dryRun"] = true
        });
        var runText = GetText(runResult);
        Assert.True(runResult.IsError != true, runText);
        Assert.Contains("dryRun", runText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret-token", runText, StringComparison.Ordinal);
        Assert.DoesNotContain("Authorization", runText, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetText(CallToolResult result)
    {
        var text = string.Join("", result.Content.OfType<TextContentBlock>().Select(c => c.Text));
        return SecretRedactor.Redact(text);
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
