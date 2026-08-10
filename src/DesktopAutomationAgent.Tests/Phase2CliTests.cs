using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class Phase2CliTests
{
    [Fact]
    public async Task ValidatePlan_MakesNoHttpCalls()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "cli.plan.json", TestSupport.MinimalPlanJson());

        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var exit = await RunAsync(
            ["validate-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(0, calls);
        AssertExactlyOneJsonDocument(exit.Stdout.Trim());
    }

    [Fact]
    public async Task RunPlanDryRun_MakesGetCallsButNoPost()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "dry.plan.json", TestSupport.MinimalPlanJson());
        var handler = ReadyHandler();

        var exit = await RunAsync(
            ["run-plan", "--file", planPath, "--dry-run", "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Contains(handler.Requests, r => r.Method == HttpMethod.Get);
        Assert.DoesNotContain(handler.Requests, r => r.Method == HttpMethod.Post);
        AssertExactlyOneJsonDocument(exit.Stdout.Trim());
        Assert.DoesNotContain("secret-token", exit.Stdout);
    }

    [Fact]
    public async Task RunPlan_PerformsOrderedPostRequests()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [
              { "id": "one", "operation": "listwindows", "arguments": {} },
              { "id": "two", "operation": "listwindows", "arguments": {} }
            ]
            """);
        var planPath = TestSupport.WritePlan(options, "run.plan.json", plan);
        var handler = ReadyHandler();

        var exit = await RunAsync(
            ["run-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        var posts = handler.Requests.Where(r => r.Method == HttpMethod.Post).ToList();
        Assert.Equal(2, posts.Count);
        Assert.All(posts, p => Assert.Equal("/ui", p.RequestUri!.AbsolutePath));
        AssertExactlyOneJsonDocument(exit.Stdout.Trim());
    }

    [Fact]
    public void Parse_RejectsEmptyValidatePlanFileArgument()
    {
        var parsed = CommandLine.Parse(["validate-plan", "--file"]);
        Assert.NotNull(parsed.Error);
        Assert.Contains("--file requires a path", parsed.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Parse_AcceptsRunnerConfigArgs()
    {
        var parsed = CommandLine.Parse([
            "validate-plan",
            "--file",
            "automation/plans/example.plan.json",
            "--Runner:MaxPlanBytes=2048",
            "Runner__CleanupTimeoutSeconds=10"
        ]);
        Assert.Null(parsed.Error);
        Assert.Contains(parsed.ConfigurationArgs, a => a.Contains("Runner", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void HelpText_IsPhase3()
    {
        Assert.Contains("Phase 3", CommandLine.HelpText, StringComparison.Ordinal);
        Assert.Contains("validate-object-repository", CommandLine.HelpText, StringComparison.Ordinal);
        Assert.Contains("capture-page", CommandLine.HelpText, StringComparison.Ordinal);
    }

    private static FakeHttpMessageHandler ReadyHandler() =>
        new(req =>
        {
            if (req.Method == HttpMethod.Post)
            {
                return FakeHttpMessageHandler.Json(new { success = true, value = true });
            }

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

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunAsync(
        string[] args,
        AgentOptions options,
        IWorkspaceManager workspace,
        FakeHttpMessageHandler handler)
    {
        var factory = TestSupport.CreateFactory(handler);

        IHost HostBuilder(string[] _, bool jsonMode)
        {
            var builder = Host.CreateApplicationBuilder();
            builder.Logging.ClearProviders();
            builder.Logging.AddSimpleConsole();
            builder.Services.Configure<Microsoft.Extensions.Logging.Console.ConsoleLoggerOptions>(o =>
            {
                o.LogToStandardErrorThreshold = LogLevel.Trace;
            });
            if (jsonMode)
                builder.Logging.AddFilter("DesktopAutomationAgent", LogLevel.Warning);

            builder.Services.AddSingleton(Options.Create(options));
            builder.Services.AddSingleton(workspace);
            builder.Services.AddSingleton<ISuiteManifestReader>(_ =>
                new SuiteManifestReader(Options.Create(options), workspace));
            builder.Services.AddSingleton<PlanManifestReader>(_ =>
                new PlanManifestReader(Options.Create(options), workspace));
            builder.Services.AddSingleton<ObjectRepositoryReader>(_ =>
                new ObjectRepositoryReader(Options.Create(options), workspace));
            builder.Services.AddSingleton<ObjectReferenceResolver>();
            builder.Services.AddSingleton<PlanObjectReferenceExpander>();
            builder.Services.AddSingleton<PlanObjectRepositoryIntegrator>();
            builder.Services.AddSingleton<ObjectArtifactWriter>();
            builder.Services.AddSingleton<ObjectCandidateGenerator>();
            builder.Services.AddSingleton<ObjectCaptureService>();
            builder.Services.AddSingleton<ObjectVerificationService>();
            builder.Services.AddSingleton<IDriverConnectionResolver>(_ =>
                new DriverConnectionResolver(Options.Create(options), factory, NullLogger<DriverConnectionResolver>.Instance));
            builder.Services.AddSingleton<IDriverCatalogClient>(_ =>
                new DriverCatalogClient(Options.Create(options), factory, NullLogger<DriverCatalogClient>.Instance));
            builder.Services.AddSingleton<IDriverUiClient>(_ =>
                new DriverUiClient(Options.Create(options), factory, NullLogger<DriverUiClient>.Instance));
            builder.Services.AddSingleton<AssertionEvaluator>();
            builder.Services.AddSingleton<RunArtifactWriter>();
            builder.Services.AddSingleton<IDeterministicPlanRunner, DeterministicPlanRunner>();
            return builder.Build();
        }

        var originalOut = Console.Out;
        var originalErr = Console.Error;
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        Console.SetOut(stdout);
        Console.SetError(stderr);
        try
        {
            var code = await AgentCli.RunAsync(args, HostBuilder);
            return (code, stdout.ToString(), stderr.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    private static void AssertExactlyOneJsonDocument(string stdoutText)
    {
        var bytes = Encoding.UTF8.GetBytes(stdoutText);
        var reader = new Utf8JsonReader(bytes);
        Assert.True(JsonDocument.TryParseValue(ref reader, out var first));
        first!.Dispose();
        Assert.Equal(bytes.Length, reader.BytesConsumed);
    }
}
