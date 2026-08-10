using System.Net;
using System.Net.Http.Headers;
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

public class ReviewGapFixTests
{
    [Fact]
    public void RedactReport_RedactsNestedSecretsAndBearerScalars()
    {
        var report = new RunReport
        {
            ReportSchemaVersion = 1,
            RunId = "run-1",
            Status = "passed",
            ExitCode = 0,
            PlanPath = "plans/a.plan.json",
            StartedAtUtc = DateTimeOffset.UtcNow,
            Steps =
            [
                new StepRunResult
                {
                    Sequence = 1,
                    Id = "s1",
                    Operation = "listwindows",
                    Phase = "steps",
                    Status = "passed",
                    Success = true,
                    CaptureResponse = true,
                    Arguments = new Dictionary<string, JsonElement>
                    {
                        ["nested"] = JsonSerializer.SerializeToElement(new
                        {
                            password = "p@ss",
                            apiKey = "key-1",
                            connectionString = "Server=.;Password=x",
                            note = "Authorization: Bearer super-secret-token-value"
                        })
                    },
                    ResponseValue = JsonSerializer.SerializeToElement(new
                    {
                        message = "Bearer super-secret-token-value accepted",
                        secret = "hidden",
                        title = "ok"
                    }),
                    Assertions =
                    [
                        new AssertionRunResult
                        {
                            Path = "",
                            Operator = "equals",
                            Passed = true,
                            Expected = JsonSerializer.SerializeToElement("Bearer super-secret-token-value"),
                            Actual = JsonSerializer.SerializeToElement("Bearer super-secret-token-value")
                        }
                    ]
                }
            ]
        };

        var redacted = RunArtifactWriter.RedactReport(report);
        var json = JsonSerializer.Serialize(redacted);
        Assert.Contains("[REDACTED]", json);
        Assert.DoesNotContain("p@ss", json);
        Assert.DoesNotContain("key-1", json);
        Assert.DoesNotContain("Server=.;Password=x", json);
        Assert.DoesNotContain("super-secret-token-value", json);
        Assert.Contains("title", json);
    }

    [Fact]
    public async Task RunPlanJson_StdoutUsesRedactedReport()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token-value");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SEC-JSON",
              "name": "sec",
              "steps": [
                {
                  "id": "one",
                  "operation": "listwindows",
                  "arguments": { "password": "p@ss", "note": "Bearer secret-token-value" },
                  "captureResponse": true
                }
              ]
            }
            """;
        var planPath = TestSupport.WritePlan(options, "sec-json.plan.json", plan);
        var handler = new FakeHttpMessageHandler(req =>
        {
            if (req.Method == HttpMethod.Post)
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new { message = "Bearer secret-token-value", title = "ok" }
                });
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

        var result = await RunCliAsync(
            ["run-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, result.ExitCode);
        Assert.DoesNotContain("p@ss", result.Stdout);
        Assert.DoesNotContain("secret-token-value", result.Stdout);
        using var doc = JsonDocument.Parse(result.Stdout.Trim());
        Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);

        var runJson = Directory.EnumerateFiles(Path.Combine(options.Workspace.Root, "runs"), "run.json", SearchOption.AllDirectories)
            .Select(File.ReadAllText)
            .Single();
        Assert.DoesNotContain("p@ss", runJson);
        Assert.DoesNotContain("secret-token-value", runJson);
    }

    [Fact]
    public async Task RunPlanJson_InvalidConfiguration_EmitsOneJsonDocument()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "token");
        options.Runner.MaxPlanBytes = -1;
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "cfg.plan.json", TestSupport.MinimalPlanJson());
        var handler = new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));

        var result = await RunCliAsync(
            ["run-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.UsageOrConfiguration, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout.Trim());
        Assert.False(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(ExitCodes.UsageOrConfiguration, doc.RootElement.GetProperty("exitCode").GetInt32());
        Assert.Empty(handler.Requests);
    }

    [Fact]
    public async Task ValidatePlan_MalformedJson_NoHttp_Exit5()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var planPath = TestSupport.WritePlan(options, "broken.plan.json", "{not-json");
        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var result = await RunCliAsync(
            ["validate-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.SuiteOrWorkspace, result.ExitCode);
        Assert.Equal(0, calls);
        using var doc = JsonDocument.Parse(result.Stdout.Trim());
        Assert.False(doc.RootElement.GetProperty("isValid").GetBoolean());
    }

    [Fact]
    public async Task DriverUiClient_RejectsStreamingOversizedBodyWithoutFullBuffer()
    {
        var options = TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: "test-token");
        options.Runner.MaxResponseBytes = 2048;
        var tracking = new TrackingReadStream();
        var handler = new FakeHttpMessageHandler(_ =>
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(tracking)
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            return response;
        });

        var client = new DriverUiClient(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverUiClient>.Instance);

        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            client.ExecuteStepAsync(
                new DriverConnection
                {
                    BaseUri = new Uri("http://127.0.0.1:33201/"),
                    BearerToken = "test-token",
                    DiscoveryMethod = "explicit"
                },
                new PlanStep { Id = "s1", Operation = "listwindows", Arguments = new Dictionary<string, JsonElement>() }));

        Assert.Contains("maximum size", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(tracking.BytesRead > options.Runner.MaxResponseBytes);
        Assert.True(
            tracking.BytesRead < 1024 * 1024,
            $"Expected bounded streaming read; bytesRead={tracking.BytesRead}");
    }

    [Theory]
    [InlineData("run-plan", "--json")]
    [InlineData("run-plan", "--json", "--file")]
    [InlineData("validate-plan", "--json", "--bogus")]
    public async Task UsageErrorsWithJsonFlag_EmitOneJsonDocument(params string[] args)
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var handler = new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));

        var result = await RunCliAsync(args, options, workspace, handler);
        Assert.Equal(ExitCodes.UsageOrConfiguration, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout.Trim());
        Assert.False(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(ExitCodes.UsageOrConfiguration, doc.RootElement.GetProperty("exitCode").GetInt32());
        Assert.Empty(handler.Requests);
    }

    [Fact]
    public void ExamplePlan_AssertsRootValueNotWindowsPath()
    {
        var json = ReadRepoFile("automation/plans/example.plan.json");
        using var doc = JsonDocument.Parse(json);
        var assertion = doc.RootElement.GetProperty("steps")[0].GetProperty("assertions")[0];
        Assert.Equal("", assertion.GetProperty("path").GetString());
        Assert.Equal("isNotNull", assertion.GetProperty("operator").GetString());
        Assert.False(doc.RootElement.GetProperty("steps")[0].GetProperty("arguments").TryGetProperty("limit", out _));
    }

    private static string ReadRepoFile(string relativePath)
    {
        var candidates = new[]
        {
            Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), relativePath)),
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", relativePath)),
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", relativePath))
        };
        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
                return File.ReadAllText(candidate);
        }

        throw new FileNotFoundException($"Unable to locate {relativePath}");
    }

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunCliAsync(
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

    private sealed class TrackingReadStream : Stream
    {
        public long BytesRead { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => BytesRead;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var n = Math.Max(1, count);
            Array.Fill(buffer, (byte)'x', offset, n);
            BytesRead += n;
            return n;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
