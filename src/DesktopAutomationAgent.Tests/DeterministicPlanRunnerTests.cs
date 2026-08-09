using System.Net;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class DeterministicPlanRunnerTests
{
    [Fact]
    public async Task ExecutesStepsInExactOrder_OneRequestEach_NoRetry()
    {
        var postBodies = new List<string>();
        var handler = CreateReadyHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postBodies.Add(await req.Content!.ReadAsStringAsync());
                return FakeHttpMessageHandler.Json(new { success = true, value = "ok" });
            }

            return DefaultGet(req);
        });

        var options = ReadyOptions();
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [
              { "id": "a", "operation": "listwindows", "arguments": {} },
              { "id": "b", "operation": "listwindows", "arguments": {} }
            ]
            """);
        var path = TestSupport.WritePlan(options, "ordered.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("passed", report.Status);
        Assert.Equal(ExitCodes.Success, report.ExitCode);
        Assert.Equal(2, postBodies.Count);
        Assert.Equal(["a", "b"], report.Steps.Select(s => s.Id));
        Assert.All(report.Steps, s => Assert.Equal("passed", s.Status));
        Assert.Empty(report.OnFailureSteps);
    }

    [Fact]
    public async Task StopsAfterFirstFailure_MarksLaterSkipped_RunsCleanup()
    {
        var posts = 0;
        var handler = CreateReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            posts++;
            if (posts == 1)
            {
                return Task.FromResult(FakeHttpMessageHandler.Json(new
                {
                    success = false,
                    error = "boom"
                }));
            }

            return Task.FromResult(FakeHttpMessageHandler.Json(new { success = true }));
        });

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "FAIL-1",
              "name": "fail",
              "steps": [
                { "id": "launch", "operation": "launch", "arguments": { "value": "C:\\\\App.exe" } },
                { "id": "after", "operation": "listwindows", "arguments": {} },
                { "id": "quit", "operation": "quit", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "cleanup-quit", "operation": "quit", "arguments": {} }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "fail.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(ExitCodes.ExecutionFailure, report.ExitCode);
        Assert.Equal("failed", report.Steps[0].Status);
        Assert.Equal("skipped", report.Steps[1].Status);
        Assert.Equal("previousStepFailed", report.Steps[1].SkipReason);
        Assert.Equal("skipped", report.Steps[2].Status);
        Assert.Contains(report.OnFailureSteps, s => s.Id == "cleanup-quit");
        Assert.Equal(2, posts); // failed launch + cleanup quit
    }

    [Fact]
    public async Task DoesNotRunCleanupAfterSuccess()
    {
        var handler = CreateReadyHandler((req, _) =>
            Task.FromResult(req.Method == HttpMethod.Post
                ? FakeHttpMessageHandler.Json(new { success = true, value = true })
                : DefaultGet(req)));

        var options = ReadyOptions();
        var path = TestSupport.WritePlan(options, "ok.plan.json", TestSupport.MinimalPlanJson());
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);
        Assert.Equal("passed", report.Status);
        Assert.Empty(report.OnFailureSteps);
        var postBodies = new List<string>();
        foreach (var request in handler.Requests.Where(r => r.Method == HttpMethod.Post && r.Content is not null))
            postBodies.Add(await request.Content!.ReadAsStringAsync());
        Assert.DoesNotContain(postBodies, body => body.Contains("\"operation\":\"quit\"", StringComparison.Ordinal));
    }

    [Fact]
    public async Task PreservesPrimaryFailureWhenCleanupFails()
    {
        var posts = 0;
        var handler = CreateReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));
            posts++;
            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = posts == 1 ? "main-fail" : "cleanup-fail"
            }));
        });

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "FAIL-2",
              "name": "fail",
              "steps": [
                { "id": "launch", "operation": "launch", "arguments": { "value": "C:\\\\App.exe" } },
                { "id": "quit", "operation": "quit", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "cleanup-quit", "operation": "quit", "arguments": {} }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "fail2.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);
        Assert.Equal(ExitCodes.ExecutionFailure, report.ExitCode);
        Assert.Contains("main-fail", report.Failure!.Message);
        Assert.Equal("launch", report.Failure.StepId);
        Assert.Contains(report.OnFailureSteps, s => !s.Success);
    }

    [Fact]
    public async Task AttemptsCleanupAfterUncertainLaunchResult()
    {
        var posts = 0;
        var handler = CreateReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));
            posts++;
            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = posts == 1 ? false : true,
                error = posts == 1 ? "launch uncertain" : null
            }));
        });

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "UNCERTAIN",
              "name": "uncertain",
              "steps": [
                { "id": "launch", "operation": "launch", "arguments": { "value": "C:\\\\App.exe" } },
                { "id": "quit", "operation": "quit", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "cleanup-quit", "operation": "quit", "arguments": {} }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "uncertain.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);
        Assert.Equal(2, posts);
        Assert.Contains(report.OnFailureSteps, s => s.Id == "cleanup-quit" && !s.Skipped);
    }

    [Fact]
    public async Task DryRunPerformsNoUiPosts()
    {
        var handler = CreateReadyHandler((req, _) => Task.FromResult(DefaultGet(req)));
        var options = ReadyOptions();
        var path = TestSupport.WritePlan(options, "dry.plan.json", TestSupport.MinimalPlanJson());
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: true);
        Assert.Equal("validated", report.Status);
        Assert.Equal(ExitCodes.Success, report.ExitCode);
        Assert.DoesNotContain(handler.Requests, r => r.Method == HttpMethod.Post);
    }

    [Fact]
    public async Task CancellationProducesExitCode7()
    {
        using var cts = new CancellationTokenSource();
        var handler = CreateReadyHandler(async (req, ct) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                cts.Cancel();
                await Task.Delay(TimeSpan.FromSeconds(30), ct);
                return FakeHttpMessageHandler.Json(new { success = true });
            }

            return DefaultGet(req);
        });
        var options = ReadyOptions();
        var path = TestSupport.WritePlan(options, "cancel.plan.json", TestSupport.MinimalPlanJson());
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false, cts.Token);
        Assert.Equal(ExitCodes.Cancelled, report.ExitCode);
        Assert.Equal("cancelled", report.Status);
    }

    [Fact]
    public async Task AssertionFailureProducesExitCode6()
    {
        var handler = CreateReadyHandler((req, _) =>
            Task.FromResult(req.Method == HttpMethod.Post
                ? FakeHttpMessageHandler.Json(new { success = true, value = "Nope" })
                : DefaultGet(req)));

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "ASSERT-1",
              "name": "assert",
              "steps": [
                {
                  "id": "check",
                  "operation": "listwindows",
                  "arguments": {},
                  "assertions": [
                    { "path": "", "operator": "equals", "expected": "Dashboard" }
                  ]
                }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "assert.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);
        Assert.Equal(ExitCodes.ExecutionFailure, report.ExitCode);
        Assert.Equal(UiFailureClassification.AssertionFailure, report.Failure!.Classification);
    }

    [Fact]
    public async Task PlanHashAndUniqueRunDirectoryArePreserved()
    {
        var handler = CreateReadyHandler((req, _) =>
            Task.FromResult(req.Method == HttpMethod.Post
                ? FakeHttpMessageHandler.Json(new { success = true, value = true })
                : DefaultGet(req)));
        var options = ReadyOptions();
        var json = TestSupport.MinimalPlanJson();
        var path = TestSupport.WritePlan(options, "hashrun.plan.json", json);
        var runner = CreateRunner(options, handler);
        var first = await runner.RunAsync(path, dryRun: false);
        var second = await runner.RunAsync(path, dryRun: false);
        Assert.NotEqual(first.RunId, second.RunId);
        Assert.False(string.IsNullOrWhiteSpace(first.PlanSha256));
        Assert.Equal(first.PlanSha256, second.PlanSha256);
        Assert.True(File.Exists(Path.Combine(options.Workspace.Root, "runs", first.RunId, "run.json")));
        Assert.True(File.Exists(Path.Combine(options.Workspace.Root, "runs", second.RunId, "run.json")));
    }

    [Fact]
    public async Task SensitiveAndCaptureResponseRulesAreApplied()
    {
        var handler = CreateReadyHandler((req, _) =>
            Task.FromResult(req.Method == HttpMethod.Post
                ? FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new { password = "secret", title = "ok" }
                })
                : DefaultGet(req)));

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "SEC-1",
              "name": "sec",
              "steps": [
                {
                  "id": "sensitive",
                  "operation": "listwindows",
                  "arguments": { "password": "p@ss" },
                  "sensitive": true,
                  "captureResponse": true
                },
                {
                  "id": "capture",
                  "operation": "listwindows",
                  "arguments": {},
                  "captureResponse": true
                },
                {
                  "id": "nocapture",
                  "operation": "listwindows",
                  "arguments": {},
                  "captureResponse": false
                }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "sec.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);
        var reportJson = await File.ReadAllTextAsync(
            Path.Combine(options.Workspace.Root, "runs", report.RunId, "run.json"));

        Assert.DoesNotContain("p@ss", reportJson);
        Assert.DoesNotContain("test-token", reportJson);
        Assert.DoesNotContain("Authorization", reportJson, StringComparison.OrdinalIgnoreCase);
        Assert.Null(report.Steps[0].Arguments);
        Assert.Null(report.Steps[0].ResponseValue);
        Assert.NotNull(report.Steps[1].ResponseValue);
        Assert.Null(report.Steps[2].ResponseValue);
        Assert.Contains("[REDACTED]", reportJson);
    }

    private static AgentOptions ReadyOptions() =>
        TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: "test-token");

    private static DeterministicPlanRunner CreateRunner(AgentOptions options, FakeHttpMessageHandler handler)
    {
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        return new DeterministicPlanRunner(
            TestSupport.Wrap(options),
            TestSupport.CreatePlanReader(options, workspace),
            workspace,
            new DriverConnectionResolver(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverCatalogClient>.Instance),
            new DriverUiClient(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverUiClient>.Instance),
            new AssertionEvaluator(TestSupport.Wrap(options)),
            new RunArtifactWriter(),
            NullLogger<DeterministicPlanRunner>.Instance);
    }

    private static FakeHttpMessageHandler CreateReadyHandler(
        Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> onRequest) =>
        new(onRequest);

    private static HttpResponseMessage DefaultGet(HttpRequestMessage req)
    {
        if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
        {
            return FakeHttpMessageHandler.Json(new
            {
                sessionId = (string?)null,
                status = 0,
                value = new
                {
                    ready = true,
                    message = "ok",
                    build = new { version = "1.0.105", revision = "", time = "2026-01-01T00:00:00Z" }
                }
            });
        }

        if (req.RequestUri.AbsolutePath.Contains("operations", StringComparison.OrdinalIgnoreCase))
        {
            return FakeHttpMessageHandler.Json(new
            {
                success = true,
                value = CatalogFixtures.Phase2Catalog()
            });
        }

        return new HttpResponseMessage(HttpStatusCode.NotFound);
    }
}
