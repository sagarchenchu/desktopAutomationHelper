using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Jira;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public sealed class Phase4eJiraSuiteTests
{
    [Fact]
    public void SecretRedactor_Redacts_Basic_And_ApiToken()
    {
        var text = """Authorization: Basic abcdef123456 "apiToken":"super-secret-token" password=hunter2""";
        var redacted = SecretRedactor.Redact(text);
        Assert.DoesNotContain("abcdef123456", redacted, StringComparison.Ordinal);
        Assert.DoesNotContain("super-secret-token", redacted, StringComparison.Ordinal);
        Assert.DoesNotContain("hunter2", redacted, StringComparison.Ordinal);
        Assert.Contains("[REDACTED]", redacted, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Invalid_Key_Rejected_Before_Http()
    {
        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return FakeHttpMessageHandler.Json(new { });
        });
        using var temp = TempDir();
        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);

        var result = await source.FetchAsync("not a key");
        Assert.False(result.Success);
        Assert.Equal(ExitCodes.UsageOrConfiguration, result.ExitCode);
        Assert.Contains("rejected before network", result.Errors[0], StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task Vendor_Mapping_PlainFields_And_Xray_FeaturePath()
    {
        using var tempPlain = TempDir();
        using var tempXray = TempDir();
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";

        var plainHandler = new FakeHttpMessageHandler(req =>
        {
            Assert.Equal(HttpMethod.Get, req.Method);
            Assert.DoesNotContain("Authorization: Bearer secret", req.Headers.ToString(), StringComparison.Ordinal);
            return FakeHttpMessageHandler.Json(new
            {
                key = "DEMO-1",
                fields = new
                {
                    summary = "Plain fields fixture",
                    description = feature,
                    updated = "2026-11-12T00:00:00.000+0000",
                    labels = new[] { "smoke" },
                    unusedSecretField = "should-be-redacted-from-cache"
                }
            });
        });

        var plain = CreateSource(tempPlain.Path, plainHandler, mappingProfile: "plain-fields", allowLocalHttp: true);
        var plainResult = await plain.FetchAsync("DEMO-1");
        Assert.True(plainResult.Success, string.Join("; ", plainResult.Errors));
        Assert.Equal("DEMO-1", plainResult.Document!.JiraKey);
        Assert.Contains("Feature:", plainResult.Document.Feature, StringComparison.Ordinal);
        Assert.Contains(
            plainResult.Metadata!.RedactedFieldPaths,
            p => p.Contains("unusedSecretField", StringComparison.OrdinalIgnoreCase));

        var xrayHandler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            key = "DEMO-1",
            fields = new Dictionary<string, object?>
            {
                ["summary"] = "Xray mapped",
                ["description"] = "ignored",
                ["updated"] = "2026-11-12T00:00:00.000+0000",
                ["customfield_10040"] = feature
            }
        }));
        var xrayOpts = CreateJiraOptions(allowLocalHttp: true, mappingProfile: "xray");
        xrayOpts.FieldMapping = new JiraFieldMappingOptions
        {
            FeaturePath = "fields.customfield_10040"
        };
        var xray = CreateSource(tempXray.Path, xrayHandler, xrayOpts);
        var xrayResult = await xray.FetchAsync("DEMO-1");
        Assert.True(xrayResult.Success, string.Join("; ", xrayResult.Errors));
        Assert.Equal(feature, xrayResult.Document!.Feature);
    }

    [Fact]
    public async Task Pagination_Aggregates_Search_Pages()
    {
        using var temp = TempDir();
        var page = 0;
        var handler = new FakeHttpMessageHandler(req =>
        {
            page++;
            Assert.Contains("/rest/api/3/search?", req.RequestUri!.PathAndQuery, StringComparison.Ordinal);
            if (page == 1)
            {
                return FakeHttpMessageHandler.Json(new
                {
                    startAt = 0,
                    maxResults = 1,
                    total = 2,
                    issues = new[] { new { key = "DEMO-1" } }
                });
            }

            return FakeHttpMessageHandler.Json(new
            {
                startAt = 1,
                maxResults = 1,
                total = 2,
                issues = new[] { new { key = "DEMO-2" } }
            });
        });

        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.PageSize = 1;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.SearchKeysAsync("project = DEMO");
        Assert.True(result.Success, string.Join("; ", result.Errors));
        Assert.Equal(2, result.PageCount);
        Assert.Equal(["DEMO-1", "DEMO-2"], result.Keys);
    }

    [Theory]
    [InlineData(HttpStatusCode.Unauthorized, ExitCodes.AuthOrCatalog)]
    [InlineData(HttpStatusCode.Forbidden, ExitCodes.AuthOrCatalog)]
    [InlineData(HttpStatusCode.NotFound, ExitCodes.SuiteOrWorkspace)]
    [InlineData(HttpStatusCode.TooManyRequests, ExitCodes.ExecutionFailure)]
    [InlineData(HttpStatusCode.InternalServerError, ExitCodes.DriverUnavailable)]
    public async Task Http_Status_Mapped(HttpStatusCode status, int exitCode)
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ =>
        {
            var response = new HttpResponseMessage(status);
            if (status == HttpStatusCode.TooManyRequests)
                response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(1));
            response.Headers.TryAddWithoutValidation("X-RateLimit-Limit", "100");
            response.Headers.TryAddWithoutValidation("X-RateLimit-Remaining", "0");
            response.Content = new StringContent("{}", Encoding.UTF8, "application/json");
            return response;
        });
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.MaxRetries = 0;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.FetchAsync("DEMO-1");
        Assert.False(result.Success);
        Assert.Equal(exitCode, result.ExitCode);
        Assert.NotNull(result.RateLimit);
    }

    [Fact]
    public async Task Size_Limit_Rejects_Oversized_Response()
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ =>
            FakeHttpMessageHandler.Text(new string('x', 4096)));
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.MaxResponseBytes = 1024;
        opts.MaxRetries = 0;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.FetchAsync("DEMO-1");
        Assert.False(result.Success);
        Assert.Contains("MaxResponseBytes", result.Errors[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Cache_Hit_Uses_ETag_NotModified()
    {
        using var temp = TempDir();
        var calls = 0;
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var handler = new FakeHttpMessageHandler(req =>
        {
            calls++;
            if (calls == 1)
            {
                var response = FakeHttpMessageHandler.Json(new
                {
                    key = "DEMO-1",
                    fields = new
                    {
                        summary = "Cached",
                        description = feature,
                        updated = "2026-11-12T00:00:00.000+0000"
                    }
                });
                response.Headers.ETag = new System.Net.Http.Headers.EntityTagHeaderValue("\"rev1\"");
                return response;
            }

            Assert.True(req.Headers.IfNoneMatch.Count > 0);
            return new HttpResponseMessage(HttpStatusCode.NotModified);
        });

        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);
        var first = await source.FetchAsync("DEMO-1");
        Assert.True(first.Success, string.Join("; ", first.Errors));
        var second = await source.FetchAsync("DEMO-1");
        Assert.True(second.Success, string.Join("; ", second.Errors));
        Assert.True(second.NotModified);
        Assert.True(second.FromCache);
        Assert.Equal(2, calls);
    }

    [Fact]
    public async Task Cache_Does_Not_Dump_Unrestricted_Fields()
    {
        using var temp = TempDir();
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            key = "DEMO-1",
            fields = new
            {
                summary = "Cached",
                description = feature,
                updated = "2026-11-12T00:00:00.000+0000",
                comment = new { comments = new[] { new { body = "secret comment" } } },
                attachment = new[] { new { filename = "secret.bin" } }
            }
        }));
        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);
        var result = await source.FetchAsync("DEMO-1");
        Assert.True(result.Success, string.Join("; ", result.Errors));

        var cacheFiles = Directory.GetFiles(Path.Combine(temp.Path, "jira-cache"), "*", SearchOption.AllDirectories);
        Assert.Contains(cacheFiles, f => f.EndsWith("test-case.json", StringComparison.OrdinalIgnoreCase));
        foreach (var file in cacheFiles)
        {
            var text = File.ReadAllText(file);
            Assert.DoesNotContain("secret comment", text, StringComparison.Ordinal);
            Assert.DoesNotContain("secret.bin", text, StringComparison.Ordinal);
            Assert.DoesNotContain("\"comment\"", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task Suite_Sequential_FailFast_And_ContinueOnFailure()
    {
        using var temp = TempDir();
        SeedWorkspace(temp.Path);
        WriteSuite(temp.Path, enabledKeys: ["DEMO-1", "DEMO-2"], disabledKey: "DEMO-99");

        var runnerCalls = new List<string>();
        var runner = new FakePlanRunner((path, dryRun, _) =>
        {
            runnerCalls.Add(path);
            var planId = Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(path))!)!;
            // Fail DEMO-1 execution.
            if (path.Contains("DEMO-1", StringComparison.OrdinalIgnoreCase))
            {
                return Task.FromResult(new RunReport
                {
                    RunId = "run-fail",
                    Status = "failed",
                    ExitCode = ExitCodes.ExecutionFailure,
                    PlanPath = path,
                    DryRun = dryRun,
                    PlanSha256 = "a"
                });
            }

            return Task.FromResult(new RunReport
            {
                RunId = "run-ok",
                Status = "passed",
                ExitCode = ExitCodes.Success,
                PlanPath = path,
                DryRun = dryRun,
                PlanSha256 = "b"
            });
        });

        var orch = CreateOrchestrator(temp.Path, runner);

        var failFast = await orch.RunSuiteAsync(new JiraSuiteRunRequest
        {
            SuiteFile = "suites/phase4e-demo.json",
            DryRun = true,
            FromJira = false,
            ContinueOnFailure = false
        });
        Assert.False(failFast.Success);
        Assert.Equal("fail-fast", failFast.StopReason);
        Assert.Single(failFast.Keys);
        Assert.Equal("DEMO-1", failFast.Keys[0].JiraKey);
        Assert.Equal(1, failFast.DisabledExcluded);
        Assert.True(File.Exists(Path.Combine(temp.Path, failFast.SuiteReportPath!.Replace('/', Path.DirectorySeparatorChar))));

        runnerCalls.Clear();
        var cont = await orch.RunSuiteAsync(new JiraSuiteRunRequest
        {
            SuiteFile = "suites/phase4e-demo.json",
            DryRun = true,
            FromJira = false,
            ContinueOnFailure = true,
            MaxFailures = 5
        });
        Assert.Equal(2, cont.Keys.Count);
        Assert.Equal(["DEMO-1", "DEMO-2"], cont.Keys.Select(k => k.JiraKey).ToArray());
        Assert.DoesNotContain(cont.Keys, k => k.JiraKey == "DEMO-99");
    }

    [Fact]
    public async Task Compile_Failure_Blocks_Execution()
    {
        using var temp = TempDir();
        SeedWorkspace(temp.Path);
        File.WriteAllText(Path.Combine(temp.Path, "test-cases", "examples", "DEMO-8.json"), """
            {
              "schemaVersion": 1,
              "jiraKey": "DEMO-8",
              "revision": "1",
              "summary": "Unmapped",
              "feature": "Feature: X\n  Scenario: Y\n    Given this step is not in the registry\n"
            }
            """);

        var runnerCalls = 0;
        var runner = new FakePlanRunner((_, _, _) =>
        {
            runnerCalls++;
            return Task.FromResult(new RunReport
            {
                RunId = "should-not-run",
                Status = "passed",
                ExitCode = ExitCodes.Success,
                PlanPath = "x"
            });
        });
        var orch = CreateOrchestrator(temp.Path, runner);
        var result = await orch.RunKeyAsync(new JiraRunKeyRequest
        {
            JiraKey = "DEMO-8",
            DryRun = true,
            FromJira = false
        });
        Assert.False(result.Success);
        Assert.Equal("compile-failed", result.Result);
        Assert.Equal(0, runnerCalls);
    }

    [Fact]
    public async Task DryRun_Does_Not_Post_Ui()
    {
        using var temp = TempDir();
        SeedWorkspace(temp.Path);
        var posts = 0;
        var handler = new FakeHttpMessageHandler(req =>
        {
            if (req.Method == HttpMethod.Post)
            {
                posts++;
                return FakeHttpMessageHandler.Json(new { success = true });
            }

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
                        build = new { version = "1.0.105", revision = "", time = "2026-11-12T00:00:00Z" }
                    }
                });
            }

            if (req.RequestUri.AbsolutePath.Contains("operations", StringComparison.OrdinalIgnoreCase))
                return FakeHttpMessageHandler.Json(new { success = true, value = CatalogFixtures.Phase2Catalog() });

            return new HttpResponseMessage(HttpStatusCode.NotFound);
        });

        var options = TestSupport.CreateOptions(
            workspaceRoot: temp.Path,
            baseUrl: "http://127.0.0.1:39999",
            bearerToken: "test-token");
        var workspace = new WorkspaceManager(Options.Create(options), NullLogger<WorkspaceManager>.Instance);
        workspace.Initialize();
        SeedWorkspace(temp.Path);

        var factory = TestSupport.CreateFactory(handler);
        var runner = new DeterministicPlanRunner(
            Options.Create(options),
            new Plans.PlanManifestReader(Options.Create(options), workspace),
            new PlanObjectRepositoryIntegrator(
                new ObjectRepositoryReader(Options.Create(options), workspace),
                new PlanObjectReferenceExpander(new ObjectReferenceResolver())),
            workspace,
            new DriverConnectionResolver(Options.Create(options), factory, NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(Options.Create(options), factory, NullLogger<DriverCatalogClient>.Instance),
            new DriverUiClient(Options.Create(options), factory, NullLogger<DriverUiClient>.Instance),
            new AssertionEvaluator(Options.Create(options)),
            new RunArtifactWriter(),
            NullLogger<DeterministicPlanRunner>.Instance);

        var orch = new JiraSuiteOrchestrator(
            Options.Create(options),
            workspace,
            new StubJiraSource(),
            new BddCompilationService(workspace, new ObjectRepositoryReader(Options.Create(options), workspace), Options.Create(options)),
            new SuiteManifestReader(Options.Create(options), workspace),
            runner,
            NullLogger<JiraSuiteOrchestrator>.Instance);

        var result = await orch.RunKeyAsync(new JiraRunKeyRequest
        {
            JiraKey = "DEMO-1",
            DryRun = true,
            FromJira = false
        });

        Assert.True(result.Success, string.Join("; ", result.Errors));
        Assert.Equal(0, posts);
    }

    [Fact]
    public async Task Cancellation_Writes_Partial_Suite_Report()
    {
        using var temp = TempDir();
        SeedWorkspace(temp.Path);
        WriteSuite(temp.Path, enabledKeys: ["DEMO-1", "DEMO-2"], disabledKey: "DEMO-99");

        using var cts = new CancellationTokenSource();
        var runner = new FakePlanRunner((_, _, ct) =>
        {
            cts.Cancel();
            ct.ThrowIfCancellationRequested();
            return Task.FromResult(new RunReport
            {
                RunId = "x",
                Status = "passed",
                ExitCode = 0,
                PlanPath = "x"
            });
        });

        var orch = CreateOrchestrator(temp.Path, runner);
        var result = await orch.RunSuiteAsync(new JiraSuiteRunRequest
        {
            SuiteFile = "suites/phase4e-demo.json",
            DryRun = true
        }, cts.Token);

        Assert.False(result.Success);
        Assert.Equal(ExitCodes.Cancelled, result.ExitCode);
        Assert.Equal("cancelled", result.StopReason);
        Assert.Contains(result.Keys, k => k.Result == "cancelled");
        Assert.True(File.Exists(Path.Combine(temp.Path, result.SuiteReportPath!.Replace('/', Path.DirectorySeparatorChar))));
    }

    [Fact]
    public void JiraUrlRules_Reject_Remote_Http_And_UserInfo()
    {
        Assert.Throws<AgentConfigurationException>(() =>
            JiraUrlRules.EnsureAllowed("http://evil.example.com", allowLocalHttpTestServer: true));
        Assert.Throws<AgentConfigurationException>(() =>
            JiraUrlRules.EnsureAllowed("http://127.0.0.1:8080", allowLocalHttpTestServer: false));
        Assert.Throws<AgentConfigurationException>(() =>
            JiraUrlRules.EnsureAllowed("https://user:secret@example.atlassian.net", allowLocalHttpTestServer: false));
        var ok = JiraUrlRules.EnsureAllowed("http://127.0.0.1:8080", allowLocalHttpTestServer: true);
        Assert.Equal(Uri.UriSchemeHttp, ok.Scheme);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(2, 3)]
    public async Task Retry_Exact_Request_Count(int maxRetries, int expectedRequests)
    {
        using var temp = TempDir();
        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            Interlocked.Increment(ref calls);
            return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
            {
                Content = new StringContent("{}", Encoding.UTF8, "application/json")
            };
        });
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.MaxRetries = maxRetries;
        opts.RequestTimeoutSeconds = 5;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.FetchAsync("DEMO-1");
        Assert.False(result.Success);
        Assert.Equal(expectedRequests, calls);
    }

    [Fact]
    public async Task Body_Read_Timeout_On_Slow_Stream()
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(async (_, ct) =>
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new SlowHttpContent(TimeSpan.FromSeconds(30))
            };
            await Task.Yield();
            return response;
        });
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.MaxRetries = 0;
        opts.RequestTimeoutSeconds = 1;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.FetchAsync("DEMO-1");
        Assert.False(result.Success);
        Assert.Equal(ExitCodes.ExecutionFailure, result.ExitCode);
        Assert.Contains("timed out", result.Errors[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Pagination_Truncates_When_MaxPages_Before_Total()
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            startAt = 0,
            maxResults = 1,
            total = 3,
            issues = new[] { new { key = "DEMO-1" } }
        }));
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.PageSize = 1;
        opts.MaxPages = 1;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.SearchKeysAsync("project = DEMO");
        Assert.False(result.Success);
        Assert.Contains("truncated", result.Errors[0], StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, result.PageCount);
        Assert.Equal(["DEMO-1"], result.Keys);
    }

    [Fact]
    public async Task Pagination_Malformed_Page_Json_Fails()
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Text("{not-json"));
        var opts = CreateJiraOptions(allowLocalHttp: true);
        opts.MaxRetries = 0;
        var source = CreateSource(temp.Path, handler, opts);
        var result = await source.SearchKeysAsync("project = DEMO");
        Assert.False(result.Success);
        Assert.Contains("JSON parse failed", result.Errors[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Pagination_Empty_Page_Before_Total_Is_Truncation()
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            startAt = 0,
            maxResults = 50,
            total = 3,
            issues = Array.Empty<object>()
        }));
        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);
        var result = await source.SearchKeysAsync("project = DEMO");
        Assert.False(result.Success);
        Assert.Contains("truncated", result.Errors[0], StringComparison.OrdinalIgnoreCase);
        Assert.Contains("empty page", result.Errors[0], StringComparison.OrdinalIgnoreCase);
        Assert.Empty(result.Keys);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Pagination_Rejects_Missing_Or_Malformed_Total(bool missing)
    {
        using var temp = TempDir();
        var handler = new FakeHttpMessageHandler(_ => missing
            ? FakeHttpMessageHandler.Json(new
            {
                startAt = 0,
                maxResults = 50,
                issues = new[] { new { key = "DEMO-1" } }
            })
            : FakeHttpMessageHandler.Text(
                """{"startAt":0,"maxResults":50,"total":"three","issues":[{"key":"DEMO-1"}]}"""));
        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);
        var result = await source.SearchKeysAsync("project = DEMO");
        Assert.False(result.Success);
        Assert.Contains("total", result.Errors[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Cache_Whitelist_Hash_And_Schema_On_Read()
    {
        using var temp = TempDir();
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            key = "DEMO-1",
            fields = new
            {
                summary = "Cached",
                description = feature,
                updated = "2026-11-12T00:00:00.000+0000",
                unusedVendorDump = new { password = "nope", token = "abc" }
            }
        }));
        var source = CreateSource(temp.Path, handler, allowLocalHttp: true);
        var first = await source.FetchAsync("DEMO-1");
        Assert.True(first.Success, string.Join("; ", first.Errors));
        Assert.False(string.IsNullOrWhiteSpace(first.Metadata!.ContentSha256));

        var cacheFiles = Directory.GetFiles(Path.Combine(temp.Path, "jira-cache"), "test-case.json", SearchOption.AllDirectories);
        Assert.Single(cacheFiles);
        var bytes = File.ReadAllBytes(cacheFiles[0]);
        var text = Encoding.UTF8.GetString(bytes);
        Assert.DoesNotContain("unusedVendorDump", text, StringComparison.Ordinal);
        Assert.DoesNotContain("password", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\"description\"", text, StringComparison.Ordinal);

        var validator = new JiraTestCaseValidator();
        var validated = validator.ValidateBytes(bytes, "cache/test-case.json");
        Assert.True(validated.IsValid, string.Join("; ", validated.Errors));
        Assert.Equal(first.Metadata.ContentSha256, validated.Sha256, ignoreCase: true);

        // Tamper hash metadata → read must fail.
        var metaPath = Path.Combine(Path.GetDirectoryName(cacheFiles[0])!, "source-metadata.json");
        var meta = File.ReadAllText(metaPath).Replace(first.Metadata.ContentSha256, new string('0', 64), StringComparison.OrdinalIgnoreCase);
        File.WriteAllText(metaPath, meta);
        var cache = new JiraRawCache(temp.Path);
        Assert.False(cache.TryRead("DEMO-1", null, out _, out _, out _, out var errors));
        Assert.Contains(errors, e => e.Contains("contentSha256 mismatch", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("abc")]
    [InlineData("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcde")] // 63
    [InlineData("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdefg")] // 65 with non-hex
    public void ContentSha256_Required_Valid_64_Hex(string? sha)
    {
        Assert.False(JiraTestCaseProjector.IsValidContentSha256(sha));
        var errors = new List<string>();
        Assert.False(JiraTestCaseProjector.TryValidateCachedBytes(
            Encoding.UTF8.GetBytes("{}"),
            "DEMO-1",
            sha,
            "cache/test-case.json",
            out _,
            out errors));
        Assert.Contains(errors, e => e.Contains("contentSha256 is required", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ContentSha256_Blank_Metadata_Fails_Cache_Read()
    {
        using var temp = TempDir();
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var doc = new JiraTestCaseDocument
        {
            SchemaVersion = 1,
            JiraKey = "DEMO-1",
            Revision = "1",
            Summary = "Blank hash",
            Feature = feature
        };
        var cache = new JiraRawCache(temp.Path);
        cache.WriteAtomic(
            "DEMO-1",
            doc,
            new JiraSourceMetadata { JiraKey = "DEMO-1", RevisionToken = "1", MappingProfile = "plain-fields" },
            out var directory,
            out var writtenMeta,
            out _);
        Assert.True(JiraTestCaseProjector.IsValidContentSha256(writtenMeta.ContentSha256));

        var metaPath = Path.Combine(directory, "source-metadata.json");
        var blanked = File.ReadAllText(metaPath)
            .Replace($"\"{writtenMeta.ContentSha256}\"", "\"\"", StringComparison.OrdinalIgnoreCase);
        File.WriteAllText(metaPath, blanked);
        Assert.False(cache.TryRead("DEMO-1", "1", out _, out _, out _, out var errors));
        Assert.Contains(errors, e => e.Contains("contentSha256 is required", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Nested_Arrays_Validated_Before_Projection()
    {
        var feature =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var withNullScenario = new JiraTestCaseDocument
        {
            SchemaVersion = 1,
            JiraKey = "DEMO-1",
            Revision = "1",
            Summary = "Bad",
            Feature = feature,
            Scenarios = [null!]
        };
        Assert.False(JiraTestCaseProjector.TrySanitizeAndHash(
            withNullScenario,
            "bad.json",
            out _,
            out _,
            out _,
            out var nullScenarioErrors));
        Assert.Contains(nullScenarioErrors, e => e.Contains("scenarios[0] must not be null", StringComparison.Ordinal));

        var withNullSteps = new JiraTestCaseDocument
        {
            SchemaVersion = 1,
            JiraKey = "DEMO-1",
            Revision = "1",
            Summary = "Bad",
            Feature = feature,
            Scenarios =
            [
                new StructuredScenario
                {
                    Name = "Broken",
                    Steps = null!
                }
            ]
        };
        Assert.False(JiraTestCaseProjector.TrySanitizeAndHash(
            withNullSteps,
            "bad.json",
            out _,
            out _,
            out _,
            out var nullStepsErrors));
        Assert.Contains(nullStepsErrors, e => e.Contains("scenarios[0].steps must not be null", StringComparison.Ordinal));

        var withNullStepEntry = new JiraTestCaseDocument
        {
            SchemaVersion = 1,
            JiraKey = "DEMO-1",
            Revision = "1",
            Summary = "Bad",
            Feature = feature,
            Scenarios =
            [
                new StructuredScenario
                {
                    Name = "Broken",
                    Steps = [null!]
                }
            ]
        };
        Assert.False(JiraTestCaseProjector.TrySanitizeAndHash(
            withNullStepEntry,
            "bad.json",
            out _,
            out _,
            out _,
            out var nullStepErrors));
        Assert.Contains(nullStepErrors, e => e.Contains("scenarios[0].steps[0] must not be null", StringComparison.Ordinal));

        // Null-safe projection still filters if called directly after structural rejection path.
        var projected = JiraTestCaseProjector.ProjectWhitelist(new JiraTestCaseDocument
        {
            SchemaVersion = 1,
            JiraKey = "DEMO-1",
            Revision = "1",
            Summary = "Ok",
            Feature = feature,
            Scenarios =
            [
                new StructuredScenario { Name = "Keep", Steps = [new StructuredStep { Keyword = "Given", Text = "x" }] },
                null!
            ]
        });
        Assert.NotNull(projected.Scenarios);
        Assert.Single(projected.Scenarios!);
    }

    [Fact]
    public async Task FromJira_MultiKey_Suite_Http_Cache_Compile_Run_Report()
    {
        using var temp = TempDir();
        SeedWorkspace(temp.Path);
        WriteSuite(temp.Path, enabledKeys: ["DEMO-1", "DEMO-2"], disabledKey: "DEMO-99");

        var feature1 =
            "Feature: List visible windows\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";
        var feature2 =
            "Feature: List visible windows again\n  Scenario: Enumerate windows\n    Given the desktop is ready\n    When I list visible windows\n    Then the window list is not empty\n";

        var handler = new FakeHttpMessageHandler(req =>
        {
            Assert.Equal(HttpMethod.Get, req.Method);
            var path = req.RequestUri!.AbsolutePath;
            if (path.Contains("/issue/DEMO-1", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    key = "DEMO-1",
                    fields = new { summary = "One", description = feature1, updated = "2026-11-12T00:00:00.000+0000" }
                });
            }

            if (path.Contains("/issue/DEMO-2", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    key = "DEMO-2",
                    fields = new { summary = "Two", description = feature2, updated = "2026-11-12T00:00:00.000+0000" }
                });
            }

            return new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent("{}", Encoding.UTF8, "application/json")
            };
        });

        var options = TestSupport.CreateOptions(workspaceRoot: temp.Path);
        options.Jira = CreateJiraOptions(allowLocalHttp: true);
        var workspace = new WorkspaceManager(Options.Create(options), NullLogger<WorkspaceManager>.Instance);
        var jiraSource = new JiraRestTestCaseSource(
            TestSupport.CreateFactory(handler),
            options,
            temp.Path,
            NullLogger<JiraRestTestCaseSource>.Instance);

        var runOrder = new List<string>();
        var runner = new FakePlanRunner((path, dryRun, _) =>
        {
            runOrder.Add(path);
            Assert.True(dryRun);
            Assert.DoesNotContain("/ui", path, StringComparison.OrdinalIgnoreCase);
            return Task.FromResult(new RunReport
            {
                RunId = "run-" + runOrder.Count,
                Status = "passed",
                ExitCode = ExitCodes.Success,
                PlanPath = path,
                DryRun = true,
                PlanSha256 = "planhash" + runOrder.Count
            });
        });

        var orch = new JiraSuiteOrchestrator(
            Options.Create(options),
            workspace,
            jiraSource,
            new BddCompilationService(workspace, new ObjectRepositoryReader(Options.Create(options), workspace), Options.Create(options)),
            new SuiteManifestReader(Options.Create(options), workspace),
            runner,
            NullLogger<JiraSuiteOrchestrator>.Instance);

        var result = await orch.RunSuiteAsync(new JiraSuiteRunRequest
        {
            SuiteFile = "suites/phase4e-demo.json",
            DryRun = true,
            FromJira = true
        });

        Assert.True(result.Success, string.Join("; ", result.Errors) + " | keys=" +
            string.Join(",", result.Keys.Select(k => k.JiraKey + ":" + k.Result + ":" + string.Join('/', k.Errors))));
        Assert.Equal(2, result.Keys.Count);
        Assert.Equal(["DEMO-1", "DEMO-2"], result.Keys.Select(k => k.JiraKey).ToArray());
        Assert.Equal(1, result.DisabledExcluded);
        Assert.All(result.Keys, k =>
        {
            Assert.False(string.IsNullOrWhiteSpace(k.CompilationFingerprint));
            Assert.False(string.IsNullOrWhiteSpace(k.PlanHash));
            Assert.False(string.IsNullOrWhiteSpace(k.RunId));
            Assert.Equal("passed", k.Result);
        });
        Assert.Equal(2, runOrder.Count);
        Assert.True(Directory.Exists(Path.Combine(temp.Path, "jira-cache", "DEMO-1")));
        Assert.True(Directory.Exists(Path.Combine(temp.Path, "jira-cache", "DEMO-2")));
        Assert.True(File.Exists(Path.Combine(temp.Path, result.SuiteReportPath!.Replace('/', Path.DirectorySeparatorChar))));
        Assert.Equal(2, handler.Requests.Count);
        Assert.All(handler.Requests, r => Assert.Null(r.Content)); // GET only — no Jira write
    }

    [Fact]
    public void WorkspaceInit_Creates_JiraCache()
    {
        using var temp = TempDir();
        var workspace = new WorkspaceManager(
            Options.Create(TestSupport.CreateOptions(workspaceRoot: temp.Path)),
            NullLogger<WorkspaceManager>.Instance);
        workspace.Initialize();
        Assert.True(Directory.Exists(Path.Combine(temp.Path, "jira-cache")));
        Assert.True(File.Exists(Path.Combine(temp.Path, "jira-cache", "README.md")));
    }

    private static JiraRestTestCaseSource CreateSource(
        string root,
        FakeHttpMessageHandler handler,
        string mappingProfile = "plain-fields",
        bool allowLocalHttp = false)
    {
        var opts = CreateJiraOptions(allowLocalHttp, mappingProfile);
        return CreateSource(root, handler, opts);
    }

    private static JiraRestTestCaseSource CreateSource(
        string root,
        FakeHttpMessageHandler handler,
        JiraOptions jiraOptions)
    {
        var agent = TestSupport.CreateOptions(workspaceRoot: root);
        agent.Jira = jiraOptions;
        return new JiraRestTestCaseSource(
            TestSupport.CreateFactory(handler),
            agent,
            root,
            NullLogger<JiraRestTestCaseSource>.Instance);
    }

    private static JiraOptions CreateJiraOptions(bool allowLocalHttp, string mappingProfile = "plain-fields") =>
        new()
        {
            BaseUrl = allowLocalHttp ? "http://127.0.0.1:9" : "https://example.atlassian.net",
            AuthMode = "basic",
            Email = "user@example.com",
            ApiToken = "test-token-not-real",
            MappingProfile = mappingProfile,
            AllowLocalHttpTestServer = allowLocalHttp,
            RequestTimeoutSeconds = 5,
            MaxRetries = 1,
            MaxResponseBytes = 1_048_576,
            PageSize = 50,
            MaxPages = 20
        };

    private static JiraSuiteOrchestrator CreateOrchestrator(string root, IDeterministicPlanRunner runner)
    {
        var options = TestSupport.CreateOptions(workspaceRoot: root);
        var workspace = new WorkspaceManager(Options.Create(options), NullLogger<WorkspaceManager>.Instance);
        return new JiraSuiteOrchestrator(
            Options.Create(options),
            workspace,
            new StubJiraSource(),
            new BddCompilationService(workspace, new ObjectRepositoryReader(Options.Create(options), workspace), Options.Create(options)),
            new SuiteManifestReader(Options.Create(options), workspace),
            runner,
            NullLogger<JiraSuiteOrchestrator>.Instance);
    }

    private static void SeedWorkspace(string root)
    {
        var workspace = new WorkspaceManager(
            Options.Create(TestSupport.CreateOptions(workspaceRoot: root)),
            NullLogger<WorkspaceManager>.Instance);
        workspace.Initialize();
        Directory.CreateDirectory(Path.Combine(root, "step-registry"));
        Directory.CreateDirectory(Path.Combine(root, "test-cases", "examples"));
        File.Copy(
            Path.Combine(RepoRoot(), "automation", "step-registry", "registry.json"),
            Path.Combine(root, "step-registry", "registry.json"),
            overwrite: true);
        File.Copy(
            Path.Combine(RepoRoot(), "automation", "test-cases", "examples", "DEMO-1.json"),
            Path.Combine(root, "test-cases", "examples", "DEMO-1.json"),
            overwrite: true);
        File.Copy(
            Path.Combine(RepoRoot(), "automation", "test-cases", "examples", "DEMO-2.json"),
            Path.Combine(root, "test-cases", "examples", "DEMO-2.json"),
            overwrite: true);
    }

    private static void WriteSuite(string root, IReadOnlyList<string> enabledKeys, string disabledKey)
    {
        var cases = enabledKeys.Select(k => $"{{ \"jiraKey\": \"{k}\", \"enabled\": true }}")
            .Concat([$"{{ \"jiraKey\": \"{disabledKey}\", \"enabled\": false }}"]);
        File.WriteAllText(Path.Combine(root, "suites", "phase4e-demo.json"), $$"""
            {
              "schemaVersion": 1,
              "name": "phase4e-demo",
              "enabled": true,
              "testCases": [ {{string.Join(", ", cases)}} ]
            }
            """);
    }

    private static string RepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.sln"))
                || File.Exists(Path.Combine(dir.FullName, "desktopAutomationHelper.sln")))
                return dir.FullName;
            if (Directory.Exists(Path.Combine(dir.FullName, "automation"))
                && Directory.Exists(Path.Combine(dir.FullName, "src")))
                return dir.FullName;
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Could not locate repo root.");
    }

    private static TempFolder TempDir() => new();

    private sealed class TempFolder : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "da-phase4e-" + Guid.NewGuid().ToString("N"));

        public TempFolder() => Directory.CreateDirectory(Path);

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

    private sealed class FakePlanRunner : IDeterministicPlanRunner
    {
        private readonly Func<string, bool, CancellationToken, Task<RunReport>> _impl;
        public FakePlanRunner(Func<string, bool, CancellationToken, Task<RunReport>> impl) => _impl = impl;
        public Task<RunReport> RunAsync(string planPath, bool dryRun, CancellationToken cancellationToken = default) =>
            _impl(planPath, dryRun, cancellationToken);
    }

    private sealed class StubJiraSource : IJiraTestCaseSource
    {
        public Task<JiraTestCaseFetchResult> FetchAsync(string jiraKey, CancellationToken cancellationToken = default) =>
            Task.FromResult(new JiraTestCaseFetchResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                Errors = ["stub source — tests should not call live Jira."]
            });

        public Task<JiraKeySearchResult> SearchKeysAsync(string jql, CancellationToken cancellationToken = default) =>
            Task.FromResult(new JiraKeySearchResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                Errors = ["stub source"]
            });
    }

    /// <summary>Produces headers immediately, then stalls body reads until cancelled.</summary>
    private sealed class SlowHttpContent : HttpContent
    {
        private readonly TimeSpan _stall;

        public SlowHttpContent(TimeSpan stall) => _stall = stall;

        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
            SerializeToStreamAsync(stream, context, CancellationToken.None);

        protected override async Task SerializeToStreamAsync(
            Stream stream,
            TransportContext? context,
            CancellationToken cancellationToken)
        {
            await Task.Delay(_stall, cancellationToken).ConfigureAwait(false);
        }

        protected override Task<Stream> CreateContentReadStreamAsync() =>
            Task.FromResult<Stream>(new SlowReadStream(_stall));

        protected override bool TryComputeLength(out long length)
        {
            length = -1;
            return false;
        }
    }

    private sealed class SlowReadStream : Stream
    {
        private readonly TimeSpan _stall;
        public SlowReadStream(TimeSpan stall) => _stall = stall;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => 0; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException("Use ReadAsync.");
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            await Task.Delay(_stall, cancellationToken).ConfigureAwait(false);
            return 0;
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
