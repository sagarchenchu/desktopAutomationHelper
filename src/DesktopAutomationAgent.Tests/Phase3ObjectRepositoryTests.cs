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
using Json.Schema;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class Phase3ObjectRepositoryTests
{
    [Fact]
    public void ObjectRepositorySchemas_ValidateAgainstDraft202012()
    {
        var repoSchemaText = ReadRepoFile("automation/schemas/object-repository.schema.json");
        var pageSchemaText = ReadRepoFile("automation/schemas/page-object.schema.json");

        Assert.Contains("draft/2020-12/schema", repoSchemaText, StringComparison.Ordinal);
        Assert.Contains("draft/2020-12/schema", pageSchemaText, StringComparison.Ordinal);

        var repoSchema = JsonSchema.FromText(repoSchemaText);
        var pageSchema = JsonSchema.FromText(pageSchemaText);
        Assert.NotNull(repoSchema);
        Assert.NotNull(pageSchema);
    }

    [Fact]
    public void WorkspaceManager_TemplateFiles_MatchCheckedInAutomationFiles()
    {
        var workspace = TestSupport.CreateWorkspace(TestSupport.CreateOptions());
        workspace.Initialize();

        var templates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["schemas/object-repository.schema.json"] = "automation/schemas/object-repository.schema.json",
            ["schemas/page-object.schema.json"] = "automation/schemas/page-object.schema.json",
            ["schemas/plan.schema.json"] = "automation/schemas/plan.schema.json",
            ["object-repository/repository.json"] = "automation/object-repository/repository.json",
            ["object-repository/README.md"] = "automation/object-repository/README.md",
            ["object-repository/.gitignore"] = "automation/object-repository/.gitignore"
        };

        foreach (var (relative, source) in templates)
        {
            var generated = NormalizeNewlines(File.ReadAllText(Path.Combine(workspace.RootPath, relative)));
            var expected = NormalizeNewlines(ReadRepoFile(source));
            Assert.Equal(expected, generated);
        }
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("   ", true)]
    [InlineData("btn", false)]
    public void ObjectLocatorValidator_RejectsBlankStringsWhenPresent(string automationId, bool expectError)
    {
        var result = ObjectLocatorValidator.Validate(
            new ObjectLocator { AutomationId = automationId, ControlType = "Button" },
            "test.locator");

        if (expectError)
            Assert.Contains(result.Errors, e => e.Contains("blank", StringComparison.OrdinalIgnoreCase));
        else
            Assert.DoesNotContain(result.Errors, e => e.Contains("blank", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsCandidateManifestPage()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoDir = Path.Combine(workspace.RootPath, "object-repository");
        Directory.CreateDirectory(Path.Combine(repoDir, "pages"));
        File.WriteAllText(Path.Combine(repoDir, "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(repoDir, "pages", "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "candidate",
              "elements": {
                "submit": {
                  "locator": { "automationId": "submit", "controlType": "Button" }
                }
              }
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("manifest-referenced pages must have state 'active'", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsDuplicatePageFilePaths()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoDir = Path.Combine(workspace.RootPath, "object-repository");
        Directory.CreateDirectory(Path.Combine(repoDir, "pages"));
        File.WriteAllText(Path.Combine(repoDir, "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" },
                { "pageId": "home", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(repoDir, "pages", "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": { "locator": { "automationId": "submit" } }
              }
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("duplicates", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RequiresPagesSubdirectory()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoDir = Path.Combine(workspace.RootPath, "object-repository");
        Directory.CreateDirectory(repoDir);
        File.WriteAllText(Path.Combine(repoDir, "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "login.page.json" }
              ]
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("pages/", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PlanObjectReferenceExpander_ReplacesObjectRefInArguments()
    {
        var snapshot = BuildSnapshot();
        var plan = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "test",
            Name = "Test",
            ObjectRepository = "object-repository/repository.json",
            Steps =
            [
                new PlanStep
                {
                    Id = "click-submit",
                    Operation = "click",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["locator"] = JsonSerializer.SerializeToElement(new Dictionary<string, string>
                        {
                            ["$objectRef"] = "login.submit"
                        })
                    }
                }
            ]
        };

        var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
        var result = expander.Expand(plan, snapshot, "plans/test.plan.json");

        Assert.True(result.Success);
        Assert.Equal(["login.submit"], result.ResolvedObjectReferences);
        var locator = plan.Steps![0].Arguments!["locator"];
        Assert.True(locator.TryGetProperty("automationId", out var automationId));
        Assert.Equal("submit", automationId.GetString());
        Assert.False(locator.TryGetProperty("$objectRef", out _));
    }

    [Fact]
    public void PlanObjectReferenceExpander_RequiresObjectRepositoryWhenMarkersPresent()
    {
        var snapshot = BuildSnapshot();
        var plan = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "test",
            Name = "Test",
            Steps =
            [
                new PlanStep
                {
                    Id = "step-1",
                    Operation = "click",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["locator"] = JsonSerializer.SerializeToElement(new Dictionary<string, string>
                        {
                            ["$objectRef"] = "login.submit"
                        })
                    }
                }
            ]
        };

        var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
        var result = expander.Expand(plan, snapshot, "plans/test.plan.json");

        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("objectRepository", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task ValidateObjectRepository_MakesNoHttpCalls()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");

        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var exit = await RunAsync(
            ["validate-object-repository", "--file", repoPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task CapturePage_PerformsStatusOpsThenSinglePost()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");
        var postCount = 0;

        var handler = new FakeHttpMessageHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postCount++;
                var body = await req.Content!.ReadAsStringAsync().ConfigureAwait(false);
                Assert.Contains("dumpuia", body, StringComparison.OrdinalIgnoreCase);
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new
                    {
                        operation = "dumpuia",
                        success = true,
                        nodes = new[]
                        {
                            new
                            {
                                path = "Window/Button[0]",
                                automationId = "submit",
                                controlType = "Button",
                                depth = 1
                            }
                        }
                    }
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

        var exit = await RunAsync(
            [
                "capture-page",
                "--file", repoPath,
                "--page", "login",
                "--name", "Login Page",
                "--json"
            ],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(1, postCount);
        Assert.Contains(handler.Requests, r => r.Method == HttpMethod.Get && r.RequestUri!.AbsolutePath.Contains("/status"));
        Assert.Contains(handler.Requests, r => r.Method == HttpMethod.Get && r.RequestUri!.AbsolutePath.Contains("/ui/operations"));
    }

    [Fact]
    public async Task VerifyObjectRepository_UsesOneFinduiaPerObject()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        SetupActiveRepository(workspace);
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");
        var postCount = 0;

        var handler = new FakeHttpMessageHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postCount++;
                var body = await req.Content!.ReadAsStringAsync().ConfigureAwait(false);
                Assert.Contains("finduia", body, StringComparison.OrdinalIgnoreCase);
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new
                    {
                        operation = "finduia",
                        success = true,
                        found = true,
                        matchCount = 1,
                        matches = new[] { new { automationId = "submit" } }
                    }
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

        var exit = await RunAsync(
            ["verify-object-repository", "--file", repoPath, "--ref", "login.submit", "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(1, postCount);
    }

    [Fact]
    public void ObjectCandidateGenerator_AssignsStrongQualityForUniqueAutomationId()
    {
        var nodes = new List<JsonElement>
        {
            JsonSerializer.SerializeToElement(new
            {
                path = "Window/Button#submit",
                automationId = "submit",
                controlType = "Button"
            })
        };

        var generator = new ObjectCandidateGenerator();
        var page = generator.Generate(nodes, "login", "Login", "capture-1");

        Assert.Equal("candidate", page.State);
        Assert.Single(page.Elements!);
        var element = page.Elements!.Values.First();
        Assert.Equal("strong", element.Quality!.Grade);
        Assert.Equal("capture", element.Source!.Kind);
        Assert.Null(element.Locator!.FoundIndex);
    }

    [Fact]
    public void CommandLine_Parse_RejectsMutuallyExclusiveVerifyFlags()
    {
        var parsed = CommandLine.Parse(
        [
            "verify-object-repository",
            "--file", "object-repository/repository.json",
            "--page", "login",
            "--ref", "login.submit"
        ]);

        Assert.NotNull(parsed.Error);
        Assert.Contains("mutually exclusive", parsed.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("validate-object-repository", "--file")]
    [InlineData("resolve-object", "--file")]
    [InlineData("capture-page", "--file")]
    [InlineData("verify-object-repository", "--file")]
    public void CommandLine_Parse_PreservesJsonOnMissingFlagValues(string command, string flag)
    {
        var parsed = CommandLine.Parse([command, flag, "--json"]);
        Assert.True(parsed.Json);
        Assert.NotNull(parsed.Error);
    }

    [Fact]
    public void PlanObjectReferenceExpander_RejectsMalformedMarkerInLocator()
    {
        var snapshot = BuildSnapshot();
        var plan = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "test",
            Name = "Test",
            ObjectRepository = "object-repository/repository.json",
            Steps =
            [
                new PlanStep
                {
                    Id = "click-submit",
                    Operation = "click",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["locator"] = JsonSerializer.SerializeToElement(new Dictionary<string, string>
                        {
                            ["$objectRef"] = "login.submit",
                            ["extra"] = "nope"
                        })
                    }
                }
            ]
        };

        var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
        var result = expander.Expand(plan, snapshot, "plans/test.plan.json");

        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("exactly one property", StringComparison.OrdinalIgnoreCase));
        Assert.True(plan.Steps![0].Arguments!["locator"].TryGetProperty("$objectRef", out _));
    }

    [Fact]
    public void PlanObjectReferenceExpander_RejectsNestedObjectRefInsideRawLocator()
    {
        var snapshot = BuildSnapshot();
        var plan = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "test",
            Name = "Test",
            ObjectRepository = "object-repository/repository.json",
            Steps =
            [
                new PlanStep
                {
                    Id = "click-submit",
                    Operation = "click",
                    Arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["locator"] = JsonSerializer.SerializeToElement(new Dictionary<string, object>
                        {
                            ["automationId"] = "submit",
                            ["nested"] = new Dictionary<string, string> { ["$objectRef"] = "login.submit" }
                        })
                    }
                }
            ]
        };

        var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
        var result = expander.Expand(plan, snapshot, "plans/test.plan.json");

        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("only allowed", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task ResolveObject_MakesNoHttpCalls()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        SetupActiveRepository(workspace);
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");

        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return new HttpResponseMessage(HttpStatusCode.OK);
        });

        var exit = await RunAsync(
            ["resolve-object", "--file", repoPath, "--ref", "login.submit", "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(0, calls);
        Assert.Contains("\"success\": true", exit.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RunPlan_ExpandedObjectRef_DoesNotReachHttpPayload()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        SetupActiveRepository(workspace);

        var planPath = TestSupport.WritePlan(options, "object-ref.plan.json", """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "login.click",
              "name": "Click submit",
              "objectRepository": "object-repository/repository.json",
              "steps": [
                {
                  "id": "click-submit",
                  "operation": "finduia",
                  "arguments": {
                    "locator": { "$objectRef": "login.submit" }
                  }
                }
              ]
            }
            """);

        string? postBody = null;
        var handler = new FakeHttpMessageHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postBody = await req.Content!.ReadAsStringAsync();
                return FakeHttpMessageHandler.Json(new { success = true, value = new { ok = true } });
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

        var exit = await RunAsync(
            ["run-plan", "--file", planPath, "--json"],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.NotNull(postBody);
        Assert.DoesNotContain("$objectRef", postBody, StringComparison.Ordinal);
        Assert.Contains("\"automationId\":\"submit\"", postBody, StringComparison.Ordinal);
        Assert.Contains("\"objectRepositorySha256\"", exit.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CapturePage_TimeoutCreatesNoCandidate()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");

        var handler = new FakeHttpMessageHandler(req =>
        {
            if (req.Method == HttpMethod.Post)
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = false,
                    value = new
                    {
                        operation = "dumpuia",
                        success = false,
                        reason = "timeout",
                        partialResults = Array.Empty<object>(),
                        nodes = Array.Empty<object>()
                    }
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

        var exit = await RunAsync(
            [
                "capture-page",
                "--file", repoPath,
                "--page", "login",
                "--name", "Login Page",
                "--json"
            ],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.ExecutionFailure, exit.ExitCode);
        Assert.Empty(Directory.GetFiles(
            Path.Combine(workspace.RootPath, "object-repository", "candidates"),
            "*",
            SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(
            Path.Combine(workspace.RootPath, "object-repository", "captures"),
            "*",
            SearchOption.AllDirectories));
    }

    [Fact]
    public async Task VerifyObjectRepository_FoundIndexUsesMatchCountWithoutIndex()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var pagesDir = Path.Combine(workspace.RootPath, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(pagesDir, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "locator": { "automationId": "submit", "controlType": "Button", "foundIndex": 0 },
                  "quality": { "grade": "weak", "warnings": [] },
                  "source": { "kind": "manual" }
                }
              }
            }
            """);

        string? postBody = null;
        var handler = new FakeHttpMessageHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postBody = await req.Content!.ReadAsStringAsync();
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new
                    {
                        operation = "finduia",
                        success = true,
                        found = true,
                        matchCount = 3,
                        matches = new[]
                        {
                            new { automationId = "submit" },
                            new { automationId = "submit" },
                            new { automationId = "submit" }
                        }
                    }
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

        var exit = await RunAsync(
            [
                "verify-object-repository",
                "--file", Path.Combine(workspace.RootPath, "object-repository", "repository.json"),
                "--ref", "login.submit",
                "--json"
            ],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.NotNull(postBody);
        Assert.DoesNotContain("foundIndex", postBody, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"fragile\": 1", exit.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsActivePageWithUnresolved()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var pagesDir = Path.Combine(workspace.RootPath, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(pagesDir, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "locator": { "automationId": "submit", "controlType": "Button" },
                  "source": { "kind": "manual" }
                }
              },
              "unresolved": [ { "path": "Window/Static" } ]
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("unresolved", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WorkspaceManager_AgentSettingsExample_IncludesObjectRepository()
    {
        var workspace = TestSupport.CreateWorkspace(TestSupport.CreateOptions());
        workspace.Initialize();
        var generated = NormalizeNewlines(
            File.ReadAllText(Path.Combine(workspace.RootPath, "config", "agentsettings.example.json")));
        var expected = NormalizeNewlines(ReadRepoFile("automation/config/agentsettings.example.json"));
        Assert.Equal(expected, generated);
        Assert.Contains("ObjectRepository", generated, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("123Submit", "e-123submit")]
    [InlineData("Submit", "submit")]
    public void ObjectCandidateGenerator_ElementIds_AreSchemaValid(string automationId, string expectedPrefix)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        var id = ObjectCandidateGenerator.ResolveElementId(
            new ObjectCandidateGenerator.DumpNode(null, automationId, null, "Button", null, 0),
            used);

        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", id);
        Assert.StartsWith(expectedPrefix, id, StringComparison.Ordinal);
        Assert.True(id.Length <= 64);
    }

    [Fact]
    public void ObjectCandidateGenerator_ElementIds_TruncateLongIdsAndKeepCollisionsWithinLimit()
    {
        var longId = new string('a', 80) + "Button";
        var used = new HashSet<string>(StringComparer.Ordinal);
        var first = ObjectCandidateGenerator.ResolveElementId(
            new ObjectCandidateGenerator.DumpNode(null, longId, null, "Button", null, 0),
            used);
        var second = ObjectCandidateGenerator.ResolveElementId(
            new ObjectCandidateGenerator.DumpNode(null, longId, null, "Button", null, 0),
            used);

        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", first);
        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", second);
        Assert.True(first.Length <= 64);
        Assert.True(second.Length <= 64);
        Assert.NotEqual(first, second);
    }

    [Fact]
    public void ObjectCandidateGenerator_TreatsCaseVariantsAsSameAutomationId()
    {
        var nodes = new List<JsonElement>
        {
            JsonSerializer.SerializeToElement(new { path = "a", automationId = "Submit", controlType = "Button" }),
            JsonSerializer.SerializeToElement(new { path = "b", automationId = "submit", controlType = "Button" })
        };

        var page = new ObjectCandidateGenerator().Generate(nodes, "login", "Login", "cap-1");
        Assert.Empty(page.Elements!);
        Assert.True(page.Unresolved is JsonElement unresolved && unresolved.GetArrayLength() == 2);
    }

    [Fact]
    public void ObjectLocatorValidator_RejectsUnknownControlType()
    {
        var result = ObjectLocatorValidator.Validate(
            new ObjectLocator { AutomationId = "x", ControlType = "NotARealType" },
            "test.locator");

        Assert.Contains(result.Errors, e => e.Contains("recognized", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsWrongPropertyCasing()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "RepositoryId": "default",
              "name": "Test",
              "pages": []
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e =>
            e.Contains("repositoryId", StringComparison.OrdinalIgnoreCase)
            || e.Contains("unknown", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsNullPageEntries()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [null]
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("must not be null", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsNullElementDefinitions()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var pagesDir = Path.Combine(workspace.RootPath, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [ { "pageId": "login", "file": "pages/login.page.json" } ]
            }
            """);
        File.WriteAllText(Path.Combine(pagesDir, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": { "submit": null }
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("must not be null", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsUnknownNestedProperties()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var pagesDir = Path.Combine(workspace.RootPath, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [ { "pageId": "login", "file": "pages/login.page.json", "extra": true } ]
            }
            """);
        File.WriteAllText(Path.Combine(pagesDir, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "locator": { "automationId": "submit", "controlType": "Button" },
                  "quality": { "grade": "strong", "unexpected": 1 },
                  "source": { "kind": "manual", "bogus": "x" }
                }
              }
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("unknown property 'extra'", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Errors, e => e.Contains("unexpected", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Errors, e => e.Contains("bogus", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectRepositoryReader_RejectsNullRequiredString()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": null,
              "pages": []
            }
            """);

        var reader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var result = reader.Read("object-repository/repository.json");

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, e => e.Contains("name is required", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ObjectCandidateGenerator_GeneratedPagePassesCandidateValidation()
    {
        var nodes = new List<JsonElement>
        {
            JsonSerializer.SerializeToElement(new
            {
                path = "Window/Button",
                automationId = "123Submit",
                controlType = "Button"
            })
        };

        var page = new ObjectCandidateGenerator().Generate(nodes, "login", "Login", "cap-1");
        var validation = new ObjectRepositoryValidator().ValidateCandidatePage(
            page,
            "candidates/login/cap-1.page.json",
            new ObjectRepositoryOptions());

        Assert.True(validation.IsValid, string.Join("; ", validation.Errors));
        Assert.All(page.Elements!.Keys, id => Assert.Matches("^[a-z][a-z0-9-]{0,63}$", id));
    }

    [Fact]
    public void CommandLine_Parse_RejectsInvalidView()
    {
        var parsed = CommandLine.Parse(
        [
            "capture-page",
            "--file", "object-repository/repository.json",
            "--page", "login",
            "--name", "Login",
            "--view", "processWindow",
            "--json"
        ]);

        Assert.NotNull(parsed.Error);
        Assert.Contains("--view", parsed.Error, StringComparison.OrdinalIgnoreCase);
        Assert.True(parsed.Json);
    }

    [Fact]
    public void CommandLine_Parse_RejectsCaptureOnlyFlagsOnValidate()
    {
        var parsed = CommandLine.Parse(
        [
            "validate-object-repository",
            "--file", "object-repository/repository.json",
            "--view", "control",
            "--json"
        ]);

        Assert.NotNull(parsed.Error);
        Assert.Contains("--view", parsed.Error, StringComparison.OrdinalIgnoreCase);
        Assert.True(parsed.Json);
    }

    [Fact]
    public async Task CapturePage_RejectsInvalidRootWithUsageExit()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoPath = Path.Combine(workspace.RootPath, "object-repository", "repository.json");
        var handler = new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));

        var exit = await RunAsync(
            [
                "capture-page",
                "--file", repoPath,
                "--page", "login",
                "--name", "Login",
                "--root", "processWindow",
                "--json"
            ],
            options,
            workspace,
            handler);

        Assert.Equal(ExitCodes.UsageOrConfiguration, exit.ExitCode);
        Assert.Equal(0, handler.Requests.Count);
    }

    [Fact]
    public void ObjectRepositoryPathSafety_RejectsSymlinkEscape_WhenSupported()
    {
        if (OperatingSystem.IsWindows())
            return;

        var root = Path.Combine(Path.GetTempPath(), "da-or-symlink-" + Guid.NewGuid().ToString("N"));
        var outside = Path.Combine(Path.GetTempPath(), "da-or-outside-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), "{}");
        var linkPath = Path.Combine(root, "pages");

        try
        {
            File.CreateSymbolicLink(linkPath, outside);
        }
        catch (Exception ex) when (ex is IOException or PlatformNotSupportedException or UnauthorizedAccessException)
        {
            return;
        }

        try
        {
            var target = Path.Combine(linkPath, "secret.json");
            var thrown = Assert.Throws<RepositoryPathException>(() =>
                ObjectRepositoryPathSafety.EnsureNotSymlinkEscape(target, root));
            Assert.Contains("symbolic link", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            try { Directory.Delete(outside, recursive: true); } catch { /* ignore */ }
        }
    }

    private static string NormalizeNewlines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal);

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

        throw new FileNotFoundException($"Unable to locate repository file '{relativePath}'.");
    }

    private static void SetupActiveRepository(WorkspaceManager workspace)
    {
        var pagesDir = Path.Combine(workspace.RootPath, "object-repository", "pages");
        Directory.CreateDirectory(pagesDir);
        File.WriteAllText(Path.Combine(workspace.RootPath, "object-repository", "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "default",
              "name": "Test",
              "pages": [
                { "pageId": "login", "file": "pages/login.page.json" }
              ]
            }
            """);
        File.WriteAllText(Path.Combine(pagesDir, "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "locator": { "automationId": "submit", "controlType": "Button" },
                  "source": { "kind": "manual" }
                }
              }
            }
            """);
    }

    private static ObjectRepositorySnapshot BuildSnapshot()
    {
        var manifest = new ObjectRepositoryManifest
        {
            SchemaVersion = 1,
            RepositoryId = "default",
            Name = "Test",
            Pages =
            [
                new PageReference { PageId = "login", File = "pages/login.page.json" }
            ]
        };

        var page = new PageObjectDocument
        {
            SchemaVersion = 1,
            PageId = "login",
            Name = "Login",
            State = "active",
            Elements = new Dictionary<string, ObjectElementDefinition>(StringComparer.Ordinal)
            {
                ["submit"] = new()
                {
                    Locator = new ObjectLocator { AutomationId = "submit", ControlType = "Button" },
                    Source = new ObjectSource { Kind = "manual" }
                }
            }
        };

        return new ObjectRepositorySnapshot(
            manifest,
            new Dictionary<string, PageObjectDocument>(StringComparer.Ordinal) { ["login"] = page },
            new Dictionary<string, string>(StringComparer.Ordinal) { ["login"] = "object-repository/pages/login.page.json" },
            new Dictionary<string, string>(StringComparer.Ordinal),
            "object-repository/repository.json",
            "abc",
            "def");
    }

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
}
