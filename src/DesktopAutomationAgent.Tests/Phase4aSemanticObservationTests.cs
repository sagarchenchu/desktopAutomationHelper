using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Workspace;
using Json.Schema;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class Phase4aSemanticObservationTests
{
    [Fact]
    public void SemanticSnapshotSchema_ValidatesSampleAndDraft202012()
    {
        var schemaText = ReadRepoFile("automation/schemas/semantic-snapshot.schema.json");
        Assert.Contains("draft/2020-12/schema", schemaText, StringComparison.Ordinal);

        var schema = JsonSchema.FromText(schemaText);
        var sample = JsonNode.Parse(ReadRepoFile("automation/observations/examples/sample-semantic-snapshot.json"));
        var result = schema.Evaluate(sample, new EvaluationOptions { OutputFormat = OutputFormat.List });
        Assert.True(result.IsValid, string.Join("; ", result.Details.Select(d => d.Errors is null ? d.InstanceLocation.ToString() : string.Join(",", d.Errors.Values))));
    }

    [Fact]
    public void WorkspaceManager_IncludesObservationsTemplates()
    {
        var workspace = TestSupport.CreateWorkspace(TestSupport.CreateOptions());
        var result = workspace.Initialize();

        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "observations/");
        Assert.Contains(result.CreatedPaths, p => p.Replace('\\', '/') == "schemas/semantic-snapshot.schema.json");

        var generated = NormalizeNewlines(File.ReadAllText(Path.Combine(workspace.RootPath, "schemas", "semantic-snapshot.schema.json")));
        var expected = NormalizeNewlines(ReadRepoFile("automation/schemas/semantic-snapshot.schema.json"));
        Assert.Equal(expected, generated);
    }

    [Fact]
    public void Normalizer_ProducesStableOrderingAndHash()
    {
        var dumpuia = BuildDumpPayload(
            nodes:
            [
                Node(0, null, 0, 0, "Window", "Main", "main", enabled: true, patterns: ["Window"]),
                Node(2, 0, 1, 1, "Button", "B", "b", enabled: true, patterns: ["Invoke"]),
                Node(1, 0, 0, 1, "Button", "A", "a", enabled: true, patterns: ["Invoke"])
            ]);

        var normalizer = new SemanticSnapshotNormalizer();
        var options = new ObservationOptions { ActionableOnly = false, IncludeFocusState = true };
        var first = normalizer.Normalize(dumpuia, options, "snap-1", DateTimeOffset.Parse("2026-01-01T00:00:00Z"), "1.0.0");
        var second = normalizer.Normalize(dumpuia, options, "snap-2", DateTimeOffset.Parse("2026-01-02T00:00:00Z"), "1.0.0");

        Assert.Equal(["n0", "n1", "n2"], first.Nodes.Select(n => n.Ref));
        Assert.Equal("n0", first.Nodes[1].ParentRef);
        Assert.Equal(first.Nodes.Select(n => n.AutomationId), second.Nodes.Select(n => n.AutomationId));
        Assert.Equal(first.SnapshotHash, second.SnapshotHash);
        Assert.NotEqual(first.SnapshotId, second.SnapshotId);
    }

    [Fact]
    public void Normalizer_RespectsMaxNodesAndEmitsTruncationWarning()
    {
        var nodes = Enumerable.Range(0, 5)
            .Select(i => Node(i, i == 0 ? null : 0, i, i == 0 ? 0 : 1, i == 0 ? "Window" : "Button", $"N{i}", $"id-{i}", true, ["Invoke"]))
            .ToArray();
        var dumpuia = BuildDumpPayload(nodes, truncated: false);
        var snapshot = new SemanticSnapshotNormalizer().Normalize(
            dumpuia,
            new ObservationOptions { MaxNodes = 3, ActionableOnly = false },
            "snap",
            DateTimeOffset.UtcNow,
            "1.0.0");

        Assert.True(snapshot.Truncated);
        Assert.Equal(3, snapshot.Nodes.Count);
        Assert.Contains(snapshot.Warnings, w => w.Contains("maxNodes", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Normalizer_MapsPatternsToCapabilities_AndPreservesDisabledOffscreen()
    {
        var dumpuia = BuildDumpPayload(
        [
            Node(0, null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(1, 0, 0, 1, "Edit", "Name", "name", enabled: false, patterns: ["Value"], offscreen: true)
        ]);

        var snapshot = new SemanticSnapshotNormalizer().Normalize(
            dumpuia,
            new ObservationOptions { ActionableOnly = false, IncludeFocusState = true },
            "snap",
            DateTimeOffset.UtcNow,
            "1.0.0");

        var edit = snapshot.Nodes.Single(n => n.AutomationId == "name");
        Assert.False(edit.Enabled);
        Assert.True(edit.Offscreen);
        Assert.Contains("setValue", edit.Capabilities);
        Assert.Contains("readValue", edit.Capabilities);
        Assert.DoesNotContain(edit.Capabilities, c => c.Contains("value", StringComparison.OrdinalIgnoreCase) && c is not ("setValue" or "readValue"));
    }

    [Fact]
    public void Normalizer_OmitsEditableValuesByDefault()
    {
        var node = new JsonObject
        {
            ["nodeIndex"] = 0,
            ["parentNodeIndex"] = null,
            ["ordinal"] = 0,
            ["depth"] = 0,
            ["controlType"] = "Edit",
            ["name"] = "Secret Field",
            ["automationId"] = "secret",
            ["isEnabled"] = true,
            ["isOffscreen"] = false,
            ["value"] = "super-secret-password",
            ["supportedPatterns"] = new JsonArray("Value")
        };
        var dumpuia = JsonSerializer.SerializeToElement(new
        {
            operation = "dumpuia",
            success = true,
            truncated = false,
            nodes = new[] { node },
            rootDiagnostics = new { processId = 1, processName = "Sample", windowsSessionId = 1 }
        });

        var json = JsonSerializer.Serialize(
            new SemanticSnapshotNormalizer().Normalize(
                dumpuia,
                new ObservationOptions { ActionableOnly = false },
                "snap",
                DateTimeOffset.UtcNow,
                "1.0.0"));

        Assert.DoesNotContain("super-secret-password", json, StringComparison.Ordinal);
        Assert.DoesNotContain("\"value\":", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Normalizer_AppliesSecretRedaction()
    {
        var dumpuia = BuildDumpPayload(
        [
            Node(0, null, 0, 0, "Window", "Bearer abcdefghijklmnop", "main", true, ["Window"])
        ]);

        var snapshot = new SemanticSnapshotNormalizer().Normalize(
            dumpuia,
            new ObservationOptions { ActionableOnly = false },
            "snap",
            DateTimeOffset.UtcNow,
            "1.0.0");

        Assert.Contains("[REDACTED]", snapshot.Nodes[0].Name, StringComparison.Ordinal);
    }

    [Fact]
    public void EphemeralRefIndex_RejectsDuplicateUnknownAndStaleRefs()
    {
        var snapshot = new SemanticSnapshot
        {
            SnapshotId = "snap-a",
            Nodes =
            [
                new SemanticNode { Ref = "n0", Ordinal = 0, Depth = 0, ControlType = "Window", Role = "Window" }
            ]
        };

        var index = new EphemeralRefIndex(snapshot);
        Assert.Equal("n0", index.Resolve("snap-a", "n0").Ref);
        Assert.Throws<StaleEphemeralRefException>(() => index.Resolve("snap-a", "n9"));
        Assert.Throws<StaleEphemeralRefException>(() => index.Resolve("snap-b", "n0"));

        var duplicate = new SemanticSnapshot
        {
            SnapshotId = "snap-dup",
            Nodes =
            [
                new SemanticNode { Ref = "n0", Ordinal = 0, Depth = 0, ControlType = "Window", Role = "Window" },
                new SemanticNode { Ref = "n0", Ordinal = 1, Depth = 1, ControlType = "Button", Role = "Button" }
            ]
        };
        Assert.Throws<InvalidOperationException>(() => new EphemeralRefIndex(duplicate));
    }

    [Fact]
    public async Task ObserveApp_UsesSingleDumpUiaPost_AndWritesArtifact()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33401",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var postCount = 0;
        string? postBody = null;

        var handler = new FakeHttpMessageHandler(async (req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                postCount++;
                postBody = await req.Content!.ReadAsStringAsync().ConfigureAwait(false);
                Assert.Contains("dumpuia", postBody, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("\"operation\":\"click\"", postBody, StringComparison.OrdinalIgnoreCase);
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new
                    {
                        operation = "dumpuia",
                        success = true,
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
                                nativeWindowHandle = 42,
                                windowState = new { isModal = false }
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
                                isKeyboardFocusable = true,
                                hasKeyboardFocus = false,
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

        var exit = await RunAsync(["observe-app", "--json", "--max-nodes", "50"], options, workspace, handler);
        Assert.Equal(ExitCodes.Success, exit.ExitCode);
        Assert.Equal(1, postCount);
        Assert.NotNull(postBody);
        Assert.Contains("\"includeTreeLinks\":true", postBody, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"maxNodes\":50", postBody, StringComparison.OrdinalIgnoreCase);

        using var doc = JsonDocument.Parse(exit.Stdout);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(2, doc.RootElement.GetProperty("nodeCount").GetInt32());
        var artifact = doc.RootElement.GetProperty("snapshotFilePath").GetString();
        Assert.False(string.IsNullOrWhiteSpace(artifact));
        Assert.True(File.Exists(Path.Combine(workspace.RootPath, artifact!.Replace('/', Path.DirectorySeparatorChar))));
        Assert.All(handler.Requests.Where(r => r.Method == HttpMethod.Post), r =>
        {
            Assert.NotNull(r.Content);
        });
        Assert.Contains("dumpuia", postBody, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\"operation\":\"click\"", postBody, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ObserveApp_RejectsOversizedDriverResponse()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33402",
            bearerToken: "secret-token");
        options.Runner.MaxResponseBytes = 64;
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var huge = new string('x', 128);
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            if (req.Method == HttpMethod.Post)
            {
                return Task.FromResult(FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new { operation = "dumpuia", success = true, nodes = new[] { new { name = huge } } }
                }));
            }

            if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
            {
                return Task.FromResult(FakeHttpMessageHandler.Json(new
                {
                    status = 0,
                    value = new { ready = true, message = "ok", build = new { version = "1.0.105" } }
                }));
            }

            return Task.FromResult(FakeHttpMessageHandler.Json(new { success = true, value = CatalogFixtures.Phase2Catalog() }));
        });

        var exit = await RunAsync(["observe-app", "--json"], options, workspace, handler);
        Assert.Equal(ExitCodes.ExecutionFailure, exit.ExitCode);
        Assert.Contains("maximum size", exit.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CommandLine_Parse_ObserveAppDefaultsAndHelp()
    {
        var parsed = CommandLine.Parse(["observe-app", "--json", "--include-bounds", "--max-nodes", "10"]);
        Assert.Equal(AgentCommandKind.ObserveApp, parsed.Kind);
        Assert.True(parsed.Json);
        Assert.True(parsed.IncludeBounds);
        Assert.Equal(10, parsed.MaxNodes);
        Assert.Contains("observe-app", CommandLine.HelpText, StringComparison.Ordinal);
        Assert.Contains("Crawling and AI are not implemented", CommandLine.HelpText, StringComparison.Ordinal);
    }

    [Fact]
    public void PatternCapabilityMapper_MapsKnownPatterns()
    {
        var caps = PatternCapabilityMapper.Map(["Invoke", "Value", "Grid", "Toggle"]);
        Assert.Equal(["grid", "invoke", "readValue", "setValue", "toggle"], caps);
        Assert.True(PatternCapabilityMapper.IsActionable(caps));
    }

    private static JsonElement BuildDumpPayload(object[] nodes, bool truncated = false) =>
        JsonSerializer.SerializeToElement(new
        {
            operation = "dumpuia",
            success = true,
            truncated,
            truncationWarnings = truncated ? new[] { "truncated" } : Array.Empty<string>(),
            nodes,
            rootDiagnostics = new
            {
                activeWindowHwnd = 99,
                processId = 12,
                processName = "SampleApp",
                windowsSessionId = 2,
                activeWindow = new
                {
                    name = "Sample",
                    className = "Wnd",
                    processId = 12,
                    nativeWindowHandle = 99,
                    windowState = new { isModal = false }
                }
            }
        });

    private static object Node(
        int nodeIndex,
        int? parentNodeIndex,
        int ordinal,
        int depth,
        string controlType,
        string name,
        string automationId,
        bool enabled,
        string[] patterns,
        bool offscreen = false) =>
        new
        {
            nodeIndex,
            parentNodeIndex,
            ordinal,
            depth,
            controlType,
            name,
            automationId,
            className = "Cls",
            isEnabled = enabled,
            isOffscreen = offscreen,
            isKeyboardFocusable = enabled,
            hasKeyboardFocus = false,
            supportedPatterns = patterns,
            capabilities = PatternCapabilityMapper.Map(patterns)
        };

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunAsync(
        string[] args,
        AgentOptions options,
        IWorkspaceManager workspace,
        FakeHttpMessageHandler handler)
    {
        var factory = TestSupport.CreateFactory(handler);
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var originalOut = Console.Out;
        var originalErr = Console.Error;
        Console.SetOut(stdout);
        Console.SetError(stderr);

        try
        {
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
                builder.Services.AddSingleton<IWorkspaceManager>(workspace);
                builder.Services.AddSingleton<IDriverConnectionResolver>(_ =>
                    new DriverConnectionResolver(Options.Create(options), factory, NullLogger<DriverConnectionResolver>.Instance));
                builder.Services.AddSingleton<IDriverCatalogClient>(_ =>
                    new DriverCatalogClient(Options.Create(options), factory, NullLogger<DriverCatalogClient>.Instance));
                builder.Services.AddSingleton<IDriverUiClient>(_ =>
                    new DriverUiClient(Options.Create(options), factory, NullLogger<DriverUiClient>.Instance));
                builder.Services.AddSingleton<SemanticSnapshotNormalizer>();
                builder.Services.AddSingleton<ISemanticObservationService, SemanticObservationService>();
                return builder.Build();
            }

            var code = await AgentCli.RunAsync(args, HostBuilder).ConfigureAwait(false);
            return (code, stdout.ToString(), stderr.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    private static string ReadRepoFile(string relativePath)
    {
        var full = Path.GetFullPath(Path.Combine(FindRepoRoot(), relativePath));
        return File.ReadAllText(full);
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomation.sln"))
                || File.Exists(Path.Combine(dir.FullName, "automation", "schemas", "plan.schema.json")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException("Unable to locate repository root.");
    }

    private static string NormalizeNewlines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd() + "\n";
}
