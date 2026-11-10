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
        // Dump order is depth-scrambled with a grandchild under A.
        // Global depth|ordinal sort of root+two children would also pass without the grandchild;
        // DFS via parent links must emit root → A → A1 → B.
        var dumpuia = BuildDumpPayload(
            nodes:
            [
                Node(0, null, 0, 0, "Window", "Main", "main", enabled: true, patterns: ["Window"], hwnd: 99),
                Node(3, 0, 1, 1, "Button", "B", "b", enabled: true, patterns: ["Invoke"]),
                Node(2, 1, 0, 2, "Text", "A1", "a1", enabled: true, patterns: ["Value"], valueIsReadOnly: true),
                Node(1, 0, 0, 1, "Button", "A", "a", enabled: true, patterns: ["Invoke"])
            ]);

        var normalizer = new SemanticSnapshotNormalizer();
        var options = new ObservationOptions { ActionableOnly = false, IncludeFocusState = true };
        var first = normalizer.Normalize(dumpuia, options, "snap-1", DateTimeOffset.Parse("2026-01-01T00:00:00Z"), "1.0.0");
        var second = normalizer.Normalize(dumpuia, options, "snap-2", DateTimeOffset.Parse("2026-01-02T00:00:00Z"), "1.0.0");

        Assert.Equal(["main", "a", "a1", "b"], first.Nodes.Select(n => n.AutomationId));
        Assert.Equal("n0", first.Nodes[1].ParentRef);
        Assert.Equal("n1", first.Nodes[2].ParentRef);
        Assert.Equal(first.Nodes.Select(n => n.AutomationId), second.Nodes.Select(n => n.AutomationId));
        Assert.Equal(first.SnapshotHash, second.SnapshotHash);
        Assert.NotEqual(first.SnapshotId, second.SnapshotId);
    }

    [Fact]
    public void Normalizer_HashIgnoresPidSessionHwndAndBounds()
    {
        // Bounds deliberately differ between payloads so the hash must ignore them.
        var left = BuildDumpPayload(
            [
                Node(0, null, 0, 0, "Window", "Main", "main", true, ["Window"], hwnd: 11, bounds: Rect(10, 20, 100, 50))
            ],
            processId: 100,
            sessionId: 1,
            hwnd: 11,
            includeBounds: true,
            windowBounds: Rect(10, 20, 100, 50));

        var right = BuildDumpPayload(
            [
                Node(0, null, 0, 0, "Window", "Main", "main", true, ["Window"], hwnd: 99, bounds: Rect(999, 888, 10, 10))
            ],
            processId: 200,
            sessionId: 7,
            hwnd: 99,
            includeBounds: true,
            windowBounds: Rect(999, 888, 10, 10));

        var options = new ObservationOptions { ActionableOnly = false, IncludeBounds = true };
        var normalizer = new SemanticSnapshotNormalizer();
        var a = normalizer.Normalize(left, options, "snap-a", DateTimeOffset.UtcNow, "1.0.0");
        var b = normalizer.Normalize(right, options, "snap-b", DateTimeOffset.UtcNow, "1.0.0");

        Assert.NotNull(a.Window!.Bounds);
        Assert.NotNull(b.Window!.Bounds);
        Assert.NotEqual(a.Window.Bounds!.X, b.Window.Bounds!.X);
        Assert.NotEqual(a.Nodes[0].Bounds!.X, b.Nodes[0].Bounds!.X);
        Assert.Equal(a.SnapshotHash, b.SnapshotHash);
        Assert.NotEqual(a.Application.ProcessId, b.Application.ProcessId);
        Assert.NotEqual(a.Window!.Hwnd, b.Window!.Hwnd);
    }

    [Fact]
    public void Normalizer_RejectsMissingSemanticDiagnosticVersion()
    {
        var payload = JsonSerializer.SerializeToElement(new
        {
            operation = "dumpuia",
            success = true,
            nodes = new[]
            {
                new { nodeIndex = 0, ordinal = 0, depth = 0, controlType = "Window", name = "Main", automationId = "main" }
            }
        });

        Assert.Throws<InvalidOperationException>(() =>
            new SemanticSnapshotNormalizer().Normalize(
                payload,
                new ObservationOptions { ActionableOnly = false },
                "snap",
                DateTimeOffset.UtcNow,
                "1.0.0"));
    }

    [Fact]
    public void Normalizer_MapsActiveWindowHwndToMatchingNodeRef()
    {
        var dumpuia = BuildDumpPayload(
            [
                Node(0, null, 0, 0, "Window", "Other", "other", true, ["Window"], hwnd: 10),
                Node(1, null, 1, 0, "Window", "Active", "active", true, ["Window"], hwnd: 42),
                Node(2, 1, 0, 1, "Button", "Go", "go", true, ["Invoke"])
            ],
            hwnd: 42);

        var snapshot = new SemanticSnapshotNormalizer().Normalize(
            dumpuia,
            new ObservationOptions { ActionableOnly = false },
            "snap",
            DateTimeOffset.UtcNow,
            "1.0.0");

        Assert.Equal("active", snapshot.Nodes.Single(n => n.Ref == snapshot.Window!.Ref).AutomationId);
        Assert.Equal(42, snapshot.Window!.Hwnd);
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
        // Fail closed: unknown/omitted valueIsReadOnly must not expose setValue.
        Assert.DoesNotContain("setValue", edit.Capabilities);
        Assert.Contains("readValue", edit.Capabilities);
    }

    [Fact]
    public void Normalizer_ExposesSetValueOnlyWhenWritable()
    {
        var dumpuia = BuildDumpPayload(
        [
            Node(0, null, 0, 0, "Window", "Main", "main", true, ["Window"]),
            Node(1, 0, 0, 1, "Edit", "Writable", "w", true, ["Value"], valueIsReadOnly: false),
            Node(2, 0, 1, 1, "Edit", "Readonly", "r", true, ["Value"], valueIsReadOnly: true)
        ]);

        var snapshot = new SemanticSnapshotNormalizer().Normalize(
            dumpuia,
            new ObservationOptions { ActionableOnly = false },
            "snap",
            DateTimeOffset.UtcNow,
            "1.0.0");

        Assert.Contains("setValue", snapshot.Nodes.Single(n => n.AutomationId == "w").Capabilities);
        Assert.DoesNotContain("setValue", snapshot.Nodes.Single(n => n.AutomationId == "r").Capabilities);
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
            ["valueIsReadOnly"] = false,
            ["supportedPatterns"] = new JsonArray("Value")
        };
        var dumpuia = JsonSerializer.SerializeToElement(new
        {
            operation = "dumpuia",
            success = true,
            semanticDiagnosticVersion = 1,
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
                                nativeWindowHandle = 42,
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
    public async Task ObserveApp_ConcurrentAndSequentialWritesDoNotOverwriteArtifact()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33402",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        const string output = "observations/manual/shared-snapshot.json";
        var absolute = Path.Combine(workspace.RootPath, "observations", "manual", "shared-snapshot.json");
        var observeOptions = new ObservationOptions
        {
            ActionableOnly = false,
            OutputPath = output
        };

        SemanticObservationService CreateService()
        {
            var handler = new FakeHttpMessageHandler(async (req, _) =>
            {
                if (req.Method == HttpMethod.Post)
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
                                    nativeWindowHandle = 42,
                                    isEnabled = true,
                                    isOffscreen = false,
                                    supportedPatterns = new[] { "Window" },
                                    capabilities = new[] { "window" }
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

            var factory = TestSupport.CreateFactory(handler);
            var wrapped = Options.Create(options);
            return new SemanticObservationService(
                wrapped,
                new DriverConnectionResolver(wrapped, factory, NullLogger<DriverConnectionResolver>.Instance),
                new DriverCatalogClient(wrapped, factory, NullLogger<DriverCatalogClient>.Instance),
                new DriverUiClient(wrapped, factory, NullLogger<DriverUiClient>.Instance),
                new SemanticSnapshotNormalizer(),
                workspace,
                NullLogger<SemanticObservationService>.Instance);
        }

        var first = await CreateService().ObserveAsync(observeOptions);
        Assert.True(first.Success);
        Assert.True(File.Exists(absolute));
        var original = await File.ReadAllTextAsync(absolute);

        var second = await CreateService().ObserveAsync(observeOptions);
        Assert.False(second.Success);
        Assert.Contains("already exists", second.Error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(original, await File.ReadAllTextAsync(absolute));

        // Concurrent writers racing the same path: exactly one succeeds; artifact is not clobbered.
        File.Delete(absolute);
        var races = await Task.WhenAll(
            CreateService().ObserveAsync(observeOptions),
            CreateService().ObserveAsync(observeOptions));

        Assert.Equal(1, races.Count(r => r.Success));
        Assert.Equal(1, races.Count(r => !r.Success));
        Assert.Contains(races, r => !r.Success && r.Error is not null
            && r.Error.Contains("already exists", StringComparison.OrdinalIgnoreCase));
        Assert.True(File.Exists(absolute));
        var raced = await File.ReadAllTextAsync(absolute);
        Assert.False(string.IsNullOrWhiteSpace(raced));
        Assert.Contains("snapshotSchemaVersion", raced, StringComparison.Ordinal);
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
        var closed = PatternCapabilityMapper.Map(["Invoke", "Value", "Grid", "Toggle"]);
        Assert.Equal(["grid", "invoke", "readValue", "toggle"], closed);

        var writable = PatternCapabilityMapper.Map(["Value"], valueIsReadOnly: false);
        Assert.Equal(["readValue", "setValue"], writable);
        Assert.True(PatternCapabilityMapper.IsActionable(writable));
    }

    private static JsonElement BuildDumpPayload(
        object[] nodes,
        bool truncated = false,
        int processId = 12,
        int sessionId = 2,
        long hwnd = 99,
        bool includeBounds = false,
        object? windowBounds = null) =>
        JsonSerializer.SerializeToElement(new
        {
            operation = "dumpuia",
            success = true,
            semanticDiagnosticVersion = 1,
            truncated,
            truncationWarnings = truncated ? new[] { "truncated" } : Array.Empty<string>(),
            nodes,
            rootDiagnostics = new
            {
                activeWindowHwnd = hwnd,
                processId,
                processName = "SampleApp",
                windowsSessionId = sessionId,
                activeWindow = new
                {
                    name = "Sample",
                    className = "Wnd",
                    processId,
                    nativeWindowHandle = hwnd,
                    windowState = new { isModal = false },
                    boundingRectangle = includeBounds
                        ? windowBounds ?? Rect(1, 2, 3, 4)
                        : null
                }
            }
        });

    private static object Rect(double x, double y, double width, double height) =>
        new
        {
            x,
            y,
            width,
            height,
            left = x,
            top = y,
            right = x + width,
            bottom = y + height
        };

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
        bool offscreen = false,
        long? hwnd = null,
        bool? valueIsReadOnly = null,
        object? bounds = null) =>
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
            nativeWindowHandle = hwnd,
            isEnabled = enabled,
            isOffscreen = offscreen,
            isKeyboardFocusable = enabled,
            hasKeyboardFocus = false,
            valueIsReadOnly,
            supportedPatterns = patterns,
            capabilities = PatternCapabilityMapper.Map(patterns, valueIsReadOnly),
            boundingRectangle = bounds
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
