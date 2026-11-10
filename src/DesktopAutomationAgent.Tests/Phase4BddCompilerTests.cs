using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class Phase4BddCompilerTests
{
    [Fact]
    public void Compile_ProducesDeterministicPlanFromMapAndRecording()
    {
        var (options, root, mapRelative, recordingRelative) = CreateAssistiveFixture();
        var workspace = TestSupport.CreateWorkspace(options);
        var compiler = new DeterministicBddPlanCompiler(workspace, NullLogger<DeterministicBddPlanCompiler>.Instance);

        var first = compiler.Compile(new BddCompileRequest
        {
            MapPath = mapRelative,
            RecordingPath = recordingRelative,
            OutputPath = "plans/abc-1234.compiled.plan.json"
        });
        Assert.True(first.Success, string.Join("; ", first.Errors));
        Assert.Equal(2, first.StepCount);
        Assert.True(File.Exists(first.OutputPath));

        var firstBytes = File.ReadAllBytes(first.OutputPath!);

        var second = compiler.Compile(new BddCompileRequest
        {
            MapPath = mapRelative,
            RecordingPath = recordingRelative,
            OutputPath = "plans/abc-1234.compiled.plan.json"
        });
        Assert.True(second.Success, string.Join("; ", second.Errors));
        Assert.Equal(firstBytes, File.ReadAllBytes(second.OutputPath!));

        using var doc = JsonDocument.Parse(File.ReadAllText(first.OutputPath!));
        var rootEl = doc.RootElement;
        Assert.Equal(1, rootEl.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(2, rootEl.GetProperty("catalogSchemaVersion").GetInt32());
        Assert.Equal(2, rootEl.GetProperty("steps").GetArrayLength());

        var step0 = rootEl.GetProperty("steps")[0];
        Assert.Equal("click", step0.GetProperty("operation").GetString());
        Assert.Equal("OK", step0.GetProperty("arguments").GetProperty("locator").GetProperty("name").GetString());

        var step1 = rootEl.GetProperty("steps")[1];
        Assert.Equal("type", step1.GetProperty("operation").GetString());
        Assert.Equal("hello", step1.GetProperty("arguments").GetProperty("value").GetString());
        Assert.True(step1.GetProperty("sensitive").GetBoolean());

        // Statement text is metadata only — not used as an operation.
        Assert.Contains("Click the OK button", File.ReadAllText(first.OutputPath!), StringComparison.Ordinal);

        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void Compile_FailsWhenEventMissingFromRecording()
    {
        var (options, root, mapRelative, _) = CreateAssistiveFixture(includeRecordingEvent: false);
        var workspace = TestSupport.CreateWorkspace(options);
        var compiler = new DeterministicBddPlanCompiler(workspace, NullLogger<DeterministicBddPlanCompiler>.Instance);

        var result = compiler.Compile(new BddCompileRequest
        {
            MapPath = mapRelative,
            RecordingPath = "recordings/recording_missing.json"
        });

        Assert.False(result.Success);
        Assert.Contains(result.Errors, e => e.Contains("eventId not found", StringComparison.OrdinalIgnoreCase));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void Compile_DoesNotInterpretBddStatementText()
    {
        var (options, root, mapRelative, recordingRelative) = CreateAssistiveFixture(
            statement: "When I invent an unsupported verb foobar the control");
        var workspace = TestSupport.CreateWorkspace(options);
        var compiler = new DeterministicBddPlanCompiler(workspace, NullLogger<DeterministicBddPlanCompiler>.Instance);

        var result = compiler.Compile(new BddCompileRequest
        {
            MapPath = mapRelative,
            RecordingPath = recordingRelative
        });

        Assert.True(result.Success, string.Join("; ", result.Errors));
        Assert.All(result.Plan!.Steps!, s => Assert.NotEqual("foobar", s.Operation, StringComparer.OrdinalIgnoreCase));
        Assert.Contains(result.Plan.Steps!, s => s.Operation == "click");
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void Compile_ResolvesRecordingBesideAssistiveLayoutWithoutExplicitPath()
    {
        var (options, root, mapRelative, _) = CreateAssistiveFixture();
        var workspace = TestSupport.CreateWorkspace(options);
        var compiler = new DeterministicBddPlanCompiler(workspace, NullLogger<DeterministicBddPlanCompiler>.Instance);

        var result = compiler.Compile(new BddCompileRequest { MapPath = mapRelative });
        Assert.True(result.Success, string.Join("; ", result.Errors));
        Assert.False(string.IsNullOrWhiteSpace(result.RecordingPath));
        Assert.True(File.Exists(result.RecordingPath));
        Directory.Delete(root, recursive: true);
    }

    [Fact]
    public void CommandLine_ParseCompileBdd_AcceptsOptionalRecordingAndOut()
    {
        var parsed = CommandLine.Parse([
            "compile-bdd",
            "--map", "artifacts/bdd-action-map.json",
            "--recording", "recordings/recording.json",
            "--out", "plans/out.plan.json",
            "--json"
        ]);
        Assert.Null(parsed.Error);
        Assert.Equal(AgentCommandKind.CompileBdd, parsed.Kind);
        Assert.Equal("artifacts/bdd-action-map.json", parsed.MapFile);
        Assert.Equal("recordings/recording.json", parsed.RecordingFile);
        Assert.Equal("plans/out.plan.json", parsed.OutputFile);
        Assert.True(parsed.Json);
    }

    [Fact]
    public void CommandLine_ParseCompileBdd_RequiresMap()
    {
        var parsed = CommandLine.Parse(["compile-bdd", "--json"]);
        Assert.Equal(AgentCommandKind.CompileBdd, parsed.Kind);
        Assert.NotNull(parsed.Error);
        Assert.Contains("--map", parsed.Error, StringComparison.Ordinal);
    }

    private static (AgentOptions Options, string Root, string MapRelative, string RecordingRelative)
        CreateAssistiveFixture(
            string statement = "Click the OK button",
            bool includeRecordingEvent = true)
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var recordingDir = Path.Combine(workspace.RootPath, "recordings");
        var artifactDir = Path.Combine(
            recordingDir, "assistive-artifacts", "ABC-1234", "rec-001");
        Directory.CreateDirectory(artifactDir);

        const string recordingFileName = "recording_test.json";
        var recordingPath = Path.Combine(recordingDir, recordingFileName);
        var actionsJson = includeRecordingEvent
            ? """
              [
                {
                  "eventId": "evt-000001",
                  "sequence": 1,
                  "actionType": "Click",
                  "operation": "click",
                  "element": { "name": "OK", "controlType": "Button", "automationId": "OkButton" },
                  "objectRef": "welcome.ok"
                },
                {
                  "eventId": "evt-000002",
                  "sequence": 2,
                  "actionType": "Type",
                  "operation": "type",
                  "value": "hello",
                  "element": { "name": "Name", "controlType": "Edit", "automationId": "NameBox" },
                  "objectRef": "welcome.name"
                }
              ]
              """
            : "[]";

        File.WriteAllText(recordingPath,
            $$"""
            {
              "recordingId": "rec-001",
              "actions": {{actionsJson}}
            }
            """);

        var mapPath = Path.Combine(artifactDir, "bdd-action-map.json");
        File.WriteAllText(mapPath,
            $$"""
            {
              "schemaVersion": 1,
              "recordingId": "rec-001",
              "jiraKey": "ABC-1234",
              "sourceRecording": "{{recordingFileName}}",
              "createdAt": "2026-08-11T00:00:00Z",
              "pages": [
                { "pageId": "welcome", "windowTitle": "Welcome", "file": "page-objects/welcome.page.json" }
              ],
              "bddGroups": [
                {
                  "groupId": "bdd-0001",
                  "statement": "{{statement}}",
                  "actions": [
                    {
                      "eventId": "evt-000001",
                      "sequence": 1,
                      "pageId": "welcome",
                      "objectRef": "welcome.ok",
                      "operation": "click"
                    },
                    {
                      "eventId": "evt-000002",
                      "sequence": 2,
                      "pageId": "welcome",
                      "objectRef": "welcome.name",
                      "operation": "type"
                    }
                  ]
                }
              ],
              "unmappedEventIds": []
            }
            """);

        if (!includeRecordingEvent)
        {
            File.WriteAllText(
                Path.Combine(workspace.RootPath, "recordings", "recording_missing.json"),
                """{ "recordingId": "rec-001", "actions": [] }""");
        }

        var mapRelative = Path.GetRelativePath(workspace.RootPath, mapPath);
        var recordingRelative = Path.GetRelativePath(workspace.RootPath, recordingPath);
        return (options, options.Workspace.Root, mapRelative, recordingRelative);
    }
}
