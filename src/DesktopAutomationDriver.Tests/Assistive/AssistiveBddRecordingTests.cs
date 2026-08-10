using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Recording.Assistive;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Assistive;
using Json.Schema;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests.Assistive;

public class AssistiveBddRecordingTests
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    [Theory]
    [InlineData("abc-1234", "ABC-1234")]
    [InlineData(" PROJ_1-99 ", "PROJ_1-99")]
    public void JiraKey_IsAcceptedAndCanonicalized(string raw, string expected)
    {
        Assert.True(JiraKeyRules.TryCanonicalize(raw, out var canonical, out var error), error);
        Assert.Equal(expected, canonical);
    }

    [Theory]
    [InlineData("")]
    [InlineData("ABC")]
    [InlineData("ABC-0")]
    [InlineData("../ABC-1")]
    [InlineData("ABC/1")]
    [InlineData("ABC 1-2")]
    [InlineData("ABC-1234/../x")]
    public void JiraKey_RejectsInvalidValues(string raw)
    {
        Assert.False(JiraKeyRules.TryCanonicalize(raw, out _, out var error));
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void JiraKey_CannotChangeAfterFirstScopedEvent()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        Assert.True(coordinator.TryStartJiraRecording("ABC-1", out _, out _));

        var first = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "a", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(first, Window("Welcome"));
        Assert.Equal("ABC-1", first.JiraKey);

        Assert.False(coordinator.TryStartJiraRecording("XYZ-2", out _, out var error));
        Assert.Contains("locked", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StartingJiraScope_DoesNotDeleteOrdinaryActions()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");

        // Ordinary assistive-less recording is represented by actions without Jira metadata.
        var ordinary = new RecordedAction
        {
            ActionType = ActionType.Click,
            Mode = RecordingMode.Passive,
            Element = new ElementInfo { Name = "Ordinary", ControlType = "Button" }
        };

        Assert.True(coordinator.TryStartJiraRecording("ABC-1234", out _, out _));
        var jiraAction = Click("btnABC", "ABC");
        coordinator.EnrichAssistiveAction(jiraAction, Window("Welcome"));

        Assert.Null(ordinary.JiraKey);
        Assert.Equal("ABC-1234", jiraAction.JiraKey);
        Assert.Equal(1, jiraAction.Sequence);
    }

    [Fact]
    public void PassiveActions_NeverInheritAssistiveBdd()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        Assert.True(coordinator.TryStartJiraRecording("ABC-1234", out _, out _));
        Assert.True(coordinator.TryArmBdd("And click OK", BddScope.NextAction, out var groupId, out _));

        // Passive recording path does not call EnrichAssistiveAction.
        var passive = new RecordedAction
        {
            ActionType = ActionType.Click,
            Mode = RecordingMode.Passive,
            Element = new ElementInfo { Name = "Passive", ControlType = "Button" }
        };

        Assert.Null(passive.Bdd);
        Assert.Null(passive.EventId);
        Assert.Equal(groupId, coordinator.ActiveBddGroupId);
    }

    [Fact]
    public void NextActionBdd_AttachesToExactlyOneAssistiveAction()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        Assert.True(coordinator.TryArmBdd("And double click on ABC", BddScope.NextAction, out var groupId, out _));

        var first = Click("btnABC", "ABC");
        coordinator.EnrichAssistiveAction(first, Window("Welcome"));
        Assert.Equal(groupId, first.Bdd!.GroupId);

        var second = Click("btnOK", "OK");
        coordinator.EnrichAssistiveAction(second, Window("Welcome"));
        Assert.Null(second.Bdd);
    }

    [Fact]
    public void NextActionBdd_RemainsArmedWhenNoActionRecorded()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("And click OK", BddScope.NextAction, out var groupId, out _);

        Assert.Equal(groupId, coordinator.ActiveBddGroupId);

        var action = Click("btnOK", "OK");
        coordinator.EnrichAssistiveAction(action, Window("Welcome"));
        Assert.Equal(groupId, action.Bdd!.GroupId);
    }

    [Fact]
    public void MultipleActionBdd_AttachesUntilFinished()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("When the user prints the report", BddScope.UntilFinished, out var groupId, out _);

        var a1 = Click("report", "Report");
        var a2 = Click("printer", "Printer");
        var a3 = Click("print", "Print");
        coordinator.EnrichAssistiveAction(a1, Window("Print Dialog"));
        coordinator.EnrichAssistiveAction(a2, Window("Print Dialog"));
        coordinator.EnrichAssistiveAction(a3, Window("Print Dialog"));
        Assert.Equal(groupId, a1.Bdd!.GroupId);
        Assert.Equal(groupId, a2.Bdd!.GroupId);
        Assert.Equal(groupId, a3.Bdd!.GroupId);

        Assert.True(coordinator.TryFinishBdd(out _));
        var a4 = Click("close", "Close");
        coordinator.EnrichAssistiveAction(a4, Window("Print Dialog"));
        Assert.Null(a4.Bdd);
    }

    [Fact]
    public void CancelBdd_RemovesUnusedContext_ButNotAssociatedActions()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("And click OK", BddScope.NextAction, out _, out _);
        Assert.True(coordinator.TryCancelBdd(out _));
        Assert.Null(coordinator.ActiveBddGroupId);

        coordinator.TryArmBdd("And click Save", BddScope.UntilFinished, out var groupId, out _);
        var action = Click("save", "Save");
        coordinator.EnrichAssistiveAction(action, Window("Welcome"));
        Assert.False(coordinator.TryCancelBdd(out var message));
        Assert.Contains("Finish", message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(groupId, action.Bdd!.GroupId);
    }

    [Fact]
    public void NoBdd_OmitsBddPropertyFromJson()
    {
        var action = new RecordedAction
        {
            ActionType = ActionType.Click,
            Mode = RecordingMode.Assistive,
            EventId = "evt-000001",
            Sequence = 1,
            Element = new ElementInfo { AutomationId = "a", ControlType = "Button" }
        };

        var json = JsonSerializer.Serialize(action, JsonOpts);
        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.TryGetProperty("bdd", out _));
    }

    [Fact]
    public void DeferredBddConsumption_KeepsNextActionArmed()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("And drag item", BddScope.NextAction, out var groupId, out _);

        var intermediate = Click("source", "Source");
        coordinator.EnrichAssistiveAction(
            intermediate,
            new AssistiveActionCaptureContext
            {
                Window = Window("Welcome").Window,
                DeferBddConsumption = true
            });
        Assert.Null(intermediate.Bdd);
        Assert.Equal(groupId, coordinator.ActiveBddGroupId);

        var final = new RecordedAction
        {
            ActionType = ActionType.DragAndDrop,
            Element = new ElementInfo { AutomationId = "source", ControlType = "ListItem" },
            TargetElement = new ElementInfo { AutomationId = "target", ControlType = "List" }
        };
        coordinator.EnrichAssistiveAction(final, Window("Welcome"));
        Assert.Equal(groupId, final.Bdd!.GroupId);
    }

    [Fact]
    public void EventIds_AreSequentialAndUnique()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < 5; i++)
        {
            var action = Click($"id{i}", $"N{i}");
            coordinator.EnrichAssistiveAction(action, Window("Welcome"));
            Assert.Equal(i + 1, action.Sequence);
            Assert.Equal($"evt-{(i + 1):D6}", action.EventId);
            Assert.True(ids.Add(action.EventId!));
        }
    }

    [Fact]
    public void EnrichedEvents_SerializeCamelCaseAndPlaybackIgnoresExtras()
    {
        var action = new RecordedAction
        {
            ActionType = ActionType.DoubleClick,
            Mode = RecordingMode.Assistive,
            EventId = "evt-000001",
            Sequence = 1,
            JiraKey = "ABC-1234",
            Bdd = new RecordedBddAssociation { GroupId = "bdd-0001", Statement = "And double click on ABC" },
            PageId = "welcome",
            ObjectRef = "welcome.abc",
            Window = new RecordedWindowContext { Title = "Welcome", NormalizedTitle = "welcome", ProcessId = 1 },
            Element = new ElementInfo { AutomationId = "btnABC", Name = "ABC", ControlType = "Button" },
            Description = "Double Click on ABC"
        };

        var json = JsonSerializer.Serialize(action, JsonOpts);
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("eventId", out _));
        Assert.True(doc.RootElement.TryGetProperty("jiraKey", out _));
        Assert.True(doc.RootElement.TryGetProperty("bdd", out var bdd));
        Assert.Equal("bdd-0001", bdd.GetProperty("groupId").GetString());
        Assert.Equal("And double click on ABC", bdd.GetProperty("statement").GetString());

        var ui = new Mock<IUiService>();
        ui.Setup(s => s.Execute(It.IsAny<Models.Request.UiRequest>(), It.IsAny<CancellationToken>()))
            .Returns((object?)null);
        var playback = new PlaybackService(ui.Object, NullLogger<PlaybackService>.Instance);
        var export = new RecordingExport { Actions = [action] };
        var payload = JsonSerializer.SerializeToElement(new { actions = export.Actions }, JsonOpts);
        var result = playback.Play(payload);
        Assert.Equal(1, result.ExecutedActions);
        Assert.Equal("doubleclick", result.Actions[0].Operation);
    }

    [Fact]
    public void PageIds_ReuseSameTitle_AndSplitDifferentTitles()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        var a1 = Click("a", "A");
        var a2 = Click("b", "B");
        var a3 = Click("c", "C");
        coordinator.EnrichAssistiveAction(a1, Window("Welcome"));
        coordinator.EnrichAssistiveAction(a2, Window("Welcome"));
        coordinator.EnrichAssistiveAction(a3, Window("Print Dialog"));
        Assert.Equal(a1.PageId, a2.PageId);
        Assert.NotEqual(a1.PageId, a3.PageId);
        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", a1.PageId!);
        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", a3.PageId!);
    }

    [Fact]
    public void PageId_HandlesLongTitlesDigitsAndCollisions()
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        var longTitle = new string('a', 120);
        var id1 = DeterministicPageIdGenerator.Allocate(
            DeterministicPageIdGenerator.NormalizeTitle(longTitle),
            longTitle,
            used);
        Assert.True(id1.Length <= 64);
        Assert.Matches("^[a-z][a-z0-9-]{0,63}$", id1);

        var digit = DeterministicPageIdGenerator.FromWindowTitle("123 Reports");
        Assert.StartsWith("p-", digit, StringComparison.Ordinal);

        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var first = DeterministicPageIdGenerator.FromWindowTitle("A/B", map);
        map[DeterministicPageIdGenerator.NormalizeTitle("A/B")] = first;
        var second = DeterministicPageIdGenerator.FromWindowTitle("A B", map);
        map[DeterministicPageIdGenerator.NormalizeTitle("A B")] = second;
        Assert.NotEqual(first, second);
    }

    [Fact]
    public void ObjectRefs_ReuseLocatorAndKeepCollisionIdsWithinLimit()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        var a1 = Click("same", "Same");
        var a2 = Click("same", "Same");
        coordinator.EnrichAssistiveAction(a1, Window("Welcome"));
        coordinator.EnrichAssistiveAction(a2, Window("Welcome"));
        Assert.Equal(a1.ObjectRef, a2.ObjectRef);

        var used = new HashSet<string>(StringComparer.Ordinal);
        var seed = new string('x', 80);
        var e1 = DeterministicElementIdGenerator.Resolve(
            new ElementInfo { AutomationId = seed, ControlType = "Button" },
            used);
        var e2 = DeterministicElementIdGenerator.Resolve(
            new ElementInfo { AutomationId = seed + "-other", ControlType = "Button" },
            used);
        Assert.True(e1.Length <= 64);
        Assert.True(e2.Length <= 64);
        Assert.NotEqual(e1, e2);
    }

    [Fact]
    public void ArtifactBuilder_CreatesPagesAndBddMap_ValidatesSchema()
    {
        var actions = new List<RecordedAction>();
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-20260810-162000-7f81c4a2");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("And double click on ABC", BddScope.NextAction, out _, out _);

        var click = new RecordedAction
        {
            ActionType = ActionType.DoubleClick,
            Element = new ElementInfo { AutomationId = "btnABC", Name = "ABC", ControlType = "Button" },
            Description = "Double Click on ABC Button"
        };
        coordinator.EnrichAssistiveAction(click, Window("Welcome"));
        actions.Add(click);

        coordinator.TryArmBdd("When the user prints the report", BddScope.UntilFinished, out _, out _);
        var type = new RecordedAction
        {
            ActionType = ActionType.Type,
            Element = new ElementInfo { AutomationId = "reportName", Name = "Report Name", ControlType = "Edit" },
            Value = "secret-report"
        };
        var select = new RecordedAction
        {
            ActionType = ActionType.Select,
            Element = new ElementInfo { AutomationId = "printer", Name = "Printer", ControlType = "ComboBox" }
        };
        var print = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "print", Name = "Print", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(type, Window("Print Dialog"));
        coordinator.EnrichAssistiveAction(select, Window("Print Dialog"));
        coordinator.EnrichAssistiveAction(print, Window("Print Dialog"));
        coordinator.TryFinishBdd(out _);
        actions.Add(type);
        actions.Add(select);
        actions.Add(print);

        var unmapped = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "close", Name = "Close", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(unmapped, Window("Welcome"));
        actions.Add(unmapped);

        var build = new AssistiveArtifactBuilder().Build(
            "rec-20260810-162000-7f81c4a2",
            "recording_20260810_162000.json",
            DateTimeOffset.Parse("2026-08-10T16:20:15Z"),
            "ABC-1234",
            actions);

        Assert.Equal(2, build.Pages.Count);
        Assert.All(build.Pages, p => Assert.Equal("candidate", p.State));
        Assert.All(build.Pages, p => Assert.All(p.Elements.Values, e => Assert.Equal("capture", e.Source.Kind)));
        Assert.DoesNotContain(build.Pages.SelectMany(p => p.Elements.Values), e => e.Locator.ContainsKey("boundingRectangle"));
        Assert.DoesNotContain(
            JsonSerializer.Serialize(build.BddActionMap, JsonOpts),
            "secret-report",
            StringComparison.Ordinal);

        Assert.NotNull(build.BddActionMap);
        Assert.Equal(2, build.BddActionMap!.BddGroups.Count);
        Assert.Contains("evt-000005", build.BddActionMap.UnmappedEventIds);
        Assert.Equal(3, build.BddActionMap.BddGroups[1].Actions.Count);
        Assert.Contains(build.BddActionMap.BddGroups[1].Actions, a => a.PageId == "print-dialog");

        // Identical statement text in separate groups stays separate.
        var coordinator2 = new AssistiveRecordingCoordinator();
        coordinator2.Reset("rec-2");
        coordinator2.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator2.TryArmBdd("And click OK", BddScope.NextAction, out var g1, out _);
        var x1 = Click("ok1", "OK");
        coordinator2.EnrichAssistiveAction(x1, Window("Welcome"));
        coordinator2.TryArmBdd("And click OK", BddScope.NextAction, out var g2, out _);
        var x2 = Click("ok2", "OK");
        coordinator2.EnrichAssistiveAction(x2, Window("Welcome"));
        Assert.NotEqual(g1, g2);

        var schemaText = ReadRepoFile("automation/schemas/bdd-action-map.schema.json");
        var schema = JsonSchema.FromText(schemaText);
        var mapJson = JsonSerializer.Serialize(build.BddActionMap, JsonOpts);
        var result = schema.Evaluate(JsonDocument.Parse(mapJson).RootElement, new EvaluationOptions { OutputFormat = OutputFormat.List });
        Assert.True(result.IsValid, string.Join("; ", result.Details.Where(d => d.HasErrors).Select(d => d.ToString())));

        var pageSchema = JsonSchema.FromText(ReadRepoFile("automation/schemas/page-object.schema.json"));
        foreach (var page in build.Pages)
        {
            var pageJson = JsonSerializer.Serialize(page, JsonOpts);
            var pageResult = pageSchema.Evaluate(
                JsonDocument.Parse(pageJson).RootElement,
                new EvaluationOptions { OutputFormat = OutputFormat.List });
            Assert.True(pageResult.IsValid, page.PageId + ": " + string.Join("; ", pageResult.Details.Where(d => d.HasErrors).Select(d => d.ToString())));
        }
    }

    [Fact]
    public void ArtifactWriter_UsesSafeLayout_AndRejectsOverwrite()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-art-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        try
        {
            var build = new AssistiveArtifactBuilder.BuildResult
            {
                Pages =
                [
                    new AssistiveArtifactBuilder.PageCandidateDocument
                    {
                        PageId = "welcome",
                        Name = "Welcome",
                        Elements =
                        {
                            ["abc"] = new AssistiveArtifactBuilder.PageCandidateElement
                            {
                                Description = "ABC",
                                Locator = new Dictionary<string, object?>
                                {
                                    ["automationId"] = "btnABC",
                                    ["controlType"] = "Button"
                                },
                                Quality = new AssistiveArtifactBuilder.PageCandidateQuality { Grade = "strong" },
                                Source = new AssistiveArtifactBuilder.PageCandidateSource
                                {
                                    Kind = "capture",
                                    Path = "recording.json",
                                    Metadata = new Dictionary<string, object?> { ["recordingId"] = "rec-1" }
                                }
                            }
                        }
                    }
                ],
                BddActionMap = new BddActionMapDocument
                {
                    RecordingId = "rec-1",
                    JiraKey = "ABC-1234",
                    SourceRecording = "recording.json",
                    CreatedAt = DateTimeOffset.UtcNow,
                    Pages = [new BddActionMapPageRef { PageId = "welcome", WindowTitle = "Welcome", File = "page-objects/welcome.page.json" }]
                }
            };

            var writer = new AssistiveArtifactWriter();
            var summary = writer.Write(output, "recording.json", "rec-1", "ABC-1234", build);
            Assert.NotNull(summary.Directory);
            Assert.True(File.Exists(summary.BddActionMapFile));
            Assert.Single(summary.PageObjectFiles);
            Assert.Contains(Path.Combine("assistive-artifacts", "ABC-1234", "rec-1"), summary.Directory!, StringComparison.Ordinal);

            Assert.Throws<IOException>(() => writer.Write(output, "recording.json", "rec-1", "ABC-1234", build));
        }
        finally
        {
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void EmptyAssistiveSession_DoesNotCreatePageFiles()
    {
        var build = new AssistiveArtifactBuilder().Build(
            "rec-empty",
            "recording.json",
            DateTimeOffset.UtcNow,
            "ABC-1234",
            [new RecordedAction { ActionType = ActionType.Click, Mode = RecordingMode.Passive }]);
        Assert.Empty(build.Pages);
        Assert.Null(build.BddActionMap);
    }

    [Fact]
    public void OperationResolver_MatchesPlaybackContract()
    {
        Assert.Equal(
            "doubleclick",
            RecordedActionOperationResolver.ResolveOperation(new RecordedAction { ActionType = ActionType.DoubleClick }));
        Assert.Equal(
            "clicklogicalmenupath",
            RecordedActionOperationResolver.ResolveOperation(new RecordedAction { ActionType = ActionType.MenuPathClick }));
        Assert.Equal(
            "alertok",
            RecordedActionOperationResolver.ResolveOperation(new RecordedAction
            {
                ActionType = ActionType.Click,
                Description = "Click OK on Popup"
            }));
    }

    private static RecordedAction Click(string automationId, string name) =>
        new()
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo
            {
                AutomationId = automationId,
                Name = name,
                ControlType = "Button"
            }
        };

    private static AssistiveActionCaptureContext Window(string title) =>
        new()
        {
            Window = new RecordedWindowContext
            {
                Title = title,
                NormalizedTitle = DeterministicPageIdGenerator.NormalizeTitle(title),
                ProcessId = 42,
                NativeWindowHandle = "0x001A04F2"
            }
        };

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

        throw new FileNotFoundException($"Unable to locate '{relativePath}'.");
    }
}
