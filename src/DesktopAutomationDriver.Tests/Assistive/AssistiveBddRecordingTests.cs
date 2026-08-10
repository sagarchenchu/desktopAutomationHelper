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
    public void ArmBdd_RejectsReplaceWhileMultiActionActive()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        Assert.True(coordinator.TryArmBdd("When group A", BddScope.UntilFinished, out var g1, out _));
        coordinator.EnrichAssistiveAction(Click("a", "A"), Window("Welcome"));

        Assert.False(coordinator.TryArmBdd("When group B", BddScope.NextAction, out _, out var error));
        Assert.Contains("Finish or Cancel", error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(g1, coordinator.ActiveBddGroupId);
    }

    [Fact]
    public void ArmBdd_RejectsReplaceWhilePendingUnused()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-1");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        Assert.True(coordinator.TryArmBdd("And click OK", BddScope.NextAction, out var g1, out _));
        Assert.False(coordinator.TryArmBdd("And click Save", BddScope.NextAction, out _, out var error));
        Assert.Contains("pending", error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(g1, coordinator.ActiveBddGroupId);
    }

    [Fact]
    public void UnsupportedOperations_AreUnmappedWithWarning_NotInventedFallbacks()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-unsupported");
        coordinator.TryStartJiraRecording("ABC-1234", out _, out _);
        coordinator.TryArmBdd("Then assert something", BddScope.NextAction, out _, out _);

        var assertAction = new RecordedAction
        {
            ActionType = ActionType.Assert,
            Element = new ElementInfo { AutomationId = "x", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(assertAction, Window("Welcome"));

        var disabled = new RecordedAction
        {
            ActionType = ActionType.IsDisabled,
            Element = new ElementInfo { AutomationId = "y", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(disabled, Window("Welcome"));

        Assert.Null(RecordedActionOperationResolver.ResolveOperation(assertAction));
        Assert.Null(RecordedActionOperationResolver.ResolveOperation(disabled));

        var build = new AssistiveArtifactBuilder().Build(
            "rec-unsupported",
            "recording.json",
            DateTimeOffset.UtcNow,
            "ABC-1234",
            [assertAction, disabled]);

        Assert.NotNull(build.BddActionMap);
        Assert.Contains("evt-000001", build.BddActionMap!.UnmappedEventIds);
        Assert.Contains("evt-000002", build.BddActionMap.UnmappedEventIds);
        Assert.Empty(build.BddActionMap.BddGroups);
        Assert.Contains(
            build.BddActionMap.Warnings ?? [],
            w => w.Contains("unsupported", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(
            JsonSerializer.Serialize(build.BddActionMap, JsonOpts),
            "\"operation\":\"assert\"",
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            JsonSerializer.Serialize(build.BddActionMap, JsonOpts),
            "isDisabled",
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ActionType.Type)]
    [InlineData(ActionType.TypeAndSelect)]
    [InlineData(ActionType.GetTableHeaders)]
    [InlineData(ActionType.GetTableData)]
    [InlineData(ActionType.SwitchWindow)]
    [InlineData(ActionType.Select)]
    public void AssistiveActions_AcrossTwoWindowTitles_KeepDistinctPageIds(ActionType actionType)
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-windows");

        var a1 = new RecordedAction
        {
            ActionType = actionType,
            Element = new ElementInfo { AutomationId = "ctrl-a", Name = "Control A", ControlType = "Edit" },
            Value = actionType is ActionType.Type or ActionType.TypeAndSelect ? "typed" : null
        };
        var a2 = new RecordedAction
        {
            ActionType = actionType,
            Element = new ElementInfo { AutomationId = "ctrl-b", Name = "Control B", ControlType = "ComboBox" },
            Value = actionType is ActionType.Type or ActionType.TypeAndSelect ? "typed" : null
        };

        coordinator.EnrichAssistiveAction(a1, Window("Orders"));
        coordinator.EnrichAssistiveAction(a2, Window("Reports"));

        Assert.NotNull(a1.Window);
        Assert.NotNull(a2.Window);
        Assert.Equal("Orders", a1.Window!.Title);
        Assert.Equal("Reports", a2.Window!.Title);
        Assert.NotEqual(a1.PageId, a2.PageId);
        Assert.Equal("orders", a1.PageId);
        Assert.Equal("reports", a2.PageId);
    }

    [Fact]
    public void PopupAction_AcrossTwoWindows_KeepsDistinctPageIds()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-popup");

        var popup1 = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "ok", Name = "OK", ControlType = "Button" },
            Description = "Click OK on Popup"
        };
        var popup2 = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "cancel", Name = "Cancel", ControlType = "Button" },
            Description = "Click Cancel on Popup"
        };

        coordinator.EnrichAssistiveAction(popup1, Window("Confirm Delete"));
        coordinator.EnrichAssistiveAction(popup2, Window("Save Changes"));
        Assert.NotEqual(popup1.PageId, popup2.PageId);
    }

    [Fact]
    public void Phase3LocatorContract_HeaderItem_OmitsUnsupportedControlType_WhenAutomationIdPresent()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-header");
        var action = new RecordedAction
        {
            ActionType = ActionType.GetTableHeaders,
            Element = new ElementInfo
            {
                AutomationId = "colAmount",
                Name = "Amount",
                ControlType = "HeaderItem"
            }
        };
        coordinator.EnrichAssistiveAction(action, Window("Grid"));

        var build = new AssistiveArtifactBuilder().Build(
            "rec-header",
            "recording.json",
            DateTimeOffset.UtcNow,
            null,
            [action]);

        Assert.Single(build.Pages);
        var element = Assert.Single(build.Pages[0].Elements);
        Assert.Equal("colAmount", element.Value.Locator["automationId"]);
        Assert.False(element.Value.Locator.ContainsKey("controlType"));
        Assert.True(Phase3LocatorContract.IsValid(element.Value.Locator, out _), "locator must pass Phase 3 contract");
        AssertNullOrEmptyUnresolved(build.Pages[0]);
    }

    [Theory]
    [InlineData("HeaderItem")]
    [InlineData("TabItem")]
    [InlineData("TitleBar")]
    [InlineData("TotallyUnknownType")]
    public void Phase3LocatorContract_UnsupportedControlTypeWithoutAutomationId_GoesUnresolved(string controlType)
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-unresolved-ct");
        var action = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo
            {
                Name = "Label",
                ControlType = controlType
            }
        };
        coordinator.EnrichAssistiveAction(action, Window("Welcome"));
        Assert.Null(action.ObjectRef);

        var build = new AssistiveArtifactBuilder().Build(
            "rec-unresolved-ct",
            "recording.json",
            DateTimeOffset.UtcNow,
            null,
            [action]);

        Assert.Single(build.Pages);
        Assert.Empty(build.Pages[0].Elements);
        Assert.NotNull(build.Pages[0].Unresolved);
        Assert.Contains(
            build.Pages[0].Unresolved!,
            u => (u.TryGetValue("reason", out var reason) ? reason?.ToString() : null)?
                .Contains("not recognized", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public void Phase3LocatorContract_AutomationIdOnly_IsValidWithoutControlType()
    {
        var locator = new Dictionary<string, object?>
        {
            ["automationId"] = "only-id"
        };
        Assert.True(Phase3LocatorContract.IsValid(locator, out var error), error);
    }

    [Fact]
    public void AssistivePathSafety_RejectsDirectorySymlinkEscape()
    {
        var root = Path.Combine(Path.GetTempPath(), "da-as-symlink-" + Guid.NewGuid().ToString("N"));
        var outside = Path.Combine(Path.GetTempPath(), "da-as-outside-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), "{}");
        var linkPath = Path.Combine(root, "assistive-artifacts");

        try
        {
            RequireSymbolicLink(() => Directory.CreateSymbolicLink(linkPath, outside), linkPath);
            var target = Path.Combine(linkPath, "secret.json");
            var thrown = Assert.Throws<IOException>(() =>
                AssistivePathSafety.EnsureNotSymlinkEscape(target, root));
            Assert.Contains("reparse point", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            try { Directory.Delete(outside, recursive: true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void AssistivePathSafety_RejectsChainedFileSymlinkEscape()
    {
        var root = Path.Combine(Path.GetTempPath(), "da-as-chain-" + Guid.NewGuid().ToString("N"));
        var outside = Path.Combine(Path.GetTempPath(), "da-as-chain-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);
        var outsideFile = Path.Combine(outside, "secret.json");
        File.WriteAllText(outsideFile, "{}");
        var mid = Path.Combine(root, "mid.json");
        var entry = Path.Combine(root, "page.json");

        try
        {
            RequireSymbolicLink(() => File.CreateSymbolicLink(mid, outsideFile), mid);
            RequireSymbolicLink(() => File.CreateSymbolicLink(entry, mid), entry);

            var thrown = Assert.Throws<IOException>(() =>
                AssistivePathSafety.EnsureNotSymlinkEscape(entry, root));
            Assert.Contains("outside", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try { File.Delete(entry); } catch { /* ignore */ }
            try { File.Delete(mid); } catch { /* ignore */ }
            try { Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            try { Directory.Delete(outside, recursive: true); } catch { /* ignore */ }
        }
    }

    [WindowsFact]
    public void AssistivePathSafety_RejectsWindowsJunctionEscape()
    {
        var root = Path.Combine(Path.GetTempPath(), "da-as-junc-" + Guid.NewGuid().ToString("N"));
        var outside = Path.Combine(Path.GetTempPath(), "da-as-junc-out-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(outside, "secret.json"), "{}");
        var junction = Path.Combine(root, "assistive-artifacts");

        try
        {
            RequireWindowsJunction(junction, outside);
            var target = Path.Combine(junction, "secret.json");
            var thrown = Assert.Throws<IOException>(() =>
                AssistivePathSafety.EnsureNotSymlinkEscape(target, root));
            Assert.Contains("reparse point", thrown.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try { Directory.Delete(junction); } catch { /* ignore */ }
            try { Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            try { Directory.Delete(outside, recursive: true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void ArtifactWriter_CleansStaging_OnInjectedWriteFailure()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-art-fail-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        try
        {
            var build = MinimalBuild("rec-fail");
            var writer = new AssistiveArtifactWriter
            {
                BeforeCommitStaging = (_, _) => throw new IOException("injected staging failure")
            };

            var thrown = Assert.Throws<IOException>(() =>
                writer.Write(output, "recording.json", "rec-fail", "ABC-1234", build));
            Assert.Contains("injected", thrown.Message, StringComparison.Ordinal);

            var artifactsRoot = Path.Combine(output, "assistive-artifacts", "ABC-1234");
            if (Directory.Exists(artifactsRoot))
            {
                Assert.Empty(Directory.GetDirectories(artifactsRoot, "rec-fail"));
                Assert.Empty(Directory.GetDirectories(artifactsRoot, ".staging-*"));
            }
        }
        finally
        {
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void Export_PreservesPrimary_OnArtifactFailure_AndAllowsPrimaryWriteRetry()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
            var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.ConfigureAssistiveSessionForTests(output, "rec-export-1");
            service.RecordAssistiveAction(
                new RecordedAction
                {
                    ActionType = ActionType.Type,
                    Element = new ElementInfo { AutomationId = "name", ControlType = "Edit" },
                    Value = "secret"
                },
                Window("Welcome"));

            service.ArtifactWriterForTests = new AssistiveArtifactWriter
            {
                BeforeCommitStaging = (_, _) => throw new IOException("injected artifact failure")
            };

            service.ExportForTests();
            Assert.True(service.ExportCompletedForTests);
            Assert.NotNull(service.ExportFilePathForTests);
            Assert.True(File.Exists(service.ExportFilePathForTests));
            var primary = File.ReadAllText(service.ExportFilePathForTests!);
            Assert.Contains("evt-000001", primary, StringComparison.Ordinal);
            Assert.Contains("Primary recording was preserved", primary, StringComparison.OrdinalIgnoreCase);
            Assert.False(
                Directory.Exists(Path.Combine(output, "assistive-artifacts", "unassigned", "rec-export-1")));

            // Primary write failure leaves export incomplete and retryable.
            service.ConfigureAssistiveSessionForTests(output, "rec-export-2");
            service.RecordAssistiveAction(
                Click("ok", "OK"),
                Window("Welcome"));
            var failOnce = true;
            service.BeforePrimaryWriteForTests = _ =>
            {
                if (failOnce)
                {
                    failOnce = false;
                    throw new IOException("injected primary failure");
                }
            };

            Assert.Throws<IOException>(() => service.ExportForTests());
            Assert.False(service.ExportCompletedForTests);

            service.ExportForTests();
            Assert.True(service.ExportCompletedForTests);
            Assert.True(File.Exists(service.ExportFilePathForTests));
            Assert.Contains("evt-000001", File.ReadAllText(service.ExportFilePathForTests!), StringComparison.Ordinal);
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void AddAction_InAssistiveMode_RequiresRecordAssistiveAction()
    {
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);
        try
        {
            service.ConfigureAssistiveSessionForTests(
                Path.Combine(Path.GetTempPath(), "da-add-" + Guid.NewGuid().ToString("N")),
                "rec-add");
            var ex = Assert.Throws<InvalidOperationException>(() =>
                service.AddAction(new RecordedAction { ActionType = ActionType.Click }));
            Assert.Contains("RecordAssistiveAction", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            service.Dispose();
        }
    }

    [Fact]
    public void Export_IsSingleFlight_ConcurrentCallersShareOneExport()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-export-sf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.ConfigureAssistiveSessionForTests(output, "rec-single-flight");
            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));

            var enteredStaging = new ManualResetEventSlim(false);
            var releaseStaging = new ManualResetEventSlim(false);
            var stagingEntries = 0;

            service.ArtifactWriterForTests = new AssistiveArtifactWriter
            {
                BeforeCommitStaging = (_, _) =>
                {
                    Interlocked.Increment(ref stagingEntries);
                    enteredStaging.Set();
                    if (!releaseStaging.Wait(TimeSpan.FromSeconds(10)))
                        throw new TimeoutException("Release staging timed out.");
                }
            };

            var exportTask = Task.Run(() => service.ExportForTests());
            Assert.True(enteredStaging.Wait(TimeSpan.FromSeconds(10)));
            Assert.Equal(RecordingExportState.InProgress, service.ExportStateForTests);

            var startWhileExporting = service.StartRecording(new Models.Request.StartRecordingRequest
            {
                OutputPath = output
            });
            Assert.NotNull(startWhileExporting.Error);
            Assert.Contains("stop/export is still pending", startWhileExporting.Error, StringComparison.OrdinalIgnoreCase);

            var statusTask = Task.Run(() => service.GetCurrentState());
            Thread.Sleep(200);
            Assert.False(statusTask.IsCompleted);

            releaseStaging.Set();
            Assert.True(exportTask.Wait(TimeSpan.FromSeconds(10)));
            var status = statusTask.Result;

            Assert.True(service.ExportCompletedForTests);
            Assert.Equal(1, stagingEntries);
            Assert.NotNull(status.ExportedFilePath);
            Assert.True(Directory.Exists(Path.Combine(output, "assistive-artifacts", "unassigned", "rec-single-flight")));
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void Export_SummaryRewriteFailure_DoesNotReportArtifactExportFailure()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-export-rewrite-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.ConfigureAssistiveSessionForTests(output, "rec-rewrite");
            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));
            service.BeforePrimarySummaryRewriteForTests = () =>
                throw new IOException("injected summary rewrite failure");

            service.ExportForTests();

            Assert.True(service.ExportCompletedForTests);
            Assert.True(Directory.Exists(Path.Combine(output, "assistive-artifacts", "unassigned", "rec-rewrite")));
            Assert.NotNull(service.GetCurrentState().Artifacts);
            Assert.Contains(
                service.GetCurrentState().Artifacts!.Warnings,
                w => w.Contains("could not be updated with the artifact summary", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(
                service.GetCurrentState().Artifacts!.Warnings,
                w => w.Contains("Assistive artifact export failed", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void CaptureContext_IsImmutable_AndUsedAtRecordTime()
    {
        var coordinator = new AssistiveRecordingCoordinator();
        coordinator.Reset("rec-immutable");
        var preCaptured = Window("Orders");
        // Mutating a different window title after capture must not affect the recorded page.
        var action = new RecordedAction
        {
            ActionType = ActionType.Click,
            Element = new ElementInfo { AutomationId = "ok", ControlType = "Button" }
        };
        coordinator.EnrichAssistiveAction(action, preCaptured);
        Assert.Equal("Orders", action.Window!.Title);
        Assert.Equal("orders", action.PageId);
    }

    [Fact]
    public void StartRecording_RejectedDuringStopToExportTransition()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-stop-gap-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.SuppressOverlayForTests = true;
            service.ConfigureAssistiveSessionForTests(output, "rec-stop-gap");
            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));
            var originalRecordingId = service.RecordingId;

            var enteredGap = new ManualResetEventSlim(false);
            var releaseGap = new ManualResetEventSlim(false);

            service.AfterBeginStopBeforeExportForTests = () =>
            {
                Assert.Equal(RecordingLifecycleState.Stopping, service.LifecycleStateForTests);
                Assert.False(service.IsActive);
                Assert.Equal(RecordingExportState.NotStarted, service.ExportStateForTests);
                enteredGap.Set();
                if (!releaseGap.Wait(TimeSpan.FromSeconds(10)))
                    throw new TimeoutException("Release stop-gap timed out.");
            };

            var exportTask = Task.Run(() => service.ExportForTests());
            Assert.True(enteredGap.Wait(TimeSpan.FromSeconds(10)));

            var startDuringGap = service.StartRecording(new Models.Request.StartRecordingRequest
            {
                OutputPath = output
            });
            Assert.NotNull(startDuringGap.Error);
            Assert.Contains("stop/export is still pending", startDuringGap.Error, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalRecordingId, service.RecordingId);
            Assert.Equal(RecordingLifecycleState.Stopping, service.LifecycleStateForTests);

            releaseGap.Set();
            Assert.True(exportTask.Wait(TimeSpan.FromSeconds(10)));
            Assert.True(service.ExportCompletedForTests);
            Assert.Equal(RecordingLifecycleState.Idle, service.LifecycleStateForTests);
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void StartRecording_TwoSimultaneousStarts_ExactlyOneSucceeds()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-dual-start-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.SuppressOverlayForTests = true;

            var barrier = new Barrier(2);
            StartRecordingResult? r1 = null;
            StartRecordingResult? r2 = null;

            var t1 = Task.Run(() =>
            {
                barrier.SignalAndWait(TimeSpan.FromSeconds(10));
                r1 = service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output });
            });
            var t2 = Task.Run(() =>
            {
                barrier.SignalAndWait(TimeSpan.FromSeconds(10));
                r2 = service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output });
            });

            Assert.True(Task.WaitAll([t1, t2], TimeSpan.FromSeconds(15)));
            Assert.NotNull(r1);
            Assert.NotNull(r2);

            var successCount = (r1!.Error is null ? 1 : 0) + (r2!.Error is null ? 1 : 0);
            var failureCount = (r1.Error is not null ? 1 : 0) + (r2.Error is not null ? 1 : 0);
            Assert.Equal(1, successCount);
            Assert.Equal(1, failureCount);

            var failed = r1.Error is not null ? r1 : r2;
            Assert.Contains("already active", failed.Error, StringComparison.OrdinalIgnoreCase);
            Assert.True(service.IsActive);
            Assert.Equal(RecordingLifecycleState.Active, service.LifecycleStateForTests);
            Assert.NotNull(service.RecordingId);
        }
        finally
        {
            try { service.StopRecording(); } catch { /* ignore */ }
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void Export_SummaryRewriteFailure_WithoutSidecars_DoesNotClaimSidecarsWereWritten()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-export-nosidecar-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.ConfigureAssistiveSessionForTests(output, "rec-nosidecar");
            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));
            service.ArtifactWriterForTests = new AssistiveArtifactWriter
            {
                BeforeCommitStaging = (_, _) => throw new IOException("injected sidecar failure")
            };
            service.BeforePrimarySummaryRewriteForTests = () =>
                throw new IOException("injected summary rewrite failure");

            service.ExportForTests();

            Assert.True(service.ExportCompletedForTests);
            var warnings = service.GetCurrentState().Artifacts!.Warnings;
            Assert.Contains(warnings, w => w.Contains("Assistive artifact export failed", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(
                warnings,
                w => w.Contains("could not be updated with the Assistive artifact summary", StringComparison.OrdinalIgnoreCase)
                     || w.Contains("could not be updated with the artifact summary", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(
                warnings,
                w => w.Contains("sidecars were written", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void PublicStopRecording_WithOverlaySuppressed_ExportsAndBecomesIdle()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-public-stop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.SuppressOverlayForTests = true;
            var start = service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output });
            Assert.Null(start.Error);
            Assert.True(service.IsActive);
            Assert.Equal(RecordingLifecycleState.Active, service.LifecycleStateForTests);

            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));
            var recordingId = service.RecordingId;

            var export = service.StopRecording();
            Assert.False(service.IsActive);
            Assert.Equal(RecordingLifecycleState.Idle, service.LifecycleStateForTests);
            Assert.True(service.ExportCompletedForTests);
            Assert.Equal(recordingId, export.RecordingId);
            Assert.NotNull(export.ExportedFilePath);
            Assert.True(File.Exists(export.ExportedFilePath));
            Assert.Contains("evt-000001", File.ReadAllText(export.ExportedFilePath!), StringComparison.Ordinal);
        }
        finally
        {
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public void PrimaryExportFailure_KeepsStopping_RejectsStart_UntilRetrySucceeds()
    {
        var output = Path.Combine(Path.GetTempPath(), "da-primary-keep-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        var session = new Mock<IUiSessionContext>();
        session.SetupGet(s => s.ActiveSession).Returns((AutomationSession?)null);
        var service = new RecordingService(NullLogger<RecordingService>.Instance, session.Object);

        try
        {
            service.SuppressOverlayForTests = true;
            Assert.Null(service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output }).Error);
            service.RecordAssistiveAction(Click("ok", "OK"), Window("Welcome"));
            var originalId = service.RecordingId;
            Assert.NotNull(originalId);

            var failOnce = true;
            service.BeforePrimaryWriteForTests = _ =>
            {
                if (failOnce)
                {
                    failOnce = false;
                    throw new IOException("injected primary failure");
                }
            };

            Assert.Throws<IOException>(() => service.ExportForTests());
            Assert.Equal(RecordingLifecycleState.Stopping, service.LifecycleStateForTests);
            Assert.False(service.ExportCompletedForTests);
            Assert.Equal(originalId, service.RecordingId);

            var blocked = service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output });
            Assert.NotNull(blocked.Error);
            Assert.Contains("stop/export is still pending", blocked.Error, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalId, service.RecordingId);

            // Retry succeeds without clearing the session.
            service.ExportForTests();
            Assert.True(service.ExportCompletedForTests);
            Assert.Equal(RecordingLifecycleState.Idle, service.LifecycleStateForTests);
            Assert.Equal(originalId, service.GetCurrentState().RecordingId);

            var allowed = service.StartRecording(new Models.Request.StartRecordingRequest { OutputPath = output });
            Assert.Null(allowed.Error);
            Assert.True(service.IsActive);
            Assert.NotEqual(originalId, service.RecordingId);
        }
        finally
        {
            try { service.StopRecording(); } catch { /* ignore */ }
            service.Dispose();
            try { Directory.Delete(output, true); } catch { /* ignore */ }
        }
    }

    private static AssistiveArtifactBuilder.BuildResult MinimalBuild(string recordingId) =>
        new()
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
                                Metadata = new Dictionary<string, object?> { ["recordingId"] = recordingId }
                            }
                        }
                    }
                }
            ],
            BddActionMap = new BddActionMapDocument
            {
                RecordingId = recordingId,
                JiraKey = "ABC-1234",
                SourceRecording = "recording.json",
                CreatedAt = DateTimeOffset.UtcNow,
                Pages =
                [
                    new BddActionMapPageRef
                    {
                        PageId = "welcome",
                        WindowTitle = "Welcome",
                        File = "page-objects/welcome.page.json"
                    }
                ]
            }
        };

    private static void AssertNullOrEmptyUnresolved(AssistiveArtifactBuilder.PageCandidateDocument page) =>
        Assert.True(page.Unresolved is null || page.Unresolved.Count == 0);

    private static void RequireSymbolicLink(Action create, string path)
    {
        try
        {
            create();
        }
        catch (Exception ex) when (ex is IOException or PlatformNotSupportedException or UnauthorizedAccessException)
        {
            Assert.Fail(
                $"Unable to create symbolic link '{path}' required for path-safety coverage. {ex.GetType().Name}: {ex.Message}");
        }

        if (!File.Exists(path) && !Directory.Exists(path))
            Assert.Fail($"Symbolic link '{path}' was not created.");
    }

    private static void RequireWindowsJunction(string junctionPath, string targetPath)
    {
        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c mklink /J \"{junctionPath}\" \"{targetPath}\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var process = System.Diagnostics.Process.Start(psi);
            Assert.NotNull(process);
            process.WaitForExit(5000);
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            if (process.ExitCode != 0 || !Directory.Exists(junctionPath))
            {
                Assert.Fail(
                    $"Unable to create Windows junction '{junctionPath}' -> '{targetPath}' " +
                    $"(exit {process.ExitCode}). stdout: {stdout} stderr: {stderr}");
            }
        }
        catch (Exception ex) when (ex is IOException or PlatformNotSupportedException or UnauthorizedAccessException)
        {
            Assert.Fail(
                $"Unable to create Windows junction '{junctionPath}' required for path-safety coverage. {ex.GetType().Name}: {ex.Message}");
        }
    }

    private sealed class WindowsFactAttribute : FactAttribute
    {
        public WindowsFactAttribute()
        {
            if (!OperatingSystem.IsWindows())
                Skip = "Windows-only junction coverage.";
        }
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
