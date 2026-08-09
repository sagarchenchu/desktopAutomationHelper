using System.Text.Json;
using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Playback;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class PlaybackContractTests
{
    private readonly Mock<IUiService> _ui = new();
    private readonly PlaybackService _service;
    private readonly PlaybackController _controller;

    public PlaybackContractTests()
    {
        _service = new PlaybackService(_ui.Object, NullLogger<PlaybackService>.Instance);
        _controller = new PlaybackController(_service, NullLogger<PlaybackController>.Instance);
    }

    [Fact]
    public void Play_RawRecordingExport_ExecutesAssistiveActions()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Returns(new { success = true });

        var payload = JsonDocument.Parse("""
        {
          "mode": "Assistive",
          "actions": [
            {
              "actionType": "click",
              "mode": "assistive",
              "element": { "automationId": "btnSubmit", "controlType": "Button" },
              "description": "Click Submit"
            }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.TotalActions);
        Assert.Equal(1, result.ExecutedActions);
        Assert.Equal(0, result.SkippedActions);
        Assert.True(result.Completed);
        _ui.Verify(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "click"), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public void Play_BareActionsArray_IsAccepted()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);

        var payload = JsonDocument.Parse("""
        [
          {
            "actionType": "type",
            "mode": "assistive",
            "value": "hello",
            "element": { "automationId": "txt" }
          }
        ]
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.ExecutedActions);
    }

    [Fact]
    public void Play_RecordingWrapper_IsAccepted()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "continueOnError": false,
          "recording": {
            "actions": [
              {
                "actionType": "hover",
                "mode": "assistive",
                "element": { "name": "Field" }
              }
            ]
          }
        }
        """).RootElement;

        Assert.Equal(1, _service.Play(payload).ExecutedActions);
    }

    [Fact]
    public void Play_WebDriverValueWrapper_IsAccepted()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "value": {
            "actions": [
              {
                "actionType": "isVisible",
                "mode": "assistive",
                "element": { "name": "Label" }
              }
            ]
          }
        }
        """).RootElement;

        Assert.Equal(1, _service.Play(payload).ExecutedActions);
    }

    [Fact]
    public void Play_PassiveActions_AreSkipped()
    {
        var payload = JsonDocument.Parse("""
        {
          "actions": [
            {
              "actionType": "click",
              "mode": "passive",
              "element": { "automationId": "btn" }
            }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.SkippedActions);
        Assert.Equal(0, result.ExecutedActions);
        Assert.Contains("Assistive", result.Actions[0].SkipReason);
        _ui.Verify(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public void Play_ExplicitOperation_TakesPrecedence()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "actions": [
            {
              "actionType": "menuPathClick",
              "mode": "assistive",
              "operation": "contextmenupath",
              "value": "Edit>Copy",
              "element": { "name": "Edit" }
            }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal("contextmenupath", result.Actions[0].Operation);
        _ui.Verify(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "contextmenupath"), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public void Play_UnsupportedAction_IsSkippedWithReason()
    {
        var payload = JsonDocument.Parse("""
        {
          "actions": [
            {
              "actionType": "isDisabled",
              "mode": "assistive",
              "element": { "name": "x" }
            }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.SkippedActions);
        Assert.Contains("Unsupported", result.Actions[0].SkipReason);
    }

    [Fact]
    public void Play_MissingLocator_IsSkipped()
    {
        var payload = JsonDocument.Parse("""
        {
          "actions": [
            {
              "actionType": "click",
              "mode": "assistive"
            }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.SkippedActions);
        Assert.Contains("element information", result.Actions[0].SkipReason);
    }

    [Fact]
    public void Play_ContinueOnErrorFalse_StopsAfterFailure()
    {
        _ui.SetupSequence(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new InvalidOperationException("boom"))
            .Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "continueOnError": false,
          "actions": [
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } },
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "b" } }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.False(result.Completed);
        Assert.Equal(1, result.FailedActions);
        // Playback returns immediately after the first failure when continueOnError is false.
        Assert.Equal(1, result.Actions.Count);
        _ui.Verify(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public void Play_ContinueOnErrorTrue_ContinuesAfterFailure()
    {
        _ui.SetupSequence(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new InvalidOperationException("boom"))
            .Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "continueOnError": true,
          "actions": [
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } },
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "b" } }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(1, result.FailedActions);
        Assert.Equal(1, result.ExecutedActions);
        _ui.Verify(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()), Times.Exactly(2));
    }

    [Fact]
    public void Play_DelayMs_IsAccepted()
    {
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);

        var payload = JsonDocument.Parse("""
        {
          "delayMs": 5,
          "actions": [
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } },
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "b" } }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(2, result.ExecutedActions);
        Assert.True(result.Completed);
    }

    [Fact]
    public void Controller_InvalidPayload_Returns400()
    {
        var payload = JsonDocument.Parse("""{ "foo": 1 }""").RootElement;
        var result = _controller.Play(payload);
        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public void Controller_UnexpectedFailure_Returns500()
    {
        var playback = new Mock<IPlaybackService>();
        playback.Setup(p => p.Play(It.IsAny<JsonElement>())).Throws(new Exception("unexpected"));
        var controller = new PlaybackController(playback.Object, NullLogger<PlaybackController>.Instance);

        var result = controller.Play(JsonDocument.Parse("""{ "actions": [] }""").RootElement);
        var obj = Assert.IsType<ObjectResult>(result);
        Assert.Equal(500, obj.StatusCode);
    }

    [Fact]
    public void Play_Counters_AreCorrect()
    {
        _ui.SetupSequence(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Returns((object?)null)
            .Throws(new ArgumentException("bad"));

        var payload = JsonDocument.Parse("""
        {
          "continueOnError": true,
          "actions": [
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } },
            { "actionType": "click", "mode": "passive", "element": { "automationId": "b" } },
            { "actionType": "click", "mode": "assistive", "element": { "automationId": "c" } }
          ]
        }
        """).RootElement;

        var result = _service.Play(payload);
        Assert.Equal(3, result.TotalActions);
        Assert.Equal(1, result.ExecutedActions);
        Assert.Equal(1, result.SkippedActions);
        Assert.Equal(1, result.FailedActions);
    }
}
