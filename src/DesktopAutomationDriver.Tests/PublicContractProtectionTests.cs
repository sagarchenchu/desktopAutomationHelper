using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class PublicContractProtectionTests
{
    private const string Token = "contract-token";

    [Fact]
    public async Task Verify_RemainsUnauthenticated()
    {
        var nextCalled = false;
        var middleware = new BearerTokenMiddleware(
            _ => { nextCalled = true; return Task.CompletedTask; },
            MockDriverContext());

        var ctx = new DefaultHttpContext();
        ctx.Request.Path = "/verify";
        ctx.Response.Body = new MemoryStream();
        await middleware.InvokeAsync(ctx);
        Assert.True(nextCalled);
    }

    [Theory]
    [InlineData("/status")]
    [InlineData("/ui")]
    [InlineData("/ui/operations")]
    public async Task ProtectedRoutes_RequireAuthentication(string path)
    {
        var middleware = new BearerTokenMiddleware(
            _ => Task.CompletedTask,
            MockDriverContext());

        var ctx = new DefaultHttpContext();
        ctx.Request.Path = path;
        ctx.Response.Body = new MemoryStream();
        await middleware.InvokeAsync(ctx);
        Assert.Equal(401, ctx.Response.StatusCode);
    }

    [Theory]
    [InlineData("/status")]
    [InlineData("/ui")]
    [InlineData("/ui/operations")]
    public async Task ProtectedRoutes_AcceptValidBearerToken(string path)
    {
        var nextCalled = false;
        var middleware = new BearerTokenMiddleware(
            _ => { nextCalled = true; return Task.CompletedTask; },
            MockDriverContext());

        var ctx = new DefaultHttpContext();
        ctx.Request.Path = path;
        ctx.Request.Headers.Authorization = $"Bearer {Token}";
        ctx.Response.Body = new MemoryStream();
        await middleware.InvokeAsync(ctx);
        Assert.True(nextCalled);
    }

    [Fact]
    public void UiRequest_JsonPropertyNames_RemainStable()
    {
        var json = """
        {
          "operation": "click",
          "locator": { "automationId": "btn", "name": "OK", "controlType": "Button", "matchMode": "exact" },
          "locator2": { "automationId": "other" },
          "value": "x",
          "index": 1,
          "columnIndex": 2,
          "timeoutMs": 1000,
          "pollIntervalMs": 100,
          "view": "raw",
          "root": "activeWindow",
          "maxDepth": 5,
          "maxChildren": 50,
          "includeOffscreen": true
        }
        """;

        var request = JsonSerializer.Deserialize<UiRequest>(json, JsonOptions());
        Assert.NotNull(request);
        Assert.Equal("click", request!.Operation);
        Assert.Equal("btn", request.Locator!.AutomationId);
        Assert.Equal("OK", request.Locator.Name);
        Assert.Equal("Button", request.Locator.ControlType);
        Assert.Equal("exact", request.Locator.MatchMode);
        Assert.Equal("other", request.Locator2!.AutomationId);
        Assert.Equal("x", request.Value);
        Assert.Equal(1, request.Index);
        Assert.Equal(2, request.ColumnIndex);
        Assert.Equal(1000, request.TimeoutMs);
        Assert.Equal(100, request.PollIntervalMs);
        Assert.Equal("raw", request.View);
        Assert.Equal("activeWindow", request.Root);
        Assert.Equal(5, request.MaxDepth);
        Assert.Equal(50, request.MaxChildren);
        Assert.True(request.IncludeOffscreen);
    }

    [Fact]
    public void UiResponse_SuccessAndFailureEnvelopes_RemainStable()
    {
        var ok = UiResponse.Ok(new { hello = "world" });
        Assert.True(ok.Success);
        Assert.NotNull(ok.Value);
        Assert.Null(ok.Error);

        var fail = UiResponse.Fail("broken", "shot.png");
        Assert.False(fail.Success);
        Assert.Equal("broken", fail.Error);
        Assert.Equal("shot.png", fail.ScreenshotPath);

        var fromFail = UiResponse.FromOperationResult(new { success = false, reason = "x", message = "y" });
        Assert.False(fromFail.Success);
        Assert.Equal("x", fromFail.Reason);
        Assert.Equal("y", fromFail.Error);
    }

    [Fact]
    public void RecordingExport_JsonRemainsBackwardCompatible()
    {
        var export = new RecordingExport
        {
            StartedAt = DateTimeOffset.Parse("2026-08-03T18:00:00Z"),
            StoppedAt = DateTimeOffset.Parse("2026-08-03T18:01:00Z"),
            Mode = "Assistive",
            ExportedFilePath = "recording.json",
            Screen = new ScreenResolutionInfo { Width = 800, Height = 600 },
            Actions =
            [
                new RecordedAction
                {
                    ActionType = ActionType.Click,
                    Mode = RecordingMode.Assistive,
                    Description = "Click",
                    Element = new ElementInfo { AutomationId = "btn", ControlType = "Button" }
                }
            ]
        };

        var json = JsonSerializer.Serialize(export, JsonOptions());
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        Assert.True(root.TryGetProperty("startedAt", out _));
        Assert.True(root.TryGetProperty("stoppedAt", out _));
        Assert.True(root.TryGetProperty("mode", out _));
        Assert.True(root.TryGetProperty("exportedFilePath", out _));
        Assert.True(root.TryGetProperty("screen", out _));
        Assert.True(root.TryGetProperty("actions", out var actions));
        Assert.Equal(JsonValueKind.Array, actions.ValueKind);
        var action = actions[0];
        Assert.True(action.TryGetProperty("actionType", out _));
        Assert.True(action.TryGetProperty("mode", out _));
        Assert.True(action.TryGetProperty("element", out _));
        Assert.True(action.TryGetProperty("description", out _));
    }

    [Fact]
    public void Playback_AcceptsDocumentedPayloadShapes()
    {
        var ui = new Mock<IUiService>();
        ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns((object?)null);
        var service = new PlaybackService(ui.Object, NullLogger<PlaybackService>.Instance);

        var shapes = new[]
        {
            """{ "actions": [ { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } } ] }""",
            """[ { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } } ]""",
            """{ "recording": { "actions": [ { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } } ] } }""",
            """{ "value": { "actions": [ { "actionType": "click", "mode": "assistive", "element": { "automationId": "a" } } ] } }"""
        };

        foreach (var shape in shapes)
        {
            var result = service.Play(JsonDocument.Parse(shape).RootElement);
            Assert.Equal(1, result.ExecutedActions);
        }
    }

    [Fact]
    public void UiOperationsEndpoint_DoesNotRequireSessionOrLaunch()
    {
        var ui = new Mock<IUiService>(MockBehavior.Strict);
        var config = new Mock<IConfiguration>();
        config.Setup(c => c["FailureScreenshotDirectory"]).Returns((string?)null);
        var controller = new UiController(
            ui.Object,
            new UiOperationCatalog(),
            NullLogger<UiController>.Instance,
            config.Object);

        var result = controller.GetOperations();
        Assert.IsType<OkObjectResult>(result);
        ui.VerifyNoOtherCalls();
    }

    private static IDriverContext MockDriverContext()
    {
        var mock = new Mock<IDriverContext>();
        mock.SetupGet(c => c.BearerToken).Returns(Token);
        return mock.Object;
    }

    private static JsonSerializerOptions JsonOptions() => new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}
