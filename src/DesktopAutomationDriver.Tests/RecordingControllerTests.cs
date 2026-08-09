using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class RecordingControllerTests
{
    private readonly Mock<IRecordingService> _recording = new(MockBehavior.Strict);
    private readonly RecordingController _controller;

    public RecordingControllerTests()
    {
        _controller = new RecordingController(_recording.Object, NullLogger<RecordingController>.Instance);
    }

    [Fact]
    public void Start_WhenSuccessful_ReturnsOkWithExportFields()
    {
        _recording.Setup(s => s.StartRecording(It.IsAny<StartRecordingRequest?>()))
            .Returns(new StartRecordingResult
            {
                OutputPath = "C:/temp/recordings",
                Launch = new LaunchInfo { Success = true, ProcessId = 42, WindowTitle = "App" }
            });

        var result = _controller.Start(new StartRecordingRequest { ExePath = "C:/app.exe" });

        var ok = Assert.IsType<OkObjectResult>(result);
        var json = System.Text.Json.JsonSerializer.Serialize(ok.Value);
        Assert.Contains("\"success\":true", json);
        Assert.Contains("Recording started.", json);
        Assert.Contains("outputPath", json);
        Assert.Contains("launch", json);
        _recording.VerifyAll();
    }

    [Fact]
    public void Start_WhenAlreadyActive_Returns409()
    {
        _recording.Setup(s => s.StartRecording(It.IsAny<StartRecordingRequest?>()))
            .Returns(new StartRecordingResult { Error = "Recording is already active." });

        var result = _controller.Start(null);

        var conflict = Assert.IsType<ConflictObjectResult>(result);
        Assert.Equal(409, conflict.StatusCode);
        _recording.VerifyAll();
    }

    [Fact]
    public void Start_WhenServiceThrows_Returns500()
    {
        _recording.Setup(s => s.StartRecording(It.IsAny<StartRecordingRequest?>()))
            .Throws(new InvalidOperationException("overlay failed"));

        var result = _controller.Start(null);

        var obj = Assert.IsType<ObjectResult>(result);
        Assert.Equal(500, obj.StatusCode);
    }

    [Fact]
    public void Status_ReturnsRecordingState()
    {
        _recording.SetupGet(s => s.IsActive).Returns(true);
        _recording.SetupGet(s => s.CurrentMode).Returns(RecordingMode.Assistive);
        _recording.SetupGet(s => s.StartedAt).Returns(DateTimeOffset.Parse("2026-08-03T18:00:00Z"));
        _recording.Setup(s => s.GetCurrentState()).Returns(new RecordingExport
        {
            Mode = "Assistive",
            Actions =
            [
                new RecordedAction { ActionType = ActionType.Click, Mode = RecordingMode.Assistive }
            ]
        });

        var result = _controller.Status();
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = System.Text.Json.JsonSerializer.Serialize(ok.Value);
        Assert.Contains("\"isActive\":true", json);
        Assert.Contains("Assistive", json);
        Assert.Contains("\"actionsCount\":1", json);
    }

    [Fact]
    public void GetActions_ReturnsRecordingExportEnvelope()
    {
        var export = new RecordingExport
        {
            StartedAt = DateTimeOffset.Parse("2026-08-03T18:00:00Z"),
            Mode = "Assistive",
            ExportedFilePath = null,
            Screen = new ScreenResolutionInfo { Width = 1920, Height = 1080 },
            Actions = []
        };
        _recording.Setup(s => s.GetCurrentState()).Returns(export);

        var result = _controller.GetActions();
        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<WebDriverResponse<RecordingExport>>(ok.Value);
        Assert.Equal(0, response.Status);
        Assert.NotNull(response.Value);
        Assert.Equal("Assistive", response.Value!.Mode);
        Assert.NotNull(response.Value.Actions);
        Assert.NotNull(response.Value.Screen);
        Assert.Equal(export.StartedAt, response.Value.StartedAt);
    }

    [Fact]
    public void Stop_WhenSuccessful_ReturnsExport()
    {
        var export = new RecordingExport
        {
            StartedAt = DateTimeOffset.Parse("2026-08-03T18:00:00Z"),
            StoppedAt = DateTimeOffset.Parse("2026-08-03T18:05:00Z"),
            Mode = "Assistive",
            ExportedFilePath = "C:/temp/recording.json",
            Actions = [new RecordedAction { ActionType = ActionType.Type, Mode = RecordingMode.Assistive, Value = "x" }]
        };
        _recording.Setup(s => s.StopRecording()).Returns(export);

        var result = _controller.Stop();
        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<WebDriverResponse<RecordingExport>>(ok.Value);
        Assert.Equal(export.ExportedFilePath, response.Value!.ExportedFilePath);
        Assert.Equal(export.StoppedAt, response.Value.StoppedAt);
        Assert.Single(response.Value.Actions);
    }

    [Fact]
    public void Stop_WhenServiceThrows_Returns500()
    {
        _recording.Setup(s => s.StopRecording()).Throws(new Exception("disk full"));

        var result = _controller.Stop();
        var obj = Assert.IsType<ObjectResult>(result);
        Assert.Equal(500, obj.StatusCode);
    }
}
