using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Provides endpoints to start, monitor and stop a window-activity recording session.
///
/// Workflow:
///   1. POST /record/start   → opens the transparent overlay.
///   2. Press Ctrl+P or Ctrl+A in the overlay to select a recording mode.
///   3. Press Ctrl+S (or call POST /record/stop) to finish and save the JSON.
///   4. GET  /record/actions → retrieve the recorded actions (or the exported file path).
/// </summary>
[ApiController]
[Route("record")]
public class RecordingController : ControllerBase
{
    private readonly IRecordingService _recordingService;
    private readonly ILogger<RecordingController> _logger;

    public RecordingController(IRecordingService recordingService, ILogger<RecordingController> logger)
    {
        _recordingService = recordingService;
        _logger = logger;
    }

    /// <summary>
    /// POST /record/start
    ///
    /// Opens a transparent always-on-top status bar at the top of the primary screen.
    /// Optionally launches an application and/or auto-stops after a given number of seconds.
    ///
    /// Request body (all fields optional):
    ///   exePath     – full path to an executable to launch before recording begins.
    ///   outputPath  – directory where the JSON export file should be saved.
    ///   waitSeconds – number of seconds before the recording auto-stops.
    ///
    /// The user can also interact with the overlay directly:
    ///   Ctrl+P  – Passive recording (mouse clicks / keyboard captured automatically)
    ///   Ctrl+A  – Assistive recording (right-click an element to choose an action)
    ///   Ctrl+S  – Stop and export to JSON
    /// </summary>
    [HttpPost("start")]
    public IActionResult Start([FromBody] StartRecordingRequest? request = null)
    {
        try
        {
            var result = _recordingService.StartRecording(request);
            if (!string.IsNullOrEmpty(result.Error))
                return Conflict(WebDriverResponse<ErrorDetail>.Error(9, result.Error, "recording already active"));

            return Ok(new
            {
                success = true,
                message = "Recording started.",
                launch = result.Launch,
                outputPath = result.OutputPath
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error starting recording");
            return StatusCode(500, WebDriverResponse<ErrorDetail>.Error(13, ex.Message, "unknown error"));
        }
    }

    /// <summary>
    /// GET /record/status
    ///
    /// Returns whether recording is active, the current mode and the number of actions captured so far.
    /// </summary>
    [HttpGet("status")]
    public IActionResult Status()
    {
        return Ok(WebDriverResponse<object>.Success(new
        {
            isActive = _recordingService.IsActive,
            mode = _recordingService.CurrentMode.ToString(),
            startedAt = _recordingService.StartedAt,
            actionsCount = _recordingService.GetCurrentState().Actions.Count
        }));
    }

    /// <summary>
    /// GET /record/actions
    ///
    /// Returns all recorded actions collected so far (or after recording stopped).
    /// The response also includes the path of the exported JSON file, if available.
    /// </summary>
    [HttpGet("actions")]
    public IActionResult GetActions()
    {
        var state = _recordingService.GetCurrentState();
        return Ok(WebDriverResponse<RecordingExport>.Success(state));
    }

    /// <summary>
    /// POST /record/stop
    ///
    /// Stops the active recording session, writes the JSON export file and returns the result.
    /// Idempotent — safe to call even if recording was already stopped (e.g. via Ctrl+S).
    /// </summary>
    [HttpPost("stop")]
    public IActionResult Stop()
    {
        try
        {
            var export = _recordingService.StopRecording();
            return Ok(WebDriverResponse<RecordingExport>.Success(export));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error stopping recording");
            return StatusCode(500, WebDriverResponse<ErrorDetail>.Error(13, ex.Message, "unknown error"));
        }
    }
}
