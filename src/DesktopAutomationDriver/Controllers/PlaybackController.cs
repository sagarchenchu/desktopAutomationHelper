using System.Text.Json;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Replays Assistive recording JSON by executing its recorded actions in order.
/// </summary>
[ApiController]
[Route("playback")]
public class PlaybackController : ControllerBase
{
    private readonly IPlaybackService _playbackService;
    private readonly ILogger<PlaybackController> _logger;

    public PlaybackController(IPlaybackService playbackService, ILogger<PlaybackController> logger)
    {
        _playbackService = playbackService;
        _logger = logger;
    }

    /// <summary>
    /// POST /playback
    ///
    /// Body can be the raw recording export JSON returned by /record/stop or
    /// /record/actions, an actions array, or { "recording": { ... }, "continueOnError": true }.
    /// </summary>
    [HttpPost]
    public IActionResult Play([FromBody] JsonElement payload)
    {
        try
        {
            var result = _playbackService.Play(payload);
            return Ok(WebDriverResponse<object>.Success(result));
        }
        catch (ArgumentException ex)
        {
            return BadRequest(WebDriverResponse<ErrorDetail>.Error(10, ex.Message, "invalid argument"));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Playback failed");
            return StatusCode(500, WebDriverResponse<ErrorDetail>.Error(13, ex.Message, "unknown error"));
        }
    }
}
