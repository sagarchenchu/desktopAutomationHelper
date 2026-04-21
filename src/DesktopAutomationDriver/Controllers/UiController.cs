using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Unified entry point for all UI automation operations.
///
/// POST (or GET) http://127.0.0.1:{port}/ui
///
/// Request body:
/// <code>
/// {
///   "operation": "&lt;operation-name&gt;",
///   "locator": { "name": "...", "automationId": "...", "className": "...", "controlType": "..." },
///   "locator2": { "name": "...", "automationId": "..." },
///   "value": "...",
///   "index": 0
/// }
/// </code>
///
/// All routes require: Authorization: Bearer &lt;token&gt;
/// </summary>
[ApiController]
[Route("ui")]
public class UiController : ControllerBase
{
    private readonly IUiService _uiService;
    private readonly ILogger<UiController> _logger;

    public UiController(IUiService uiService, ILogger<UiController> logger)
    {
        _uiService = uiService;
        _logger = logger;
    }

    /// <summary>
    /// Executes the requested UI automation operation.
    /// </summary>
    [HttpPost]
    [HttpGet]
    public IActionResult Execute([FromBody] UiRequest? request)
    {
        if (request == null)
            return BadRequest(UiResponse.Fail("Request body is required."));

        try
        {
            var result = _uiService.Execute(request);
            return Ok(UiResponse.Ok(result));
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid argument for operation '{Op}'",
                SanitizeOp(request.Operation));
            var screenshot = _uiService.TakeFailureScreenshot(FailureScreenshotDirectory);
            return BadRequest(UiResponse.Fail(ex.Message, screenshot));
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(ex, "Operation failed: '{Op}'",
                SanitizeOp(request.Operation));
            var screenshot = _uiService.TakeFailureScreenshot(FailureScreenshotDirectory);
            return NotFound(UiResponse.Fail(ex.Message, screenshot));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error for operation '{Op}'",
                SanitizeOp(request.Operation));
            var screenshot = _uiService.TakeFailureScreenshot(FailureScreenshotDirectory);
            return StatusCode(500, UiResponse.Fail(ex.Message, screenshot));
        }
    }

    /// <summary>
    /// Directory where failure screenshots are automatically saved.
    /// Resolves to <c>test/resources</c> relative to the current working directory.
    /// </summary>
    private static readonly string FailureScreenshotDirectory =
        Path.Combine(Directory.GetCurrentDirectory(), "test", "resources");

    /// <summary>
    /// Strips control characters from operation names before logging
    /// to prevent log-injection attacks.
    /// </summary>
    private static string SanitizeOp(string op) =>
        System.Text.RegularExpressions.Regex.Replace(op ?? string.Empty, @"[\r\n\t]", "_");
}
