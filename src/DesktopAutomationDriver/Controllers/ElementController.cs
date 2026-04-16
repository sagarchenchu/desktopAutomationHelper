using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Handles all element-level automation operations within a session.
/// </summary>
[ApiController]
[Route("session/{sessionId}")]
public class ElementController : ControllerBase
{
    private readonly ISessionManager _sessionManager;
    private readonly IAutomationService _automationService;
    private readonly ILogger<ElementController> _logger;

    public ElementController(
        ISessionManager sessionManager,
        IAutomationService automationService,
        ILogger<ElementController> logger)
    {
        _sessionManager = sessionManager;
        _automationService = automationService;
        _logger = logger;
    }

    // ------------------------------------------------------------------ Element finding

    /// <summary>
    /// POST /session/{sessionId}/element
    /// Finds a single element in the application window.
    /// </summary>
    /// <remarks>
    /// Request body example:
    /// <code>
    /// { "using": "automation id", "value": "OKButton" }
    /// { "using": "name",          "value": "Save" }
    /// { "using": "class name",    "value": "Edit" }
    /// { "using": "tag name",      "value": "Button" }
    /// { "using": "xpath",         "value": "//Button[@AutomationId='okButton']" }
    /// </code>
    /// </remarks>
    [HttpPost("element")]
    public IActionResult FindElement(string sessionId, [FromBody] FindElementRequest request)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var elementId = _automationService.FindElement(session, request);
            return Ok(WebDriverResponse<ElementResponse>.Success(
                new ElementResponse { ElementId = elementId }, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/elements
    /// Finds all elements matching the given strategy.
    /// </summary>
    [HttpPost("elements")]
    public IActionResult FindElements(string sessionId, [FromBody] FindElementRequest request)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var elementIds = _automationService.FindElements(session, request);
            var responses = elementIds
                .Select(id => new ElementResponse { ElementId = id })
                .ToList();
            return Ok(WebDriverResponse<List<ElementResponse>>.Success(responses, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/element
    /// Finds a child element within the specified parent element.
    /// </summary>
    [HttpPost("element/{elementId}/element")]
    public IActionResult FindChildElement(string sessionId, string elementId, [FromBody] FindElementRequest request)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var childId = _automationService.FindElement(session, request, elementId);
            return Ok(WebDriverResponse<ElementResponse>.Success(
                new ElementResponse { ElementId = childId }, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/elements
    /// Finds all child elements within the specified parent element.
    /// </summary>
    [HttpPost("element/{elementId}/elements")]
    public IActionResult FindChildElements(string sessionId, string elementId, [FromBody] FindElementRequest request)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var childIds = _automationService.FindElements(session, request, elementId);
            var responses = childIds
                .Select(id => new ElementResponse { ElementId = id })
                .ToList();
            return Ok(WebDriverResponse<List<ElementResponse>>.Success(responses, sessionId));
        });
    }

    // ------------------------------------------------------------------ Actions

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/click
    /// Performs a mouse click on the element.
    /// </summary>
    [HttpPost("element/{elementId}/click")]
    public IActionResult Click(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            _automationService.Click(session, elementId);
            return Ok(WebDriverResponse<object?>.Success(null, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/doubleclick
    /// Performs a double-click on the element.
    /// </summary>
    [HttpPost("element/{elementId}/doubleclick")]
    public IActionResult DoubleClick(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            _automationService.DoubleClick(session, elementId);
            return Ok(WebDriverResponse<object?>.Success(null, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/rightclick
    /// Performs a right-click on the element.
    /// </summary>
    [HttpPost("element/{elementId}/rightclick")]
    public IActionResult RightClick(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            _automationService.RightClick(session, elementId);
            return Ok(WebDriverResponse<object?>.Success(null, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/value
    /// Types text into the element. Supports WebDriver special key codes.
    /// </summary>
    /// <remarks>
    /// Request body example:
    /// <code>
    /// { "value": ["Hello World", "\uE007"] }
    /// </code>
    /// The array items are concatenated / typed in sequence.
    /// "\uE007" is the Enter key.
    /// </remarks>
    [HttpPost("element/{elementId}/value")]
    public IActionResult SendKeys(string sessionId, string elementId, [FromBody] SendKeysRequest request)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            _automationService.SendKeys(session, elementId, request.Value);
            return Ok(WebDriverResponse<object?>.Success(null, sessionId));
        });
    }

    /// <summary>
    /// POST /session/{sessionId}/element/{elementId}/clear
    /// Clears the value of a text element.
    /// </summary>
    [HttpPost("element/{elementId}/clear")]
    public IActionResult Clear(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            _automationService.Clear(session, elementId);
            return Ok(WebDriverResponse<object?>.Success(null, sessionId));
        });
    }

    // ------------------------------------------------------------------ Queries

    /// <summary>
    /// GET /session/{sessionId}/element/{elementId}/text
    /// Returns the visible text of the element.
    /// </summary>
    [HttpGet("element/{elementId}/text")]
    public IActionResult GetText(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var text = _automationService.GetText(session, elementId);
            return Ok(WebDriverResponse<string>.Success(text, sessionId));
        });
    }

    /// <summary>
    /// GET /session/{sessionId}/element/{elementId}/name
    /// Returns the control type name of the element (e.g. "Button", "Edit").
    /// </summary>
    [HttpGet("element/{elementId}/name")]
    public IActionResult GetName(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var name = _automationService.GetControlType(session, elementId);
            return Ok(WebDriverResponse<string>.Success(name, sessionId));
        });
    }

    /// <summary>
    /// GET /session/{sessionId}/element/{elementId}/attribute/{name}
    /// Returns the named attribute/property of the element.
    /// Common attribute names: Name, AutomationId, ClassName, IsEnabled, IsOffscreen,
    /// ControlType, Value, HelpText, BoundingRectangle.
    /// </summary>
    [HttpGet("element/{elementId}/attribute/{name}")]
    public IActionResult GetAttribute(string sessionId, string elementId, string name)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var value = _automationService.GetAttribute(session, elementId, name);
            return Ok(WebDriverResponse<string?>.Success(value, sessionId));
        });
    }

    /// <summary>
    /// GET /session/{sessionId}/element/{elementId}/enabled
    /// Returns true if the element is enabled.
    /// </summary>
    [HttpGet("element/{elementId}/enabled")]
    public IActionResult IsEnabled(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var enabled = _automationService.IsEnabled(session, elementId);
            return Ok(WebDriverResponse<bool>.Success(enabled, sessionId));
        });
    }

    /// <summary>
    /// GET /session/{sessionId}/element/{elementId}/displayed
    /// Returns true if the element is visible on screen.
    /// </summary>
    [HttpGet("element/{elementId}/displayed")]
    public IActionResult IsDisplayed(string sessionId, string elementId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var displayed = _automationService.IsDisplayed(session, elementId);
            return Ok(WebDriverResponse<bool>.Success(displayed, sessionId));
        });
    }

    // ------------------------------------------------------------------ Screenshot

    /// <summary>
    /// GET /session/{sessionId}/screenshot
    /// Returns a Base64-encoded PNG screenshot of the application window.
    /// </summary>
    [HttpGet("screenshot")]
    public IActionResult TakeScreenshot(string sessionId)
    {
        return ExecuteWithSession(sessionId, session =>
        {
            var base64 = _automationService.TakeScreenshot(session);
            return Ok(WebDriverResponse<string>.Success(base64, sessionId));
        });
    }

    // ------------------------------------------------------------------ helpers

    private IActionResult ExecuteWithSession(string sessionId, Func<AutomationSession, IActionResult> action)
    {
        var session = _sessionManager.GetSession(sessionId);
        if (session == null)
            return NotFound(WebDriverResponse<ErrorDetail>.Error(
                6, $"Session '{sessionId}' not found.", "invalid session id"));

        try
        {
            return action(session);
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(ex, "Operation failed in session {SessionId}", sessionId);
            return NotFound(WebDriverResponse<ErrorDetail>.Error(
                7, ex.Message, "no such element"));
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid argument in session {SessionId}", sessionId);
            return BadRequest(WebDriverResponse<ErrorDetail>.Error(
                13, ex.Message, "invalid argument"));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error in session {SessionId}", sessionId);
            return StatusCode(500, WebDriverResponse<ErrorDetail>.Error(
                13, ex.Message, "unknown error"));
        }
    }
}
