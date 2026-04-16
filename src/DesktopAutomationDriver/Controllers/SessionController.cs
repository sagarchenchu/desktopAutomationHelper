using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Handles session lifecycle: create, get and delete automation sessions.
/// </summary>
[ApiController]
[Route("session")]
public class SessionController : ControllerBase
{
    private readonly ISessionManager _sessionManager;
    private readonly ILogger<SessionController> _logger;

    public SessionController(ISessionManager sessionManager, ILogger<SessionController> logger)
    {
        _sessionManager = sessionManager;
        _logger = logger;
    }

    /// <summary>
    /// POST /session
    /// Creates a new automation session by launching an application or attaching
    /// to an already-running process.
    /// </summary>
    /// <remarks>
    /// Request body example (launch):
    /// <code>
    /// {
    ///   "desiredCapabilities": {
    ///     "app": "C:\\Windows\\System32\\notepad.exe",
    ///     "uiaType": "UIA3",
    ///     "launchDelay": 1000
    ///   }
    /// }
    /// </code>
    /// Request body example (attach):
    /// <code>
    /// {
    ///   "desiredCapabilities": {
    ///     "appName": "notepad",
    ///     "uiaType": "UIA3"
    ///   }
    /// }
    /// </code>
    /// </remarks>
    [HttpPost]
    public IActionResult CreateSession([FromBody] CreateSessionRequest request)
    {
        try
        {
            var session = _sessionManager.CreateSession(request.DesiredCapabilities);
            var response = WebDriverResponse<SessionResponse>.Success(
                new SessionResponse
                {
                    SessionId = session.SessionId,
                    Capabilities = new SessionCapabilities
                    {
                        App = session.AppPath,
                        AppName = session.AppName,
                        UiaType = session.UiaType,
                        ProcessId = session.Application.ProcessId
                    }
                },
                session.SessionId);

            return Ok(response);
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid session creation request");
            return BadRequest(WebDriverResponse<ErrorDetail>.Error(13, ex.Message, "invalid argument"));
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(ex, "Could not create session");
            return UnprocessableEntity(
                WebDriverResponse<ErrorDetail>.Error(33, ex.Message, "no such window"));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error creating session");
            return StatusCode(500,
                WebDriverResponse<ErrorDetail>.Error(13, ex.Message, "unknown error"));
        }
    }

    /// <summary>
    /// GET /session/{sessionId}
    /// Returns information about the session.
    /// </summary>
    [HttpGet("{sessionId}")]
    public IActionResult GetSession(string sessionId)
    {
        var session = _sessionManager.GetSession(sessionId);
        if (session == null)
            return NotFound(WebDriverResponse<ErrorDetail>.Error(6, $"Session '{sessionId}' not found.", "invalid session id"));

        return Ok(WebDriverResponse<SessionCapabilities>.Success(
            new SessionCapabilities
            {
                App = session.AppPath,
                AppName = session.AppName,
                UiaType = session.UiaType,
                ProcessId = session.Application.ProcessId
            },
            sessionId));
    }

    /// <summary>
    /// DELETE /session/{sessionId}
    /// Closes the session and optionally terminates the application.
    /// </summary>
    [HttpDelete("{sessionId}")]
    public IActionResult DeleteSession(string sessionId)
    {
        var session = _sessionManager.GetSession(sessionId);
        if (session == null)
            return NotFound(WebDriverResponse<ErrorDetail>.Error(6, $"Session '{sessionId}' not found.", "invalid session id"));

        _sessionManager.CloseSession(sessionId);
        return Ok(WebDriverResponse<object?>.Success(null, sessionId));
    }

    /// <summary>
    /// GET /sessions
    /// Lists all active session IDs.
    /// </summary>
    [HttpGet("/sessions")]
    public IActionResult ListSessions()
    {
        var ids = _sessionManager.GetAllSessionIds();
        return Ok(WebDriverResponse<IReadOnlyList<string>>.Success(ids));
    }
}
