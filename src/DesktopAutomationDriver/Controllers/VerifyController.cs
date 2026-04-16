using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Unauthenticated endpoint that tells callers whether this driver instance
/// is running and provides the connection details (port + Bearer token)
/// needed to use the automation endpoints.
///
/// This endpoint is intentionally exempt from Bearer-token authentication
/// so clients can bootstrap without knowing the token in advance.
/// </summary>
[ApiController]
[Route("verify")]
public class VerifyController : ControllerBase
{
    private readonly IDriverContext _driverContext;

    public VerifyController(IDriverContext driverContext)
    {
        _driverContext = driverContext;
    }

    /// <summary>
    /// GET /verify
    /// Returns driver status, connection port, and the Bearer token required
    /// for all other endpoints.
    /// </summary>
    [HttpGet]
    public IActionResult Verify()
    {
        return Ok(WebDriverResponse<VerifyResponse>.Success(new VerifyResponse
        {
            Running = true,
            Username = _driverContext.Username,
            Port = _driverContext.MainPort,
            ProbePort = _driverContext.ProbePortActive ? _driverContext.ProbePort : null,
            Token = _driverContext.BearerToken,
            AuthorizationHeader = $"Bearer {_driverContext.BearerToken}"
        }));
    }
}
