using DesktopAutomationDriver.Models.Response;
using Microsoft.AspNetCore.Mvc;
using System.Reflection;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Exposes the /status endpoint which is used by clients to check if the
/// driver is running and ready to accept sessions.
/// </summary>
[ApiController]
[Route("status")]
public class StatusController : ControllerBase
{
    /// <summary>
    /// Returns the current driver status.
    /// </summary>
    [HttpGet]
    public IActionResult GetStatus()
    {
        var version = Assembly
            .GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "1.0.0";

        var response = WebDriverResponse<StatusResponse>.Success(new StatusResponse
        {
            Ready = true,
            Message = "Desktop Automation Driver is running",
            Build = new BuildInfo
            {
                Version = version,
                Time = DateTimeOffset.UtcNow.ToString("o")
            }
        });

        return Ok(response);
    }
}
