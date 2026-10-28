using DesktopAutomationDriver.Models.Response;
using DesktopExecutionRuntime;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>Read-only diagnostics, protected by the existing bearer middleware.</summary>
[ApiController]
[Route("runtime")]
public sealed class RuntimeController(IDesktopEnvironmentProbe probe) : ControllerBase
{
    [HttpGet("diagnostics")]
    public IActionResult GetDiagnostics() =>
        Ok(WebDriverResponse<DesktopDiagnostics>.Success(DesktopDiagnostics.Evaluate(probe.Capture())));
}
