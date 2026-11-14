using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Bootstrap;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Discovery endpoint. In <see cref="BootstrapDiscoveryMode.Secure"/> (default)
/// this never returns the bearer token — clients must use the SID ACL-protected
/// named pipe / connection descriptor. Legacy token return requires an explicit
/// <see cref="BootstrapDiscoveryMode.LegacyVerify"/> opt-in with risk acknowledgement.
/// </summary>
[ApiController]
[Route("verify")]
public class VerifyController : ControllerBase
{
    private readonly IDriverContext _driverContext;
    private readonly BootstrapOptions _bootstrap;

    public VerifyController(
        IDriverContext driverContext,
        IOptions<BootstrapOptions> bootstrapOptions)
    {
        _driverContext = driverContext;
        _bootstrap = bootstrapOptions?.Value ?? new BootstrapOptions();
    }

    /// <summary>GET /verify</summary>
    [HttpGet]
    public IActionResult Verify()
    {
        if (!_bootstrap.AllowsLegacyVerifyHttp)
        {
            return StatusCode(StatusCodes.Status403Forbidden, new
            {
                status = 403,
                value = new
                {
                    error = "secure_bootstrap_required",
                    message =
                        "Unauthenticated HTTP /verify does not return the bearer token in secure mode. " +
                        "Use the per-user SID ACL-protected named pipe or connection descriptor. " +
                        "Legacy /verify token return requires Bootstrap:Mode=LegacyVerify and " +
                        "Bootstrap:AcknowledgeLegacyVerifyRisk=true (unsafe on shared hosts)."
                }
            });
        }

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
