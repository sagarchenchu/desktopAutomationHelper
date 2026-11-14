using DesktopAutomationDriver.Services.Debug;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>HTTP surface for Debug Mode start/stop/inspect/hover. Feature logic lives in Debug services.</summary>
[ApiController]
[Route("debug")]
public sealed class DebugController : ControllerBase
{
    private readonly IDebugSessionCoordinator _coordinator;

    public DebugController(IDebugSessionCoordinator coordinator)
    {
        _coordinator = coordinator;
    }

    [HttpGet("status")]
    public IActionResult Status() => Ok(_coordinator.Snapshot);

    [HttpPost("start")]
    public IActionResult Start([FromQuery] string? outputDirectory = null) =>
        Ok(_coordinator.Start(outputDirectory));

    [HttpPost("stop")]
    public IActionResult Stop() => Ok(_coordinator.Stop());

    [HttpPost("toggle")]
    public IActionResult Toggle([FromQuery] string? outputDirectory = null) =>
        Ok(_coordinator.Toggle(outputDirectory));

    public sealed class InspectRequest
    {
        public int X { get; set; }
        public int Y { get; set; }
        public bool WriteArtifact { get; set; }
    }

    [HttpPost("inspect")]
    public IActionResult Inspect([FromBody] InspectRequest request)
    {
        if (!_coordinator.IsActive)
            return Conflict(new { error = "Debug Mode is not active. POST /debug/start first." });

        return Ok(_coordinator.InspectUnderCursor(new Point(request.X, request.Y), request.WriteArtifact));
    }

    [HttpPost("verify")]
    public IActionResult Verify()
    {
        if (!_coordinator.IsActive)
            return Conflict(new { error = "Debug Mode is not active." });
        return Ok(_coordinator.VerifyLastCandidate());
    }

    public sealed class PointRequest
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    /// <summary>Hover is a mouse probe (not a UIA capability). Requires active Debug Mode.</summary>
    [HttpPost("hover")]
    public IActionResult Hover([FromBody] PointRequest request)
    {
        if (!_coordinator.IsActive)
            return Conflict(new { error = "Debug Mode is not active. POST /debug/start first." });

        return Ok(_coordinator.ProbeHover(new Point(request.X, request.Y)));
    }

    public sealed class DoubleClickPolicyRequest
    {
        public string? ControlType { get; set; }
        public bool PolicyAllowsDoubleClick { get; set; }
        public bool KnownBehavior { get; set; }
    }

    /// <summary>Double-click policy gate exposed for Debug Mode / acceptance (fail-closed).</summary>
    [HttpPost("doubleclick-allowed")]
    public IActionResult DoubleClickAllowed([FromBody] DoubleClickPolicyRequest request)
    {
        var allowed = DoubleClickPolicy.IsAllowed(
            request.ControlType,
            request.PolicyAllowsDoubleClick,
            request.KnownBehavior);
        return Ok(new
        {
            allowed,
            reason = allowed
                ? "Double-click allowed by policy or known behavior."
                : DoubleClickPolicy.DescribeDenial(request.ControlType)
        });
    }
}
