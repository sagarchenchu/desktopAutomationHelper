using DesktopAutomationDriver.Services.Debug;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>HTTP surface for Debug Mode start/stop/inspect. Feature logic lives in Debug services.</summary>
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
}
