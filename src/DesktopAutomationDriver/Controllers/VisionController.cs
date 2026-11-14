using DesktopAutomationDriver.Services.Debug;
using DesktopAutomationDriver.Services.Vision;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Config-gated vision surface. Candidates are server-captured; approval uses the operator channel
/// (interactive presence); authorize consumes a short-lived one-use execution token.
/// </summary>
[ApiController]
[Route("vision")]
public sealed class VisionController : ControllerBase
{
    private readonly VisionFallbackOptions _options;
    private readonly VisionFallbackCoordinator _coordinator;
    private readonly VisionFallbackPolicy _policy;

    public VisionController(
        VisionFallbackOptions options,
        VisionFallbackCoordinator coordinator,
        VisionFallbackPolicy policy)
    {
        _options = options;
        _coordinator = coordinator;
        _policy = policy;
    }

    [HttpGet("status")]
    public IActionResult Status() =>
        Ok(new
        {
            enabled = _options.Enabled,
            provider = _options.Provider,
            allowWhenUiaInaccessible = _options.AllowWhenUiaInaccessible,
            allowWhenPolicyPermits = _options.AllowWhenPolicyPermits,
            maxScreenshotAgeMilliseconds = _options.MaxScreenshotAgeMilliseconds,
            candidateTtlMilliseconds = _options.CandidateTtlMilliseconds,
            executionTokenTtlMilliseconds = _options.ExecutionTokenTtlMilliseconds,
            maxStoredCandidates = _options.MaxStoredCandidates,
            minimumConfidence = _options.MinimumConfidence
        });

    public sealed class EvaluateRequest
    {
        public DebugUiaNode? SemanticTarget { get; set; }
        public bool PolicyOverride { get; set; }
        public bool UiaReportsInaccessible { get; set; }
    }

    [HttpPost("evaluate")]
    public IActionResult Evaluate([FromBody] EvaluateRequest request)
    {
        VisionEligibility eligibility;
        if (request.SemanticTarget is not null)
            eligibility = _policy.Evaluate(request.SemanticTarget, request.PolicyOverride);
        else
            eligibility = _coordinator.EvaluateEligibility(request.UiaReportsInaccessible, request.PolicyOverride);

        return Ok(new
        {
            enabled = _options.Enabled,
            eligibility = eligibility.ToString()
        });
    }

    public sealed class ProduceCandidatesRequest
    {
        /// <summary>HWND hex (e.g. 0x1234ABCD) or decimal. Server probes this handle — client window/PNG ignored.</summary>
        public string? WindowHandle { get; set; }
        public int? CursorX { get; set; }
        public int? CursorY { get; set; }
        public bool PolicyOverride { get; set; }
        public string? Query { get; set; }
    }

    /// <summary>
    /// Server-side capture from HWND (+ optional cursor). Client PNG/window/UIA flags are not trusted.
    /// </summary>
    [HttpPost("candidates")]
    public IActionResult ProduceCandidates([FromBody] ProduceCandidatesRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (!TryParseHandle(request.WindowHandle, out var hwnd) || hwnd == IntPtr.Zero)
            return BadRequest(new { error = "windowHandle is required (server captures window/screenshot/UIA)." });

        Point? cursor = request.CursorX is int x && request.CursorY is int y
            ? new Point(x, y)
            : null;

        var (eligibility, _, candidates) = _coordinator.TryProduceCandidatesFromHandle(
            hwnd,
            cursor,
            query: request.Query,
            policyExplicitlyPermits: request.PolicyOverride);

        return Ok(new
        {
            enabled = _options.Enabled,
            eligibility = eligibility.ToString(),
            candidates,
            autoPromoted = false,
            requiresHumanApproval = true,
            requiresInteractivePresence = true,
            requiresOperatorToken = false
        });
    }

    public sealed class ApproveRequest
    {
        public required string CandidateId { get; set; }
        /// <summary>Action bound to this approval (click/hover/doubleclick).</summary>
        public string Action { get; set; } = "click";
    }

    /// <summary>
    /// Operator approval: bearer auth + driver-owned interactive presence
    /// (ACL file self-read is not human approval). Revalidates pixels against the
    /// original baseline and returns a short-lived one-use execution token.
    /// </summary>
    [HttpPost("approve")]
    public IActionResult Approve([FromBody] ApproveRequest request)
    {
        var result = _coordinator.ApproveCandidate(request.CandidateId, request.Action);
        if (!result.Allowed)
        {
            return BadRequest(new
            {
                allowed = false,
                reason = result.Reason,
                candidateId = request.CandidateId
            });
        }

        return Ok(new
        {
            allowed = true,
            reason = result.Reason,
            candidateId = result.CandidateId,
            action = result.Action,
            executionToken = result.ExecutionToken,
            executionTokenExpiresAtUtc = result.ExecutionTokenExpiresAtUtc,
            screenPoint = result.ScreenPoint,
            oneUse = true
        });
    }

    public sealed class AuthorizeRequest
    {
        /// <summary>Short-lived one-use token from operator approve.</summary>
        public required string ExecutionToken { get; set; }
        public string? Action { get; set; }
        public bool IsDestructive { get; set; }
    }

    [HttpPost("authorize")]
    public IActionResult Authorize([FromBody] AuthorizeRequest request)
    {
        var decision = _coordinator.AuthorizeCoordinateAction(
            request.ExecutionToken,
            request.Action,
            request.IsDestructive);
        return Ok(new
        {
            decision.Allowed,
            decision.Reason,
            decision.ScreenPoint,
            visionEnabled = _options.Enabled,
            consumed = decision.Allowed
        });
    }

    [HttpPost("hover")]
    public IActionResult AuthorizeHover([FromBody] AuthorizeRequest request)
    {
        var decision = _coordinator.AuthorizeGatedPointerAction(
            request.ExecutionToken,
            "hover",
            request.IsDestructive);
        return Ok(new
        {
            decision.Allowed,
            decision.Reason,
            decision.ScreenPoint,
            visionEnabled = _options.Enabled,
            consumed = decision.Allowed
        });
    }

    public sealed class DoubleClickRequest
    {
        public string? ControlType { get; set; }
        public bool PolicyAllowsDoubleClick { get; set; }
        public bool KnownBehavior { get; set; }
        public string? ExecutionToken { get; set; }
    }

    [HttpPost("doubleclick")]
    public IActionResult AuthorizeDoubleClick([FromBody] DoubleClickRequest request)
    {
        var policyAllowed = DoubleClickPolicy.IsAllowed(
            request.ControlType,
            request.PolicyAllowsDoubleClick,
            request.KnownBehavior);

        if (!policyAllowed)
        {
            return Ok(new
            {
                allowed = false,
                reason = DoubleClickPolicy.DescribeDenial(request.ControlType),
                visionEnabled = _options.Enabled
            });
        }

        if (string.IsNullOrWhiteSpace(request.ExecutionToken))
        {
            return Ok(new
            {
                allowed = false,
                reason = "Double-click with vision requires an operator executionToken.",
                visionEnabled = _options.Enabled,
                policyAllowed
            });
        }

        var decision = _coordinator.AuthorizeGatedPointerAction(request.ExecutionToken, "doubleclick");
        return Ok(new
        {
            allowed = decision.Allowed,
            reason = decision.Reason,
            screenPoint = decision.ScreenPoint,
            visionEnabled = _options.Enabled,
            policyAllowed,
            consumed = decision.Allowed
        });
    }

    private static bool TryParseHandle(string? value, out IntPtr hwnd)
    {
        hwnd = IntPtr.Zero;
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var trimmed = value.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            if (long.TryParse(trimmed[2..], System.Globalization.NumberStyles.HexNumber, null, out var hex))
            {
                hwnd = new IntPtr(hex);
                return true;
            }

            return false;
        }

        if (long.TryParse(trimmed, out var dec))
        {
            hwnd = new IntPtr(dec);
            return true;
        }

        return false;
    }
}
