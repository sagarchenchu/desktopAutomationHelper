using DesktopAutomationDriver.Services.Debug;
using DesktopAutomationDriver.Services.Vision;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Config-gated production surface for vision fallback (default off via Vision:Enabled=false).
/// Exposes candidate production, coordinate authorization, hover, and double-click policy checks.
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
        public required DebugWindowIdentity Window { get; set; }
        public DebugUiaNode? SemanticTarget { get; set; }
        public string? PngBase64 { get; set; }
        public bool UiaReportsInaccessible { get; set; }
        public bool PolicyOverride { get; set; }
        public string? Query { get; set; }
    }

    /// <summary>
    /// Production path for vision candidates. Returns empty when Vision:Enabled=false.
    /// </summary>
    [HttpPost("candidates")]
    public IActionResult ProduceCandidates([FromBody] ProduceCandidatesRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Window);

        byte[] png;
        if (!string.IsNullOrWhiteSpace(request.PngBase64))
            png = Convert.FromBase64String(request.PngBase64);
        else
            png = System.Text.Encoding.UTF8.GetBytes($"vision-candidates:{request.Window.WindowHandleHex}");

        var semantic = request.SemanticTarget;
        if (semantic is null && request.UiaReportsInaccessible)
        {
            semantic = new DebugUiaNode
            {
                ControlType = "Custom",
                IsSemanticAccessible = false,
                InaccessibilityReason = "caller-reported-inaccessible"
            };
        }

        var (eligibility, _, candidates) = _coordinator.TryProduceCandidates(
            semantic,
            request.Window,
            png,
            query: request.Query,
            policyExplicitlyPermits: request.PolicyOverride);

        return Ok(new
        {
            enabled = _options.Enabled,
            eligibility = eligibility.ToString(),
            candidates,
            autoPromoted = false,
            requiresHumanApproval = true
        });
    }

    public sealed class AuthorizeRequest
    {
        public required VisionCandidateTarget Candidate { get; set; }
        public required DebugWindowIdentity CurrentWindow { get; set; }
        public string Action { get; set; } = "click";
        public bool IsDestructive { get; set; }
        public DateTimeOffset? NowUtc { get; set; }
    }

    [HttpPost("authorize")]
    public IActionResult Authorize([FromBody] AuthorizeRequest request)
    {
        var decision = _coordinator.AuthorizeCoordinateAction(new CoordinateActionRequest
        {
            Candidate = request.Candidate,
            CurrentWindow = request.CurrentWindow,
            Action = request.Action,
            IsDestructive = request.IsDestructive,
            NowUtc = request.NowUtc ?? DateTimeOffset.UtcNow
        });
        return Ok(decision);
    }

    /// <summary>Gated hover authorization (same fail-closed guard; requires Vision:Enabled).</summary>
    [HttpPost("hover")]
    public IActionResult AuthorizeHover([FromBody] AuthorizeRequest request)
    {
        var decision = _coordinator.AuthorizeGatedPointerAction(new CoordinateActionRequest
        {
            Candidate = request.Candidate,
            CurrentWindow = request.CurrentWindow,
            Action = "hover",
            IsDestructive = request.IsDestructive,
            NowUtc = request.NowUtc ?? DateTimeOffset.UtcNow
        });
        return Ok(decision);
    }

    public sealed class DoubleClickRequest
    {
        public string? ControlType { get; set; }
        public bool PolicyAllowsDoubleClick { get; set; }
        public bool KnownBehavior { get; set; }
        public VisionCandidateTarget? Candidate { get; set; }
        public DebugWindowIdentity? CurrentWindow { get; set; }
    }

    /// <summary>
    /// Double-click requires DoubleClickPolicy and (when a vision candidate is supplied) Vision:Enabled + guard.
    /// </summary>
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

        if (request.Candidate is null || request.CurrentWindow is null)
        {
            return Ok(new
            {
                allowed = policyAllowed && _options.Enabled,
                reason = _options.Enabled
                    ? "Double-click policy allowed (no vision candidate to authorize)."
                    : "Vision fallback disabled (Vision:Enabled=false).",
                visionEnabled = _options.Enabled,
                policyAllowed
            });
        }

        var decision = _coordinator.AuthorizeGatedPointerAction(new CoordinateActionRequest
        {
            Candidate = request.Candidate,
            CurrentWindow = request.CurrentWindow,
            Action = "doubleclick",
            NowUtc = DateTimeOffset.UtcNow
        });

        return Ok(new
        {
            allowed = decision.Allowed,
            reason = decision.Reason,
            screenPoint = decision.ScreenPoint,
            visionEnabled = _options.Enabled,
            policyAllowed
        });
    }
}
