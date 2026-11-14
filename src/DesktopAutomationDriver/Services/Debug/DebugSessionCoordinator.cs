using DesktopAutomationDriver.Services.Vision;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Debug Mode session coordinator. Keeps feature logic out of RecordingOverlayWindow / UiService.
/// When Vision:Enabled and UIA reports inaccessible, produces vision candidates via the gated coordinator.
/// </summary>
public sealed class DebugSessionCoordinator : IDebugSessionCoordinator, IDisposable
{
    private readonly UiaTreeCaptureService _treeCapture;
    private readonly LocatorCandidateGenerator _candidateGenerator;
    private readonly NonInteractiveHighlightService _highlight;
    private readonly CandidateVerificationService _verification;
    private readonly DebugArtifactWriter _artifactWriter;
    private readonly IUiaElementProbe _probe;
    private readonly VisionFallbackCoordinator? _vision;
    private readonly VisionScreenshotService? _screenshots;
    private readonly ILogger<DebugSessionCoordinator> _logger;
    private readonly object _gate = new();

    private bool _active;
    private string? _sessionId;
    private DateTimeOffset? _startedAtUtc;
    private string? _outputDirectory;
    private DebugTreeCapture? _lastCapture;
    private IReadOnlyList<LocatorCandidate> _lastCandidates = [];
    private CandidateVerificationResult? _lastVerification;
    private IReadOnlyList<VisionCandidateTarget> _lastVisionCandidates = [];
    private string? _lastVisionEligibility;
    private string? _artifactPath;
    private bool _disposed;

    public DebugSessionCoordinator(
        UiaTreeCaptureService treeCapture,
        LocatorCandidateGenerator candidateGenerator,
        NonInteractiveHighlightService highlight,
        CandidateVerificationService verification,
        DebugArtifactWriter artifactWriter,
        IUiaElementProbe probe,
        VisionFallbackCoordinator? vision = null,
        VisionScreenshotService? screenshots = null,
        ILogger<DebugSessionCoordinator>? logger = null)
    {
        _treeCapture = treeCapture ?? throw new ArgumentNullException(nameof(treeCapture));
        _candidateGenerator = candidateGenerator ?? throw new ArgumentNullException(nameof(candidateGenerator));
        _highlight = highlight ?? throw new ArgumentNullException(nameof(highlight));
        _verification = verification ?? throw new ArgumentNullException(nameof(verification));
        _artifactWriter = artifactWriter ?? throw new ArgumentNullException(nameof(artifactWriter));
        _probe = probe ?? throw new ArgumentNullException(nameof(probe));
        _vision = vision;
        _screenshots = screenshots;
        _logger = logger ?? NullLogger<DebugSessionCoordinator>.Instance;
    }

    public static DebugSessionCoordinator ForTests(
        IUiaElementProbe? probe = null,
        IHighlightSurface? highlightSurface = null,
        VisionFallbackCoordinator? vision = null,
        VisionScreenshotService? screenshots = null)
    {
        probe ??= new NullUiaElementProbe();
        highlightSurface ??= new FakeHighlightSurface();
        return new DebugSessionCoordinator(
            new UiaTreeCaptureService(probe),
            new LocatorCandidateGenerator(),
            new NonInteractiveHighlightService(highlightSurface),
            new CandidateVerificationService(),
            new DebugArtifactWriter(),
            probe,
            vision,
            screenshots);
    }

    public bool IsActive
    {
        get { lock (_gate) return _active; }
    }

    public DebugSessionSnapshot Snapshot
    {
        get { lock (_gate) return BuildSnapshotUnlocked(); }
    }

    public DebugSessionSnapshot Start(string? outputDirectory = null)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_active)
                return BuildSnapshotUnlocked();

            _active = true;
            _sessionId = $"dbg-{DateTimeOffset.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}"[..28];
            _startedAtUtc = DateTimeOffset.UtcNow;
            _outputDirectory = string.IsNullOrWhiteSpace(outputDirectory)
                ? Path.Combine(Path.GetTempPath(), "desktop-automation-debug")
                : outputDirectory;
            _lastCapture = null;
            _lastCandidates = [];
            _lastVerification = null;
            _lastVisionCandidates = [];
            _lastVisionEligibility = null;
            _artifactPath = null;
            _logger.LogInformation("Debug Mode started ({SessionId})", _sessionId);
            return BuildSnapshotUnlocked("Debug Mode ON — Ctrl+D to stop; hover inspects without click");
        }
    }

    public DebugSessionSnapshot Stop()
    {
        lock (_gate)
        {
            if (!_active)
                return BuildSnapshotUnlocked("Debug Mode OFF");

            _highlight.Clear();
            _active = false;
            _logger.LogInformation("Debug Mode stopped ({SessionId})", _sessionId);
            return BuildSnapshotUnlocked("Debug Mode OFF");
        }
    }

    public DebugSessionSnapshot Toggle(string? outputDirectory = null) =>
        IsActive ? Stop() : Start(outputDirectory);

    public DebugSessionSnapshot InspectUnderCursor(Point screenPoint, bool writeArtifact = false)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_active)
                return BuildSnapshotUnlocked("Debug Mode is not active");

            var capture = _treeCapture.CaptureUnderCursor(screenPoint);
            if (capture is null)
            {
                _highlight.Clear();
                _lastCapture = null;
                _lastCandidates = [];
                _lastVerification = null;
                _lastVisionCandidates = [];
                _lastVisionEligibility = null;
                return BuildSnapshotUnlocked("No UIA element under cursor");
            }

            _lastCapture = capture;
            _highlight.Highlight(capture.Target);

            var siblings = capture.Children.Concat(capture.Ancestry).ToList();
            _lastCandidates = _candidateGenerator.Generate(capture.Target, siblings);
            _lastVerification = null;
            _lastVisionCandidates = [];
            _lastVisionEligibility = null;

            if (_lastCandidates.Count > 0)
            {
                // Live probe only — never convert probe failure / zero-match into unique via stale in-memory capture.
                var probeResult = _probe.CountLocatorMatches(
                    _lastCandidates[0].Locator,
                    capture.Window);
                _lastVerification = _verification.Verify(_lastCandidates[0], probeResult);
            }

            TryProduceVisionFallbackUnlocked(capture);

            if (writeArtifact && _lastCandidates.Count > 0)
            {
                _artifactPath = _artifactWriter.WriteCandidateArtifact(
                    _outputDirectory ?? Path.GetTempPath(),
                    _sessionId ?? "debug",
                    capture,
                    _lastCandidates,
                    _lastVerification);
            }

            var label = ElementLabel(capture.Target);
            var status = _lastVisionCandidates.Count > 0
                ? $"Inspect: {label} (vision candidates={_lastVisionCandidates.Count})"
                : $"Inspect: {label}";
            return BuildSnapshotUnlocked(status);
        }
    }

    public DebugSessionSnapshot VerifyLastCandidate()
    {
        lock (_gate)
        {
            if (!_active)
                return BuildSnapshotUnlocked("Debug Mode is not active");
            if (_lastCandidates.Count == 0 || _lastCapture is null)
                return BuildSnapshotUnlocked("No candidate to verify");

            // Live probe only — never fall back to stale in-memory capture for uniqueness.
            var probeResult = _probe.CountLocatorMatches(
                _lastCandidates[0].Locator,
                _lastCapture.Window);
            _lastVerification = _verification.Verify(_lastCandidates[0], probeResult);
            return BuildSnapshotUnlocked(_lastVerification.Message ?? "Verified");
        }
    }

    public HoverProbeResult ProbeHover(Point screenPoint)
    {
        lock (_gate)
        {
            return _probe.ProbeHover(screenPoint);
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
                return;
            _disposed = true;
            _active = false;
            _highlight.Dispose();
        }
    }

    private void TryProduceVisionFallbackUnlocked(DebugTreeCapture capture)
    {
        if (_vision is null)
            return;

        if (capture.Window is null || capture.Window.Bounds.IsEmpty)
        {
            _lastVisionCandidates = [];
            _lastVisionEligibility = null;
            return;
        }

        // Synthetic window-scoped PNG for the fake provider path (real OCR adapters replace this later).
        // When Vision:Enabled and UIA reports inaccessible, this is the production fallback path.
        var png = _screenshots is not null
            ? System.Text.Encoding.UTF8.GetBytes(
                $"phase8-vision-fallback:{capture.Window.WindowHandleHex}:{capture.Cursor.X},{capture.Cursor.Y}")
            : System.Text.Encoding.UTF8.GetBytes("phase8-vision-fallback");

        var query = !string.IsNullOrWhiteSpace(capture.Target.Name)
            ? capture.Target.Name
            : "Custom Drawn Target";

        var (eligibility, _, candidates) = _vision.TryProduceCandidates(
            capture.Target,
            capture.Window,
            png,
            query: query);
        _lastVisionEligibility = eligibility.ToString();
        _lastVisionCandidates = candidates;
        if (candidates.Count > 0)
        {
            _logger.LogInformation(
                "Vision fallback produced {Count} candidates (eligibility={Eligibility})",
                candidates.Count,
                eligibility);
        }
    }

    private DebugSessionSnapshot BuildSnapshotUnlocked(string? status = null) =>
        new()
        {
            IsActive = _active,
            SessionId = _sessionId,
            StartedAtUtc = _startedAtUtc,
            LastCapture = _lastCapture,
            LastCandidates = _lastCandidates,
            LastVerification = _lastVerification,
            ArtifactPath = _artifactPath,
            VisionCandidates = _lastVisionCandidates,
            VisionEligibility = _lastVisionEligibility,
            StatusText = status
                ?? (_active
                    ? "Debug Mode ON — Ctrl+D to stop"
                    : "Debug Mode OFF")
        };

    private static string ElementLabel(DebugUiaNode node) =>
        !string.IsNullOrWhiteSpace(node.Name) ? node.Name! :
        !string.IsNullOrWhiteSpace(node.AutomationId) ? node.AutomationId! :
        !string.IsNullOrWhiteSpace(node.ControlType) ? node.ControlType! :
        "(element)";
}
