using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Debug Mode session coordinator. Keeps feature logic out of RecordingOverlayWindow / UiService.
/// </summary>
public sealed class DebugSessionCoordinator : IDebugSessionCoordinator, IDisposable
{
    private readonly UiaTreeCaptureService _treeCapture;
    private readonly LocatorCandidateGenerator _candidateGenerator;
    private readonly NonInteractiveHighlightService _highlight;
    private readonly CandidateVerificationService _verification;
    private readonly DebugArtifactWriter _artifactWriter;
    private readonly IUiaElementProbe _probe;
    private readonly ILogger<DebugSessionCoordinator> _logger;
    private readonly object _gate = new();

    private bool _active;
    private string? _sessionId;
    private DateTimeOffset? _startedAtUtc;
    private string? _outputDirectory;
    private DebugTreeCapture? _lastCapture;
    private IReadOnlyList<LocatorCandidate> _lastCandidates = [];
    private CandidateVerificationResult? _lastVerification;
    private string? _artifactPath;
    private bool _disposed;

    public DebugSessionCoordinator(
        UiaTreeCaptureService treeCapture,
        LocatorCandidateGenerator candidateGenerator,
        NonInteractiveHighlightService highlight,
        CandidateVerificationService verification,
        DebugArtifactWriter artifactWriter,
        IUiaElementProbe probe,
        ILogger<DebugSessionCoordinator>? logger = null)
    {
        _treeCapture = treeCapture ?? throw new ArgumentNullException(nameof(treeCapture));
        _candidateGenerator = candidateGenerator ?? throw new ArgumentNullException(nameof(candidateGenerator));
        _highlight = highlight ?? throw new ArgumentNullException(nameof(highlight));
        _verification = verification ?? throw new ArgumentNullException(nameof(verification));
        _artifactWriter = artifactWriter ?? throw new ArgumentNullException(nameof(artifactWriter));
        _probe = probe ?? throw new ArgumentNullException(nameof(probe));
        _logger = logger ?? NullLogger<DebugSessionCoordinator>.Instance;
    }

    public static DebugSessionCoordinator ForTests(
        IUiaElementProbe? probe = null,
        IHighlightSurface? highlightSurface = null)
    {
        probe ??= new NullUiaElementProbe();
        highlightSurface ??= new FakeHighlightSurface();
        return new DebugSessionCoordinator(
            new UiaTreeCaptureService(probe),
            new LocatorCandidateGenerator(),
            new NonInteractiveHighlightService(highlightSurface),
            new CandidateVerificationService(),
            new DebugArtifactWriter(),
            probe);
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
                return BuildSnapshotUnlocked("No UIA element under cursor");
            }

            _lastCapture = capture;
            _highlight.Highlight(capture.Target);

            var siblings = capture.Children.Concat(capture.Ancestry).ToList();
            _lastCandidates = _candidateGenerator.Generate(capture.Target, siblings);
            _lastVerification = null;

            if (_lastCandidates.Count > 0)
            {
                var matchCount = _probe.CountLocatorMatches(
                    _lastCandidates[0].Locator,
                    capture.Window);
                // Fall back to in-memory uniqueness when live probe has no scope.
                if (matchCount == 0)
                {
                    var pool = new List<DebugUiaNode> { capture.Target };
                    pool.AddRange(siblings);
                    matchCount = _verification.CountMatches(_lastCandidates[0].Locator, pool);
                }

                _lastVerification = _verification.Verify(_lastCandidates[0], matchCount);
            }

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
            return BuildSnapshotUnlocked($"Inspect: {label}");
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

            var matchCount = _probe.CountLocatorMatches(
                _lastCandidates[0].Locator,
                _lastCapture.Window);
            if (matchCount == 0)
            {
                var pool = new List<DebugUiaNode> { _lastCapture.Target };
                pool.AddRange(_lastCapture.Children);
                pool.AddRange(_lastCapture.Ancestry);
                matchCount = _verification.CountMatches(_lastCandidates[0].Locator, pool);
            }

            _lastVerification = _verification.Verify(_lastCandidates[0], matchCount);
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
