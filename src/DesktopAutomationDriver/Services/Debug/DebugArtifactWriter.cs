using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Services.Assistive;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Writes Debug Mode candidate artifacts only. Never auto-promotes into approved pages.
/// </summary>
public sealed class DebugArtifactWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public const int DefaultMaxArtifactBytes = 262_144;

    public string WriteCandidateArtifact(
        string outputDirectory,
        string sessionId,
        DebugTreeCapture capture,
        IReadOnlyList<LocatorCandidate> candidates,
        CandidateVerificationResult? verification,
        int maxArtifactBytes = DefaultMaxArtifactBytes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(sessionId);
        ArgumentNullException.ThrowIfNull(capture);
        ArgumentNullException.ThrowIfNull(candidates);

        foreach (var candidate in candidates)
        {
            if (candidate.AutoPromoted)
                throw new InvalidOperationException("Debug Mode never auto-promotes candidates.");
        }

        var root = Path.GetFullPath(outputDirectory);
        Directory.CreateDirectory(root);
        AssistivePathSafety.EnsureWritablePathInside(root, root);

        var safeSession = Sanitize(sessionId);
        var folder = Path.Combine(root, "debug-candidates", safeSession);
        AssistivePathSafety.EnsureWritablePathInside(folder, root);
        Directory.CreateDirectory(folder);
        AssistivePathSafety.EnsureWritablePathInside(folder, root);

        var payload = new
        {
            kind = "debug-candidate",
            sessionId = safeSession,
            writtenAtUtc = DateTimeOffset.UtcNow,
            autoPromoted = false,
            requiresHumanApproval = false,
            requiresOperatorConfirmationUx = true,
            cursor = new { x = capture.Cursor.X, y = capture.Cursor.Y },
            window = capture.Window is null
                ? null
                : new
                {
                    capture.Window.ProcessName,
                    capture.Window.ProcessId,
                    capture.Window.WindowTitle,
                    bounds = Rect(capture.Window.Bounds),
                    capture.Window.DpiScaleX,
                    capture.Window.DpiScaleY
                },
            target = Node(capture.Target),
            ancestry = capture.Ancestry.Select(Node).ToList(),
            children = capture.Children.Select(Node).ToList(),
            candidates = candidates.Select(c => new
            {
                c.ElementId,
                locator = c.Locator,
                c.Grade,
                c.Strategy,
                c.Warnings,
                autoPromoted = false
            }).ToList(),
            verification
        };

        var json = JsonSerializer.Serialize(payload, JsonOptions);
        var bytes = Encoding.UTF8.GetBytes(json);
        if (bytes.Length > maxArtifactBytes)
        {
            throw new IOException(
                $"Debug candidate artifact size {bytes.Length} exceeds max {maxArtifactBytes} bytes.");
        }

        var target = Path.Combine(folder, "candidate.json");
        AssistivePathSafety.EnsureWritablePathInside(target, root);
        if (File.Exists(target))
            throw new IOException($"Debug candidate artifact already exists at '{target}'.");

        var temp = target + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllBytes(temp, bytes);
            AssistivePathSafety.EnsureWritablePathInside(temp, root);
            File.Move(temp, target);
            AssistivePathSafety.EnsureWritablePathInside(target, root);
        }
        catch
        {
            try
            {
                if (File.Exists(temp))
                    File.Delete(temp);
            }
            catch
            {
                // best-effort
            }

            throw;
        }

        return target;
    }

    private static object Node(DebugUiaNode node) => new
    {
        node.AutomationId,
        node.Name,
        node.ClassName,
        node.ControlType,
        node.ControlTypeId,
        node.Path,
        node.Depth,
        bounds = node.Bounds is null ? null : Rect(node.Bounds.Value),
        node.AvailablePatterns,
        node.IsOffscreen,
        node.IsEnabled,
        node.IsSemanticAccessible,
        node.InaccessibilityReason
    };

    private static object Rect(Rectangle r) => new { r.X, r.Y, r.Width, r.Height };

    private static string Sanitize(string value)
    {
        var trimmed = (value ?? "debug").Trim();
        if (trimmed.Length == 0)
            trimmed = "debug";
        foreach (var c in Path.GetInvalidFileNameChars())
            trimmed = trimmed.Replace(c, '_');
        return trimmed.Length > 64 ? trimmed[..64] : trimmed;
    }
}
