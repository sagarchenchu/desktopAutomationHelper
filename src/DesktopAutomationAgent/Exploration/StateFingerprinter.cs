using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Stable state identity for exploration. Ignores volatile timestamps, runtime refs,
/// handles, coordinates, PID, and session id. Collision risk: expand/selected-only
/// mutations that do not change the visible node set or names may fingerprint as equal
/// until pattern state is projected into semantic snapshots.
/// </summary>
public sealed class StateFingerprinter
{
    private static readonly JsonSerializerOptions CanonicalJsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    public string Fingerprint(SemanticSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);

        var payload = JsonSerializer.Serialize(new
        {
            schema = SemanticSnapshot.SchemaVersion,
            application = new
            {
                processName = snapshot.Application.ProcessName
            },
            window = snapshot.Window is null
                ? null
                : new
                {
                    title = snapshot.Window.Title,
                    className = snapshot.Window.ClassName,
                    isModal = snapshot.Window.IsModal
                },
            nodes = snapshot.Nodes.Select(n => new
            {
                ordinal = n.Ordinal,
                depth = n.Depth,
                controlType = n.ControlType,
                role = n.Role,
                name = n.Name,
                automationId = n.AutomationId,
                className = n.ClassName,
                enabled = n.Enabled,
                offscreen = n.Offscreen,
                patterns = n.Patterns,
                capabilities = n.Capabilities
            }),
            truncated = snapshot.Truncated
        }, CanonicalJsonOptions);

        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(payload));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }
}
