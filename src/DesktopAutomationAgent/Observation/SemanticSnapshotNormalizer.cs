using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Observation;

public sealed class SemanticSnapshotNormalizer
{
    public const int MaxStringLength = 256;
    public const int DefaultMaxNodes = 500;

    private static readonly JsonSerializerOptions CanonicalJsonOptions = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly HashSet<string> StructuralControlTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "Pane",
        "Group",
        "Custom",
        "Thumb",
        "Separator",
        "TitleBar",
        "ScrollBar"
    };

    public SemanticSnapshot Normalize(
        JsonElement dumpuiaValue,
        ObservationOptions options,
        string snapshotId,
        DateTimeOffset capturedAtUtc,
        string driverVersion)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (string.IsNullOrWhiteSpace(snapshotId))
            throw new ArgumentException("snapshotId is required.", nameof(snapshotId));

        var warnings = new List<string>();
        var truncated = dumpuiaValue.TryGetProperty("truncated", out var truncatedElement)
                        && truncatedElement.ValueKind is JsonValueKind.True;

        if (dumpuiaValue.TryGetProperty("truncationWarnings", out var truncationWarnings)
            && truncationWarnings.ValueKind == JsonValueKind.Array)
        {
            foreach (var warning in truncationWarnings.EnumerateArray())
            {
                if (warning.ValueKind == JsonValueKind.String)
                {
                    var text = Bound(SecretRedactor.Redact(warning.GetString()));
                    if (!string.IsNullOrWhiteSpace(text))
                        warnings.Add(text);
                }
            }
        }

        var rawNodes = ReadRawNodes(dumpuiaValue);
        var maxNodes = Math.Clamp(options.MaxNodes ?? DefaultMaxNodes, 1, 2000);
        if (rawNodes.Count > maxNodes)
        {
            truncated = true;
            warnings.Add($"Agent truncated nodes from {rawNodes.Count} to maxNodes={maxNodes}.");
            rawNodes = rawNodes.Take(maxNodes).ToList();
        }

        var selected = options.ActionableOnly
            ? SelectActionable(rawNodes, warnings)
            : rawNodes.OrderBy(static n => n.SortKey, StringComparer.Ordinal).ToList();

        ValidateRawTree(selected);

        var nodes = AssignRefs(selected, options);
        var application = ReadApplication(dumpuiaValue);
        var session = ReadSession(dumpuiaValue);
        var window = ReadWindow(dumpuiaValue, nodes, options);

        // Editable values are never projected into the semantic snapshot.
        // Runtime IDs stay in-memory only when requested and are omitted from persisted hash inputs below.
        var snapshotForHash = new SemanticSnapshot
        {
            SnapshotSchemaVersion = SemanticSnapshot.SchemaVersion,
            SnapshotId = string.Empty,
            CapturedAtUtc = default,
            DriverVersion = driverVersion,
            Application = application,
            WindowsSession = session,
            Window = StripRuntimeSensitiveWindow(window),
            Nodes = nodes.Select(StripRuntimeId).ToList(),
            Truncated = truncated,
            Warnings = warnings.OrderBy(static w => w, StringComparer.Ordinal).ToList(),
            Ordering = SnapshotOrdering.Default
        };

        var hash = ComputeHash(snapshotForHash);

        return new SemanticSnapshot
        {
            SnapshotSchemaVersion = SemanticSnapshot.SchemaVersion,
            SnapshotId = snapshotId,
            CapturedAtUtc = capturedAtUtc,
            DriverVersion = driverVersion,
            SnapshotHash = hash,
            Application = application,
            WindowsSession = session,
            Window = window,
            Nodes = nodes,
            Truncated = truncated,
            Warnings = warnings,
            Ordering = SnapshotOrdering.Default
        };
    }

    public static string ComputeHash(SemanticSnapshot snapshot)
    {
        var payload = JsonSerializer.Serialize(new
        {
            snapshot.SnapshotSchemaVersion,
            snapshot.DriverVersion,
            snapshot.Application,
            snapshot.WindowsSession,
            window = snapshot.Window is null
                ? null
                : new
                {
                    snapshot.Window.Hwnd,
                    snapshot.Window.Title,
                    snapshot.Window.ClassName,
                    snapshot.Window.ProcessId,
                    snapshot.Window.IsModal,
                    snapshot.Window.Bounds
                },
            nodes = snapshot.Nodes.Select(n => new
            {
                n.Ref,
                n.ParentRef,
                n.Ordinal,
                n.Depth,
                n.ControlType,
                n.Role,
                n.Name,
                n.AutomationId,
                n.ClassName,
                n.HelpText,
                n.Enabled,
                n.Offscreen,
                n.Focusable,
                n.HasKeyboardFocus,
                n.Patterns,
                n.Capabilities,
                n.Bounds
            }),
            snapshot.Truncated,
            warnings = snapshot.Warnings,
            ordering = snapshot.Ordering.Rule
        }, CanonicalJsonOptions);

        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(payload));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static List<RawNode> ReadRawNodes(JsonElement dumpuiaValue)
    {
        if (!dumpuiaValue.TryGetProperty("nodes", out var nodesElement)
            || nodesElement.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var nodes = new List<RawNode>();
        var index = 0;
        foreach (var element in nodesElement.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object)
                continue;

            var patterns = ReadStringArray(element, "supportedPatterns");
            var capabilities = element.TryGetProperty("capabilities", out var capsElement)
                               && capsElement.ValueKind == JsonValueKind.Array
                ? capsElement.EnumerateArray()
                    .Where(static c => c.ValueKind == JsonValueKind.String)
                    .Select(static c => c.GetString()!)
                    .Where(static c => !string.IsNullOrWhiteSpace(c))
                    .ToList()
                : PatternCapabilityMapper.Map(patterns).ToList();

            var nodeIndex = GetNullableInt(element, "nodeIndex") ?? index;
            var parentNodeIndex = GetNullableInt(element, "parentNodeIndex");
            var ordinal = GetNullableInt(element, "ordinal") ?? index;
            var depth = GetInt(element, "depth");
            var controlType = Bound(GetString(element, "controlType")) ?? string.Empty;
            var name = Bound(GetString(element, "name"));
            var automationId = Bound(GetString(element, "automationId"));
            var className = Bound(GetString(element, "className"));

            // Never persist editable values even if a future dumpuia field exposes them.
            _ = GetString(element, "value");

            nodes.Add(new RawNode(
                nodeIndex,
                parentNodeIndex,
                ordinal,
                depth,
                controlType,
                name,
                automationId,
                className,
                Bound(GetString(element, "helpText")),
                GetBool(element, "isEnabled", defaultValue: true),
                GetBool(element, "isOffscreen", defaultValue: false),
                GetNullableBool(element, "isKeyboardFocusable"),
                GetNullableBool(element, "hasKeyboardFocus"),
                patterns,
                capabilities.OrderBy(static c => c, StringComparer.Ordinal).ToList(),
                ReadBounds(element),
                Bound(GetString(element, "runtimeId")),
                ReadIsModal(element),
                GetNullableLong(element, "nativeWindowHandle"),
                GetNullableInt(element, "processId"),
                BuildSortKey(depth, ordinal, automationId, name, controlType, nodeIndex)));

            index++;
        }

        return nodes
            .OrderBy(static n => n.SortKey, StringComparer.Ordinal)
            .ToList();
    }

    private static List<RawNode> SelectActionable(IReadOnlyList<RawNode> nodes, List<string> warnings)
    {
        var keep = new HashSet<int>();
        foreach (var node in nodes)
        {
            if (ShouldKeepActionable(node))
                keep.Add(node.NodeIndex);
        }

        // Preserve ancestors so parent refs remain coherent.
        var byIndex = nodes.ToDictionary(static n => n.NodeIndex);
        foreach (var index in keep.ToList())
        {
            var current = byIndex[index];
            while (current.ParentNodeIndex is int parent && byIndex.TryGetValue(parent, out var parentNode))
            {
                if (!keep.Add(parent))
                    break;
                current = parentNode;
            }
        }

        var selected = nodes.Where(n => keep.Contains(n.NodeIndex)).ToList();
        if (selected.Count < nodes.Count)
        {
            warnings.Add(
                $"Filtered to actionable/content-relevant nodes ({selected.Count} of {nodes.Count}).");
        }

        return selected.OrderBy(static n => n.SortKey, StringComparer.Ordinal).ToList();
    }

    private static bool ShouldKeepActionable(RawNode node)
    {
        if (node.Depth == 0)
            return true;

        if (!string.IsNullOrWhiteSpace(node.AutomationId) || !string.IsNullOrWhiteSpace(node.Name))
            return true;

        if (PatternCapabilityMapper.IsActionable(node.Capabilities))
            return true;

        if (node.ControlType is "Button" or "Edit" or "ComboBox" or "List" or "DataGrid" or "DataItem"
            or "Menu" or "MenuItem" or "MenuBar" or "CheckBox" or "RadioButton" or "TabItem"
            or "Hyperlink" or "TreeItem" or "Window" or "Dialog")
        {
            return true;
        }

        return !StructuralControlTypes.Contains(node.ControlType);
    }

    private static void ValidateRawTree(IReadOnlyList<RawNode> nodes)
    {
        var seenIndexes = new HashSet<int>();
        foreach (var node in nodes)
        {
            if (!seenIndexes.Add(node.NodeIndex))
                throw new InvalidOperationException($"Duplicate dumpuia nodeIndex '{node.NodeIndex}'.");
        }
    }

    private static List<SemanticNode> AssignRefs(IReadOnlyList<RawNode> nodes, ObservationOptions options)
    {
        var refByIndex = new Dictionary<int, string>();
        for (var i = 0; i < nodes.Count; i++)
            refByIndex[nodes[i].NodeIndex] = $"n{i}";

        var result = new List<SemanticNode>(nodes.Count);
        for (var i = 0; i < nodes.Count; i++)
        {
            var node = nodes[i];
            string? parentRef = null;
            if (node.ParentNodeIndex is int parentIndex)
            {
                if (!refByIndex.TryGetValue(parentIndex, out parentRef))
                {
                    // Parent was filtered out; leave null rather than inventing a cross-tree link.
                    parentRef = null;
                }
            }

            result.Add(new SemanticNode
            {
                Ref = refByIndex[node.NodeIndex],
                ParentRef = parentRef,
                Ordinal = node.Ordinal,
                Depth = node.Depth,
                ControlType = node.ControlType,
                Role = node.ControlType,
                Name = NullIfEmpty(node.Name),
                AutomationId = NullIfEmpty(node.AutomationId),
                ClassName = NullIfEmpty(node.ClassName),
                HelpText = options.IncludeHelpText ? NullIfEmpty(node.HelpText) : null,
                Enabled = node.Enabled,
                Offscreen = node.Offscreen,
                Focusable = options.IncludeFocusState ? node.Focusable : null,
                HasKeyboardFocus = options.IncludeFocusState ? node.HasKeyboardFocus : null,
                Patterns = node.Patterns.OrderBy(static p => p, StringComparer.Ordinal).ToList(),
                Capabilities = node.Capabilities,
                Bounds = options.IncludeBounds ? node.Bounds : null,
                RuntimeId = options.IncludeRuntimeId ? NullIfEmpty(node.RuntimeId) : null
            });
        }

        // Reject invalid refs before returning.
        _ = new EphemeralRefIndex(new SemanticSnapshot
        {
            SnapshotId = "validate",
            Nodes = result
        });

        return result;
    }

    private static ApplicationIdentity ReadApplication(JsonElement dumpuiaValue)
    {
        if (!dumpuiaValue.TryGetProperty("rootDiagnostics", out var root)
            || root.ValueKind != JsonValueKind.Object)
        {
            return new ApplicationIdentity();
        }

        return new ApplicationIdentity
        {
            ProcessId = GetNullableInt(root, "processId"),
            ProcessName = NullIfEmpty(Bound(GetString(root, "processName")))
        };
    }

    private static WindowsSessionIdentity ReadSession(JsonElement dumpuiaValue)
    {
        if (!dumpuiaValue.TryGetProperty("rootDiagnostics", out var root)
            || root.ValueKind != JsonValueKind.Object)
        {
            return new WindowsSessionIdentity();
        }

        return new WindowsSessionIdentity
        {
            SessionId = GetNullableInt(root, "windowsSessionId")
        };
    }

    private static WindowIdentity? ReadWindow(
        JsonElement dumpuiaValue,
        IReadOnlyList<SemanticNode> nodes,
        ObservationOptions options)
    {
        if (!dumpuiaValue.TryGetProperty("rootDiagnostics", out var root)
            || root.ValueKind != JsonValueKind.Object)
        {
            return nodes.Count == 0
                ? null
                : new WindowIdentity { Ref = nodes[0].Ref };
        }

        long? hwnd = GetNullableLong(root, "activeWindowHwnd");
        string? title = null;
        string? className = null;
        int? processId = GetNullableInt(root, "processId");
        bool? isModal = null;
        BoundingRectangleDto? bounds = null;

        if (root.TryGetProperty("activeWindow", out var active) && active.ValueKind == JsonValueKind.Object)
        {
            title = NullIfEmpty(Bound(GetString(active, "name")));
            className = NullIfEmpty(Bound(GetString(active, "className")));
            processId ??= GetNullableInt(active, "processId");
            hwnd ??= GetNullableLong(active, "nativeWindowHandle");
            isModal = ReadIsModal(active);
            if (options.IncludeBounds)
                bounds = ReadBounds(active);
        }

        return new WindowIdentity
        {
            Ref = nodes.Count > 0 ? nodes[0].Ref : null,
            Hwnd = hwnd is > 0 ? hwnd : null,
            Title = title,
            ClassName = className,
            ProcessId = processId,
            IsModal = isModal,
            Bounds = bounds
        };
    }

    private static WindowIdentity? StripRuntimeSensitiveWindow(WindowIdentity? window) =>
        window is null
            ? null
            : new WindowIdentity
            {
                Hwnd = window.Hwnd,
                Title = window.Title,
                ClassName = window.ClassName,
                ProcessId = window.ProcessId,
                IsModal = window.IsModal,
                Bounds = window.Bounds
            };

    private static SemanticNode StripRuntimeId(SemanticNode node) =>
        new()
        {
            Ref = node.Ref,
            ParentRef = node.ParentRef,
            Ordinal = node.Ordinal,
            Depth = node.Depth,
            ControlType = node.ControlType,
            Role = node.Role,
            Name = node.Name,
            AutomationId = node.AutomationId,
            ClassName = node.ClassName,
            HelpText = node.HelpText,
            Enabled = node.Enabled,
            Offscreen = node.Offscreen,
            Focusable = node.Focusable,
            HasKeyboardFocus = node.HasKeyboardFocus,
            Patterns = node.Patterns,
            Capabilities = node.Capabilities,
            Bounds = node.Bounds
        };

    private static BoundingRectangleDto? ReadBounds(JsonElement element)
    {
        if (!element.TryGetProperty("boundingRectangle", out var bounds)
            || bounds.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        return new BoundingRectangleDto
        {
            X = GetNullableDouble(bounds, "x") ?? GetNullableDouble(bounds, "left"),
            Y = GetNullableDouble(bounds, "y") ?? GetNullableDouble(bounds, "top"),
            Width = GetNullableDouble(bounds, "width"),
            Height = GetNullableDouble(bounds, "height"),
            Left = GetNullableDouble(bounds, "left"),
            Top = GetNullableDouble(bounds, "top"),
            Right = GetNullableDouble(bounds, "right"),
            Bottom = GetNullableDouble(bounds, "bottom")
        };
    }

    private static bool? ReadIsModal(JsonElement element)
    {
        if (element.TryGetProperty("windowState", out var windowState)
            && windowState.ValueKind == JsonValueKind.Object
            && windowState.TryGetProperty("isModal", out var isModal)
            && isModal.ValueKind is JsonValueKind.True or JsonValueKind.False)
        {
            return isModal.GetBoolean();
        }

        return null;
    }

    private static List<string> ReadStringArray(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value) || value.ValueKind != JsonValueKind.Array)
            return [];

        return value.EnumerateArray()
            .Where(static item => item.ValueKind == JsonValueKind.String)
            .Select(static item => Bound(item.GetString())!)
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .OrderBy(static item => item, StringComparer.Ordinal)
            .ToList()!;
    }

    private static string BuildSortKey(
        int depth,
        int ordinal,
        string? automationId,
        string? name,
        string? controlType,
        int nodeIndex) =>
        $"{depth:D4}|{ordinal:D6}|{automationId}|{name}|{controlType}|{nodeIndex:D6}";

    private static string? GetString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static int GetInt(JsonElement element, string propertyName) =>
        GetNullableInt(element, propertyName) ?? 0;

    private static int? GetNullableInt(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            return number;
        return null;
    }

    private static long? GetNullableLong(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out var number))
            return number;
        return null;
    }

    private static double? GetNullableDouble(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        return null;
    }

    private static bool GetBool(JsonElement element, string propertyName, bool defaultValue)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return defaultValue;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => defaultValue
        };
    }

    private static bool? GetNullableBool(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => null
        };
    }

    private static string? Bound(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        var redacted = SecretRedactor.Redact(value);
        return redacted.Length <= MaxStringLength
            ? redacted
            : redacted[..MaxStringLength];
    }

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private sealed record RawNode(
        int NodeIndex,
        int? ParentNodeIndex,
        int Ordinal,
        int Depth,
        string ControlType,
        string? Name,
        string? AutomationId,
        string? ClassName,
        string? HelpText,
        bool Enabled,
        bool Offscreen,
        bool? Focusable,
        bool? HasKeyboardFocus,
        List<string> Patterns,
        List<string> Capabilities,
        BoundingRectangleDto? Bounds,
        string? RuntimeId,
        bool? IsModal,
        long? NativeWindowHandle,
        int? ProcessId,
        string SortKey);
}
