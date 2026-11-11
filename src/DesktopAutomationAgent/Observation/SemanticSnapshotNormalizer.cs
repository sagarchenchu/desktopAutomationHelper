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
    public const int RequiredSemanticDiagnosticVersion = 1;

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

        RequireSemanticDiagnostics(dumpuiaValue);

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
        if (rawNodes.Count == 0)
            throw new InvalidOperationException("dumpuia returned zero nodes with semantic tree links.");

        ValidateRawTree(rawNodes);

        var maxNodes = Math.Clamp(options.MaxNodes ?? DefaultMaxNodes, 1, 2000);
        if (rawNodes.Count > maxNodes)
        {
            truncated = true;
            warnings.Add($"Agent truncated nodes from {rawNodes.Count} to maxNodes={maxNodes}.");
            rawNodes = TruncatePreservingAncestors(rawNodes, maxNodes);
        }

        var selected = options.ActionableOnly
            ? SelectActionable(rawNodes, warnings)
            : rawNodes;

        var ordered = OrderDepthFirst(selected);
        var nodes = AssignRefs(ordered, options);
        var application = ReadApplication(dumpuiaValue);
        var session = ReadSession(dumpuiaValue);
        var window = ReadWindow(dumpuiaValue, ordered, nodes, options);

        var hash = ComputeHash(new SemanticSnapshot
        {
            SnapshotSchemaVersion = SemanticSnapshot.SchemaVersion,
            DriverVersion = driverVersion,
            Application = application,
            WindowsSession = session,
            Window = window,
            Nodes = nodes,
            Truncated = truncated,
            Warnings = warnings.OrderBy(static w => w, StringComparer.Ordinal).ToList(),
            Ordering = SnapshotOrdering.Default
        });

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
        // Exclude volatile identity: PID, session ID, HWND, and optional bounds.
        var payload = JsonSerializer.Serialize(new
        {
            snapshot.SnapshotSchemaVersion,
            snapshot.DriverVersion,
            application = new
            {
                processName = snapshot.Application.ProcessName
            },
            windowsSession = new { },
            window = snapshot.Window is null
                ? null
                : new
                {
                    snapshot.Window.Ref,
                    snapshot.Window.Title,
                    snapshot.Window.ClassName,
                    snapshot.Window.IsModal
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
                n.Capabilities
            }),
            snapshot.Truncated,
            warnings = snapshot.Warnings,
            ordering = snapshot.Ordering.Rule
        }, CanonicalJsonOptions);

        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(payload));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    private static void RequireSemanticDiagnostics(JsonElement dumpuiaValue)
    {
        var version = GetNullableInt(dumpuiaValue, "semanticDiagnosticVersion");
        if (version is null || version < RequiredSemanticDiagnosticVersion)
        {
            throw new InvalidOperationException(
                $"Driver dumpuia lacks semanticDiagnosticVersion>={RequiredSemanticDiagnosticVersion}. Upgrade the driver; flat/degraded snapshots are not accepted.");
        }

        if (!dumpuiaValue.TryGetProperty("nodes", out var nodes)
            || nodes.ValueKind != JsonValueKind.Array
            || nodes.GetArrayLength() == 0)
        {
            throw new InvalidOperationException("dumpuia response did not include nodes.");
        }

        foreach (var node in nodes.EnumerateArray())
        {
            if (node.ValueKind != JsonValueKind.Object)
                continue;

            if (!node.TryGetProperty("nodeIndex", out _))
            {
                throw new InvalidOperationException(
                    "Driver dumpuia omitted required tree-link field 'nodeIndex'. Upgrade the driver; flat/degraded snapshots are not accepted.");
            }
        }
    }

    private static List<RawNode> ReadRawNodes(JsonElement dumpuiaValue)
    {
        if (!dumpuiaValue.TryGetProperty("nodes", out var nodesElement)
            || nodesElement.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var nodes = new List<RawNode>();
        foreach (var element in nodesElement.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object)
                continue;

            if (!element.TryGetProperty("nodeIndex", out var nodeIndexElement)
                || !nodeIndexElement.TryGetInt32(out var nodeIndex))
            {
                throw new InvalidOperationException("dumpuia node is missing required nodeIndex tree link.");
            }

            var patterns = ReadStringArray(element, "supportedPatterns");
            var valueIsReadOnly = GetNullableBool(element, "valueIsReadOnly");
            var capabilities = element.TryGetProperty("capabilities", out var capsElement)
                               && capsElement.ValueKind == JsonValueKind.Array
                ? capsElement.EnumerateArray()
                    .Where(static c => c.ValueKind == JsonValueKind.String)
                    .Select(static c => c.GetString()!)
                    .Where(static c => !string.IsNullOrWhiteSpace(c))
                    .ToList()
                : PatternCapabilityMapper.Map(patterns, valueIsReadOnly).ToList();

            // Fail closed for Phase 4B: never invent setValue when writable state is unknown/false.
            if (valueIsReadOnly != false)
                capabilities.RemoveAll(static c => c == "setValue");
            else if (patterns.Any(static p => string.Equals(p, "Value", StringComparison.OrdinalIgnoreCase))
                     && !capabilities.Contains("setValue", StringComparer.Ordinal))
            {
                capabilities.Add("setValue");
            }

            if (patterns.Any(static p => string.Equals(p, "Value", StringComparison.OrdinalIgnoreCase))
                && !capabilities.Contains("readValue", StringComparer.Ordinal))
            {
                capabilities.Add("readValue");
            }

            var parentNodeIndex = GetNullableInt(element, "parentNodeIndex");
            var ordinal = GetNullableInt(element, "ordinal") ?? 0;
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
                valueIsReadOnly,
                ReadInteraction(element)));
        }

        // Preserve driver dump order; do not globally re-sort by depth.
        return nodes;
    }

    private static List<RawNode> TruncatePreservingAncestors(IReadOnlyList<RawNode> nodes, int maxNodes)
    {
        if (nodes.Count <= maxNodes)
            return nodes.ToList();

        var byIndex = nodes.ToDictionary(static n => n.NodeIndex);
        var keep = new HashSet<int>();
        foreach (var node in nodes)
        {
            if (keep.Count >= maxNodes)
                break;

            keep.Add(node.NodeIndex);
            var current = node;
            while (current.ParentNodeIndex is int parent && byIndex.TryGetValue(parent, out var parentNode))
            {
                if (!keep.Add(parent) || keep.Count >= maxNodes)
                    break;
                current = parentNode;
            }
        }

        return OrderDepthFirst(nodes.Where(n => keep.Contains(n.NodeIndex)).ToList());
    }

    private static List<RawNode> SelectActionable(IReadOnlyList<RawNode> nodes, List<string> warnings)
    {
        var keep = new HashSet<int>();
        foreach (var node in nodes)
        {
            if (ShouldKeepActionable(node))
                keep.Add(node.NodeIndex);
        }

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

        return selected;
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

    /// <summary>
    /// Depth-first preorder via parent links. Sibling order only is sorted
    /// (ordinal, then automationId/name/controlType/nodeIndex).
    /// </summary>
    internal static List<RawNode> OrderDepthFirst(IReadOnlyList<RawNode> nodes)
    {
        var byIndex = nodes.ToDictionary(static n => n.NodeIndex);
        var children = new Dictionary<int, List<RawNode>>();
        var roots = new List<RawNode>();

        foreach (var node in nodes)
        {
            if (node.ParentNodeIndex is int parent && byIndex.ContainsKey(parent))
            {
                if (!children.TryGetValue(parent, out var list))
                {
                    list = [];
                    children[parent] = list;
                }

                list.Add(node);
            }
            else
            {
                roots.Add(node);
            }
        }

        static int Compare(RawNode left, RawNode right) =>
            string.CompareOrdinal(
                BuildSiblingSortKey(left),
                BuildSiblingSortKey(right));

        foreach (var pair in children)
            pair.Value.Sort(Compare);

        roots.Sort(Compare);

        var ordered = new List<RawNode>(nodes.Count);
        var visiting = new HashSet<int>();

        void Visit(RawNode node)
        {
            if (!visiting.Add(node.NodeIndex))
                return;

            ordered.Add(node);
            if (children.TryGetValue(node.NodeIndex, out var kids))
            {
                foreach (var child in kids)
                    Visit(child);
            }
        }

        foreach (var root in roots)
            Visit(root);

        // Orphans with broken parent links already treated as roots.
        foreach (var node in nodes)
        {
            if (!visiting.Contains(node.NodeIndex))
                Visit(node);
        }

        return ordered;
    }

    private static string BuildSiblingSortKey(RawNode node) =>
        $"{node.Ordinal:D6}|{node.AutomationId}|{node.Name}|{node.ControlType}|{node.NodeIndex:D6}";

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
                refByIndex.TryGetValue(parentIndex, out parentRef);

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
                Interaction = node.Interaction,
                Bounds = options.IncludeBounds ? node.Bounds : null,
                RuntimeId = options.IncludeRuntimeId ? NullIfEmpty(node.RuntimeId) : null
            });
        }

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

        int? processId = GetNullableInt(root, "processId");
        if (root.TryGetProperty("activeWindow", out var active)
            && active.ValueKind == JsonValueKind.Object)
        {
            processId = GetNullableInt(active, "processId") ?? processId;
        }

        return new ApplicationIdentity
        {
            ProcessId = processId,
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
        IReadOnlyList<RawNode> orderedRaw,
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
        // Prefer the observed window's processId over a possibly stale session PID.
        int? processId = null;
        bool? isModal = null;
        BoundingRectangleDto? bounds = null;

        if (root.TryGetProperty("activeWindow", out var active) && active.ValueKind == JsonValueKind.Object)
        {
            title = NullIfEmpty(Bound(GetString(active, "name")));
            className = NullIfEmpty(Bound(GetString(active, "className")));
            processId = GetNullableInt(active, "processId") ?? GetNullableInt(root, "processId");
            hwnd ??= GetNullableLong(active, "nativeWindowHandle");
            isModal = ReadIsModal(active);
            if (options.IncludeBounds)
                bounds = ReadBounds(active);
        }

        string? windowRef = null;
        if (hwnd is > 0)
        {
            for (var i = 0; i < orderedRaw.Count; i++)
            {
                if (orderedRaw[i].NativeWindowHandle == hwnd)
                {
                    windowRef = nodes[i].Ref;
                    break;
                }
            }
        }

        windowRef ??= nodes.Count > 0 ? nodes[0].Ref : null;

        return new WindowIdentity
        {
            Ref = windowRef,
            Hwnd = hwnd is > 0 ? hwnd : null,
            Title = title,
            ClassName = className,
            ProcessId = processId,
            IsModal = isModal,
            Bounds = bounds
        };
    }

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

    private static NodeInteractionState? ReadInteraction(JsonElement element)
    {
        if (!element.TryGetProperty("interaction", out var interaction)
            || interaction.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        string? expandCollapse = null;
        bool? selected = null;
        string? toggle = null;

        if (interaction.TryGetProperty("expandCollapse", out var expandElement)
            && expandElement.ValueKind == JsonValueKind.String)
        {
            expandCollapse = Bound(expandElement.GetString());
        }

        if (interaction.TryGetProperty("selected", out var selectedElement))
        {
            if (selectedElement.ValueKind is JsonValueKind.True or JsonValueKind.False)
                selected = selectedElement.GetBoolean();
        }

        if (interaction.TryGetProperty("toggle", out var toggleElement)
            && toggleElement.ValueKind == JsonValueKind.String)
        {
            toggle = Bound(toggleElement.GetString());
        }

        if (expandCollapse is null && selected is null && toggle is null)
            return null;

        return new NodeInteractionState
        {
            ExpandCollapse = expandCollapse,
            Selected = selected,
            Toggle = toggle
        };
    }

    internal sealed record RawNode(
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
        bool? ValueIsReadOnly,
        NodeInteractionState? Interaction);
}
