using System.Diagnostics;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Live FlaUI probe for Debug Mode. Kept outside RecordingOverlayWindow / UiService.
/// Requires an interactive desktop session.
/// </summary>
public sealed class FlaUIUiaElementProbe : IUiaElementProbe, IDisposable
{
    private readonly UIA3Automation _automation = new();
    private bool _disposed;

    public DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        try
        {
            var element = _automation.FromPoint(screenPoint);
            if (element is null)
                return null;

            var target = ToNode(element, depth: 0, path: "target");
            var ancestry = new List<DebugUiaNode>();
            var parent = SafeParent(element);
            var depth = 1;
            while (parent is not null && ancestry.Count < maxAncestry)
            {
                ancestry.Add(ToNode(parent, depth, path: $"ancestor/{depth}"));
                parent = SafeParent(parent);
                depth++;
            }

            var children = new List<DebugUiaNode>();
            try
            {
                var kids = element.FindAllChildren();
                for (var i = 0; i < kids.Length && children.Count < maxChildren; i++)
                    children.Add(ToNode(kids[i], 1, path: $"child/{i}"));
            }
            catch
            {
                // children optional
            }

            return new DebugTreeCapture
            {
                Target = target,
                Ancestry = ancestry,
                Children = children,
                Window = CaptureWindowIdentity(element),
                Cursor = screenPoint,
                CapturedAtUtc = DateTimeOffset.UtcNow
            };
        }
        catch
        {
            return null;
        }
    }

    public int CountLocatorMatches(IReadOnlyDictionary<string, object?> locator, DebugWindowIdentity? scope)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        try
        {
            AutomationElement root;
            if (scope is { WindowHandle: var hwnd } && hwnd != IntPtr.Zero)
                root = _automation.FromHandle(hwnd);
            else
                root = _automation.GetDesktop();

            var condition = BuildCondition(locator);
            if (condition is null)
                return 0;

            var matches = root.FindAllDescendants(condition);
            return matches.Length;
        }
        catch
        {
            return 0;
        }
    }

    public HoverProbeResult ProbeHover(Point screenPoint)
    {
        var beforeCapture = CaptureAtPoint(screenPoint, maxAncestry: 2, maxChildren: 5);
        var before = beforeCapture?.Target;

        // Mouse probe only — move cursor; do not click.
        try
        {
            Cursor.Position = screenPoint;
        }
        catch
        {
            // ignore in non-interactive hosts
        }

        Thread.Sleep(120);

        var afterCapture = CaptureAtPoint(screenPoint, maxAncestry: 2, maxChildren: 5);
        var after = afterCapture?.Target;
        string? tooltip = null;
        try
        {
            if (afterCapture is not null)
            {
                var tip = afterCapture.Children.FirstOrDefault(c =>
                    string.Equals(c.ControlType, "ToolTip", StringComparison.OrdinalIgnoreCase));
                tooltip = tip?.Name;
            }
        }
        catch
        {
            // ignore
        }

        var structureChanged = !string.Equals(before?.Path, after?.Path, StringComparison.Ordinal)
            || !string.Equals(before?.AutomationId, after?.AutomationId, StringComparison.Ordinal);
        var stateChanged = before?.IsEnabled != after?.IsEnabled
            || before?.IsOffscreen != after?.IsOffscreen
            || !string.Equals(before?.Name, after?.Name, StringComparison.Ordinal);

        return new HoverProbeResult
        {
            ScreenPoint = screenPoint,
            Before = before,
            After = after,
            TooltipText = tooltip,
            StructureChanged = structureChanged,
            StateChanged = stateChanged,
            NoOp = !structureChanged && !stateChanged && string.IsNullOrWhiteSpace(tooltip)
        };
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _automation.Dispose();
    }

    private ConditionBase? BuildCondition(IReadOnlyDictionary<string, object?> locator)
    {
        var cf = _automation.ConditionFactory;
        var parts = new List<ConditionBase>();

        if (locator.TryGetValue("automationId", out var a) && a is string automationId && !string.IsNullOrWhiteSpace(automationId))
            parts.Add(cf.ByAutomationId(automationId));
        if (locator.TryGetValue("name", out var n) && n is string name && !string.IsNullOrWhiteSpace(name))
            parts.Add(cf.ByName(name));
        if (locator.TryGetValue("className", out var c) && c is string className && !string.IsNullOrWhiteSpace(className))
            parts.Add(cf.ByClassName(className));
        if (locator.TryGetValue("controlType", out var t) && t is string controlType
            && Enum.TryParse<ControlType>(controlType, ignoreCase: true, out var parsed))
        {
            parts.Add(cf.ByControlType(parsed));
        }

        if (parts.Count == 0)
            return null;
        if (parts.Count == 1)
            return parts[0];
        return new AndCondition(parts.ToArray());
    }

    private static AutomationElement? SafeParent(AutomationElement element)
    {
        try
        {
            return element.Parent;
        }
        catch
        {
            return null;
        }
    }

    private static DebugUiaNode ToNode(AutomationElement element, int depth, string path)
    {
        Rectangle? bounds = null;
        try
        {
            var r = element.BoundingRectangle;
            if (!r.IsEmpty && r.Width > 0 && r.Height > 0)
            {
                bounds = new Rectangle(
                    (int)Math.Round(r.X),
                    (int)Math.Round(r.Y),
                    (int)Math.Round(r.Width),
                    (int)Math.Round(r.Height));
            }
        }
        catch
        {
            // ignore
        }

        var patterns = new List<string>();
        try
        {
            if (element.Patterns.Invoke.IsSupported) patterns.Add("Invoke");
            if (element.Patterns.Value.IsSupported) patterns.Add("Value");
            if (element.Patterns.Toggle.IsSupported) patterns.Add("Toggle");
            if (element.Patterns.SelectionItem.IsSupported) patterns.Add("SelectionItem");
            if (element.Patterns.ExpandCollapse.IsSupported) patterns.Add("ExpandCollapse");
            if (element.Patterns.Window.IsSupported) patterns.Add("Window");
            if (element.Patterns.Scroll.IsSupported) patterns.Add("Scroll");
            if (element.Patterns.Grid.IsSupported) patterns.Add("Grid");
        }
        catch
        {
            // ignore
        }

        string? automationId = null;
        string? name = null;
        string? className = null;
        string? controlType = null;
        var enabled = true;
        var offscreen = false;
        try
        {
            automationId = element.AutomationId;
            name = element.Name;
            className = element.ClassName;
            controlType = element.ControlType.ToString();
            enabled = element.IsEnabled;
            offscreen = element.IsOffscreen;
        }
        catch
        {
            // ignore
        }

        var inaccessible = string.Equals(controlType, "Custom", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(automationId)
            && string.IsNullOrWhiteSpace(name);

        return new DebugUiaNode
        {
            AutomationId = automationId,
            Name = name,
            ClassName = className,
            ControlType = controlType,
            Path = path,
            Depth = depth,
            Bounds = bounds,
            AvailablePatterns = patterns,
            IsEnabled = enabled,
            IsOffscreen = offscreen,
            IsSemanticAccessible = !inaccessible,
            InaccessibilityReason = inaccessible ? "custom-drawn-or-empty-identity" : null
        };
    }

    private static DebugWindowIdentity? CaptureWindowIdentity(AutomationElement element)
    {
        try
        {
            var current = element;
            while (current is not null && current.ControlType != ControlType.Window)
                current = SafeParent(current);

            current ??= element;
            var hwnd = new IntPtr(current.Properties.NativeWindowHandle.ValueOrDefault);
            var bounds = current.BoundingRectangle;
            int? pid = null;
            string? processName = null;
            try
            {
                var processId = current.Properties.ProcessId.ValueOrDefault;
                if (processId > 0)
                {
                    pid = processId;
                    processName = Process.GetProcessById(processId).ProcessName;
                }
            }
            catch
            {
                // ignore
            }

            return new DebugWindowIdentity
            {
                WindowHandle = hwnd,
                WindowTitle = current.Name,
                ProcessId = pid,
                ProcessName = processName,
                Bounds = new Rectangle(
                    (int)Math.Round(bounds.X),
                    (int)Math.Round(bounds.Y),
                    (int)Math.Round(bounds.Width),
                    (int)Math.Round(bounds.Height)),
                DpiScaleX = 1.0,
                DpiScaleY = 1.0
            };
        }
        catch
        {
            return null;
        }
    }
}
