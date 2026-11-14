using System.Diagnostics;
using System.Runtime.InteropServices;
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

    public LocatorMatchProbeResult CountLocatorMatches(
        IReadOnlyDictionary<string, object?> locator,
        DebugWindowIdentity? scope)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        try
        {
            // Fail closed: missing HWND means the live probe cannot safely scope — report unavailable,
            // never invent a zero-match that Debug Mode could confuse with uniqueness from stale captures.
            if (scope is null || scope.WindowHandle == IntPtr.Zero)
                return LocatorMatchProbeResult.Unavailable("Window HWND required for live locator match probe.");

            var root = _automation.FromHandle(scope.WindowHandle);
            if (root is null)
                return LocatorMatchProbeResult.Unavailable("Unable to resolve window HWND for locator match probe.");

            var condition = BuildCondition(locator);
            if (condition is null)
                return LocatorMatchProbeResult.Unavailable("Locator condition could not be built.");

            var matches = root.FindAllDescendants(condition);
            return LocatorMatchProbeResult.Ok(matches.Length);
        }
        catch (Exception ex)
        {
            return LocatorMatchProbeResult.Unavailable(ex.Message);
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
                    (int)Math.Round((double)r.X),
                    (int)Math.Round((double)r.Y),
                    (int)Math.Round((double)r.Width),
                    (int)Math.Round((double)r.Height));
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

        var inaccessible = IsEmptyIdentityInaccessible(controlType, automationId, name);

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

    /// <summary>
    /// Empty AutomationId+Name on Custom/Pane surfaces (WinForms UserPaint often reports Pane).
    /// </summary>
    internal static bool IsEmptyIdentityInaccessible(string? controlType, string? automationId, string? name)
    {
        if (!string.IsNullOrWhiteSpace(automationId) || !string.IsNullOrWhiteSpace(name))
            return false;
        return string.Equals(controlType, "Custom", StringComparison.OrdinalIgnoreCase)
               || string.Equals(controlType, "Pane", StringComparison.OrdinalIgnoreCase);
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

            var (dpiX, dpiY) = ResolveDpiScale(hwnd);
            return new DebugWindowIdentity
            {
                WindowHandle = hwnd,
                WindowTitle = current.Name,
                ProcessId = pid,
                ProcessName = processName,
                Bounds = new Rectangle(
                    (int)Math.Round((double)bounds.X),
                    (int)Math.Round((double)bounds.Y),
                    (int)Math.Round((double)bounds.Width),
                    (int)Math.Round((double)bounds.Height)),
                DpiScaleX = dpiX,
                DpiScaleY = dpiY
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Convert Win32 window DPI to scale factors (96 DPI = 1.0).</summary>
    internal static (double DpiScaleX, double DpiScaleY) ScaleFromWindowDpi(uint dpi)
    {
        if (dpi == 0)
            return (1.0, 1.0);
        var scale = dpi / 96.0;
        return (scale, scale);
    }

    private static (double DpiScaleX, double DpiScaleY) ResolveDpiScale(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return (1.0, 1.0);
        try
        {
            var dpi = GetDpiForWindow(hwnd);
            return ScaleFromWindowDpi(dpi);
        }
        catch
        {
            return (1.0, 1.0);
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);
}
