using System.Drawing;
using System.Drawing.Imaging;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// FlaUI-backed implementation of <see cref="IUiService"/>.
/// All element-locating operations include a built-in retry of up to 5 seconds
/// (500 ms polling interval) so that callers do not need to add their own waits.
/// </summary>
public class UiService : IUiService
{
    private readonly IUiSessionContext _ctx;
    private readonly ILogger<UiService> _logger;

    private static readonly TimeSpan DefaultRetry = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan RetryInterval = TimeSpan.FromMilliseconds(500);

    // Named keys for the "sendkeys" operation (AutoIt / keyboard-shorthand format).
    private static readonly Dictionary<string, VirtualKeyShort> NamedKeys =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["ENTER"] = VirtualKeyShort.RETURN,
            ["RETURN"] = VirtualKeyShort.RETURN,
            ["TAB"] = VirtualKeyShort.TAB,
            ["ESC"] = VirtualKeyShort.ESCAPE,
            ["ESCAPE"] = VirtualKeyShort.ESCAPE,
            ["BACKSPACE"] = VirtualKeyShort.BACK,
            ["BS"] = VirtualKeyShort.BACK,
            ["DELETE"] = VirtualKeyShort.DELETE,
            ["DEL"] = VirtualKeyShort.DELETE,
            ["INSERT"] = VirtualKeyShort.INSERT,
            ["INS"] = VirtualKeyShort.INSERT,
            ["HOME"] = VirtualKeyShort.HOME,
            ["END"] = VirtualKeyShort.END,
            ["UP"] = VirtualKeyShort.UP,
            ["DOWN"] = VirtualKeyShort.DOWN,
            ["LEFT"] = VirtualKeyShort.LEFT,
            ["RIGHT"] = VirtualKeyShort.RIGHT,
            ["PGUP"] = VirtualKeyShort.PRIOR,
            ["PAGEUP"] = VirtualKeyShort.PRIOR,
            ["PGDN"] = VirtualKeyShort.NEXT,
            ["PAGEDOWN"] = VirtualKeyShort.NEXT,
            ["SPACE"] = VirtualKeyShort.SPACE,
            ["F1"] = VirtualKeyShort.F1,
            ["F2"] = VirtualKeyShort.F2,
            ["F3"] = VirtualKeyShort.F3,
            ["F4"] = VirtualKeyShort.F4,
            ["F5"] = VirtualKeyShort.F5,
            ["F6"] = VirtualKeyShort.F6,
            ["F7"] = VirtualKeyShort.F7,
            ["F8"] = VirtualKeyShort.F8,
            ["F9"] = VirtualKeyShort.F9,
            ["F10"] = VirtualKeyShort.F10,
            ["F11"] = VirtualKeyShort.F11,
            ["F12"] = VirtualKeyShort.F12,
        };

    public UiService(IUiSessionContext ctx, ILogger<UiService> logger)
    {
        _ctx = ctx;
        _logger = logger;
    }

    // =========================================================================
    // Public dispatch entry point
    // =========================================================================

    /// <inheritdoc/>
    public object? Execute(UiRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Operation))
            throw new ArgumentException("'operation' is required.");

        _logger.LogDebug("UI operation: {Operation}", SanitizeValue(request.Operation));

        return request.Operation.ToLowerInvariant() switch
        {
            // ----- Session & Window Management -----
            "launch"       => Launch(request),
            "close"        => Close(),
            "maximize"     => Maximize(),
            "minimize"     => Minimize(),
            "switchwindow" => SwitchWindow(request),
            "refresh"      => Refresh(),
            "screenshot"   => Screenshot(request),
            "listelements" => ListElements(request),
            "listwindows"  => ListWindows(request),

            // ----- Element Query -----
            "exists"         => Exists(request),
            "waitfor"        => WaitFor(request),
            "isenabled"      => IsEnabled(request),
            "isvisible"      => IsVisible(request),
            "isclickable"    => IsClickable(request),
            "ischecked"      => IsChecked(request),
            "getvalue"       => GetValue(request),
            "gettext"        => GetText(request),
            "getname"        => GetName(request),
            "getcontroltype" => GetControlType(request),
            "getselected"    => GetSelected(request),
            "gettable"       => GetTable(request),
            "gettableheaders"=> GetTableHeaders(request),

            // ----- Position Comparison -----
            "isrightof"   => IsRightOf(request),
            "isleftof"    => IsLeftOf(request),
            "isabove"     => IsAbove(request),
            "isbelow"     => IsBelow(request),
            "getposition" => GetPosition(request),

            // ----- Element Actions -----
            "click"       => Click(request),
            "doubleclick" => DoubleClick(request),
            "rightclick"  => RightClick(request),
            "hover"       => Hover(request),
            "focus"       => Focus(request),
            "type"        => TypeText(request),
            "clear"       => Clear(request),
            "sendkeys"    => SendKeys(request),
            "scroll"      => Scroll(request),
            "check"       => Check(request),
            "uncheck"     => Uncheck(request),
            "select"      => Select(request),
            "selectaid"   => SelectByAid(request),

            _ => throw new ArgumentException(
                $"Unknown operation '{request.Operation}'. " +
                "See GET /ui/operations for the full list.")
        };
    }

    // =========================================================================
    // Session & Window Management
    // =========================================================================

    private object? Launch(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' must be the path to the executable for 'launch'.");

        var session = _ctx.Launch(req.Value);
        return new { sessionId = session.SessionId, app = req.Value };
    }

    private object? Close()
    {
        RequireSession();
        _ctx.Close();
        return null;
    }

    private object? Maximize()
    {
        var session = RequireSession();
        var window = GetWindowRoot(session).AsWindow();
        window.Patterns.Window.Pattern.SetWindowVisualState(WindowVisualState.Maximized);
        return null;
    }

    private object? Minimize()
    {
        var session = RequireSession();
        var window = GetWindowRoot(session).AsWindow();
        window.Patterns.Window.Pattern.SetWindowVisualState(WindowVisualState.Minimized);
        return null;
    }

    private object? SwitchWindow(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' must be a partial window title for 'switchwindow'.");

        var session = RequireSession();
        var allWindows = session.Application.GetAllTopLevelWindows(session.Automation);
        var match = allWindows.FirstOrDefault(w =>
            w.Name.Contains(req.Value, StringComparison.OrdinalIgnoreCase));

        if (match == null)
            throw new InvalidOperationException(
                $"No window with title containing '{req.Value}' was found.");

        session.ActiveWindow = match;
        match.SetForeground();
        return new { title = match.Name };
    }

    private object? Refresh()
    {
        var session = RequireSession();
        // Reset the tracked window so the next call re-queries GetMainWindow().
        session.ActiveWindow = null;
        return null;
    }

    private object? Screenshot(UiRequest req)
    {
        var session = RequireSession();
        var root = GetWindowRoot(session);
        var rect = root.BoundingRectangle;

        using var bitmap = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(rect.Width, rect.Height));

        var filePath = !string.IsNullOrWhiteSpace(req.Value)
            ? req.Value
            : Path.Combine(Path.GetTempPath(),
                $"screenshot_{DateTime.UtcNow:yyyyMMddHHmmss}.png");

        bitmap.Save(filePath, ImageFormat.Png);
        return new { path = filePath };
    }

    private object? ListElements(UiRequest req)
    {
        var session = RequireSession();
        var root = GetWindowRoot(session);
        var cf = session.Automation.ConditionFactory;

        AutomationElement[] elements;
        if (!string.IsNullOrWhiteSpace(req.Value))
        {
            var ct = ParseControlType(req.Value);
            elements = root.FindAllDescendants(cf.ByControlType(ct));
        }
        else
        {
            elements = root.FindAllDescendants();
        }

        return elements.Select(e => new
        {
            name = e.Name,
            automationId = e.AutomationId,
            className = e.ClassName,
            controlType = e.ControlType.ToString(),
            enabled = e.IsEnabled,
            visible = !e.IsOffscreen
        }).ToList();
    }

    private object? ListWindows(UiRequest req)
    {
        var session = RequireSession();
        var allWindows = session.Application.GetAllTopLevelWindows(session.Automation);

        IEnumerable<AutomationElement> filtered = allWindows;
        if (!string.IsNullOrWhiteSpace(req.Value))
            filtered = allWindows.Where(w =>
                w.Name.Contains(req.Value, StringComparison.OrdinalIgnoreCase));

        return filtered.Select(w => new
        {
            title = w.Name,
            automationId = w.AutomationId,
            className = w.ClassName
        }).ToList();
    }

    // =========================================================================
    // Element Query
    // =========================================================================

    private object? Exists(UiRequest req)
    {
        var locator = RequireLocator(req);
        var session = RequireSession();
        var root = GetWindowRoot(session);

        // Exists checks immediately — no retry — and never throws.
        var element = TryFindElement(root, session, locator);
        return new { exists = element != null };
    }

    private object? WaitFor(UiRequest req)
    {
        var locator = RequireLocator(req);
        var session = RequireSession();
        var root = GetWindowRoot(session);

        double timeoutSeconds = 10;
        if (!string.IsNullOrWhiteSpace(req.Value) &&
            double.TryParse(req.Value, out var parsed))
            timeoutSeconds = parsed;

        var timeout = TimeSpan.FromSeconds(timeoutSeconds);
        var deadline = DateTime.UtcNow + timeout;

        while (true)
        {
            var element = TryFindElement(root, session, locator);
            if (element != null && element.IsEnabled && !element.IsOffscreen)
                return new { found = true };

            if (DateTime.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"Element not found or not ready within {timeoutSeconds}s: " +
                    DescribeLocator(locator));

            Thread.Sleep(RetryInterval);
        }
    }

    private object? IsEnabled(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { enabled = element.IsEnabled };
    }

    private object? IsVisible(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { visible = !element.IsOffscreen };
    }

    private object? IsClickable(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { clickable = element.IsEnabled && !element.IsOffscreen };
    }

    private object? IsChecked(UiRequest req)
    {
        var element = FindWithRetry(req);

        if (element.Patterns.Toggle.IsSupported)
            return new { @checked = element.Patterns.Toggle.Pattern.ToggleState == ToggleState.On };

        if (element.Patterns.SelectionItem.IsSupported)
            return new { @checked = element.Patterns.SelectionItem.Pattern.IsSelected };

        return new { @checked = false };
    }

    private object? GetValue(UiRequest req)
    {
        var element = FindWithRetry(req);

        if (element.Patterns.Value.IsSupported)
            return new { value = element.Patterns.Value.Pattern.Value ?? string.Empty };

        if (element.Patterns.RangeValue.IsSupported)
            return new { value = element.Patterns.RangeValue.Pattern.Value.ToString() };

        return new { value = element.Name ?? string.Empty };
    }

    private object? GetText(UiRequest req)
    {
        var element = FindWithRetry(req);
        const int MaxText = 1_048_576;

        try
        {
            if (element.Patterns.Text.IsSupported)
                return new { text = element.Patterns.Text.Pattern.DocumentRange.GetText(MaxText) };
        }
        catch { /* fall through */ }

        try
        {
            if (element.Patterns.Value.IsSupported)
                return new { text = element.Patterns.Value.Pattern.Value ?? string.Empty };
        }
        catch { /* fall through */ }

        return new { text = element.Name ?? string.Empty };
    }

    private object? GetName(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { name = element.Name ?? string.Empty };
    }

    private object? GetControlType(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { controlType = element.ControlType.ToString() };
    }

    private object? GetSelected(UiRequest req)
    {
        var element = FindWithRetry(req);

        if (element.Patterns.Selection.IsSupported)
        {
            var selected = element.Patterns.Selection.Pattern.Selection.Value;
            var first = selected.FirstOrDefault();
            return new { selected = first?.Name ?? string.Empty };
        }

        if (element.Patterns.SelectionItem.IsSupported)
            return new { selected = element.Name ?? string.Empty };

        throw new InvalidOperationException(
            "Element does not support the Selection or SelectionItem pattern.");
    }

    private object? GetTable(UiRequest req)
    {
        var element = FindWithRetry(req);
        return ReadTableData(element, headersOnly: false);
    }

    private object? GetTableHeaders(UiRequest req)
    {
        var element = FindWithRetry(req);
        return ReadTableData(element, headersOnly: true);
    }

    // =========================================================================
    // Position Comparison
    // =========================================================================

    private object? IsRightOf(UiRequest req)
    {
        var (r1, r2) = GetTwoRects(req);
        return new { isRightOf = r1.Left >= r2.Right };
    }

    private object? IsLeftOf(UiRequest req)
    {
        var (r1, r2) = GetTwoRects(req);
        return new { isLeftOf = r1.Right <= r2.Left };
    }

    private object? IsAbove(UiRequest req)
    {
        var (r1, r2) = GetTwoRects(req);
        return new { isAbove = r1.Bottom <= r2.Top };
    }

    private object? IsBelow(UiRequest req)
    {
        var (r1, r2) = GetTwoRects(req);
        return new { isBelow = r1.Top >= r2.Bottom };
    }

    private object? GetPosition(UiRequest req)
    {
        var (r1, r2) = GetTwoRects(req);
        return new
        {
            element1  = RectObj(r1),
            element2  = RectObj(r2),
            isRightOf = r1.Left  >= r2.Right,
            isLeftOf  = r1.Right <= r2.Left,
            isAbove   = r1.Bottom <= r2.Top,
            isBelow   = r1.Top   >= r2.Bottom
        };
    }

    // =========================================================================
    // Element Actions
    // =========================================================================

    private object? Click(UiRequest req)
    {
        FindWithRetry(req).Click();
        return null;
    }

    private object? DoubleClick(UiRequest req)
    {
        FindWithRetry(req).DoubleClick();
        return null;
    }

    private object? RightClick(UiRequest req)
    {
        FindWithRetry(req).RightClick();
        return null;
    }

    private object? Hover(UiRequest req)
    {
        var element = FindWithRetry(req);
        var pt = element.GetClickablePoint();
        Mouse.MoveTo(pt);
        return null;
    }

    private object? Focus(UiRequest req)
    {
        FindWithRetry(req).Focus();
        return null;
    }

    private object? TypeText(UiRequest req)
    {
        if (req.Value == null)
            throw new ArgumentException("'value' is required for 'type'.");

        var element = FindWithRetry(req);
        element.Focus();
        Keyboard.Type(req.Value);
        return null;
    }

    private object? Clear(UiRequest req)
    {
        var element = FindWithRetry(req);
        element.AsTextBox().Text = string.Empty;
        return null;
    }

    private object? SendKeys(UiRequest req)
    {
        if (req.Value == null)
            throw new ArgumentException("'value' is required for 'sendkeys'.");

        var element = FindWithRetry(req);
        element.Focus();
        SendKeysString(req.Value);
        return null;
    }

    private object? Scroll(UiRequest req)
    {
        var element = FindWithRetry(req);
        element.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();
        return null;
    }

    private object? Check(UiRequest req)
    {
        SetToggle(req, wantChecked: true);
        return null;
    }

    private object? Uncheck(UiRequest req)
    {
        SetToggle(req, wantChecked: false);
        return null;
    }

    private object? Select(UiRequest req)
    {
        if (req.Value == null && req.Index == null)
            throw new ArgumentException("Either 'value' (item name) or 'index' is required for 'select'.");

        var element = FindWithRetry(req);
        var cf = RequireSession().Automation.ConditionFactory;

        // Expand the combo/list so that items become available.
        element.Patterns.ExpandCollapse.PatternOrDefault?.Expand();
        Thread.Sleep(100);

        var items = element.FindAllDescendants(cf.ByControlType(ControlType.ListItem));

        if (req.Value != null)
        {
            var match = items.FirstOrDefault(i =>
                i.Name.Equals(req.Value, StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException(
                    $"ComboBox item '{req.Value}' not found.");
            match.Patterns.SelectionItem.Pattern.Select();
        }
        else
        {
            var idx = req.Index!.Value;
            if (idx < 0 || idx >= items.Length)
                throw new ArgumentException(
                    $"Index {idx} is out of range. ComboBox has {items.Length} item(s).");
            items[idx].Patterns.SelectionItem.Pattern.Select();
        }

        return null;
    }

    private object? SelectByAid(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' (item AutomationId) is required for 'selectaid'.");

        var element = FindWithRetry(req);
        var cf = RequireSession().Automation.ConditionFactory;

        // Expand the combo/list so that items become available.
        element.Patterns.ExpandCollapse.PatternOrDefault?.Expand();
        Thread.Sleep(100);

        var items = element.FindAllDescendants(cf.ByControlType(ControlType.ListItem));
        var match = items.FirstOrDefault(i =>
            i.AutomationId.Equals(req.Value, StringComparison.OrdinalIgnoreCase));

        if (match == null)
            throw new InvalidOperationException(
                $"ComboBox item with AutomationId '{req.Value}' not found.");

        match.Patterns.SelectionItem.Pattern.Select();
        return null;
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    /// <summary>
    /// Returns the active session or throws if none exists.
    /// </summary>
    private AutomationSession RequireSession()
    {
        var session = _ctx.ActiveSession;
        if (session == null)
            throw new InvalidOperationException(
                "No active session. Use 'launch' first.");
        return session;
    }

    /// <summary>
    /// Returns the locator from the request or throws if missing.
    /// </summary>
    private static UiLocator RequireLocator(UiRequest req)
    {
        if (req.Locator == null)
            throw new ArgumentException("'locator' is required for this operation.");
        return req.Locator;
    }

    /// <summary>
    /// Returns the current window root for element searches.
    /// Falls back to the application's main window when no window switch has occurred.
    /// </summary>
    private AutomationElement GetWindowRoot(AutomationSession session) =>
        session.ActiveWindow
            ?? session.Application.GetMainWindow(session.Automation)
            ?? throw new InvalidOperationException("Could not find the main window of the application.");

    /// <summary>
    /// Finds an element using a locator with up to 5 s retry (500 ms interval).
    /// </summary>
    private AutomationElement FindWithRetry(UiRequest req)
    {
        var locator = RequireLocator(req);
        var session = RequireSession();
        var root = GetWindowRoot(session);
        return FindLocatorWithRetry(session, root, locator);
    }

    /// <summary>
    /// Finds an element by <paramref name="locator"/> with up to 5 s retry.
    /// Used directly when a locator is already in hand (e.g. locator2 lookups).
    /// </summary>
    private AutomationElement FindLocatorWithRetry(
        AutomationSession session, AutomationElement root, UiLocator locator)
    {
        var deadline = DateTime.UtcNow + DefaultRetry;
        while (true)
        {
            var element = TryFindElement(root, session, locator);
            if (element != null)
                return element;

            if (DateTime.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"Element not found within {DefaultRetry.TotalSeconds}s: " +
                    DescribeLocator(locator));

            Thread.Sleep(RetryInterval);
        }
    }

    /// <summary>
    /// Attempts a single element search. Returns null when not found.
    /// </summary>
    private static AutomationElement? TryFindElement(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        try
        {
            var condition = BuildCondition(session, locator);
            return root.FindFirstDescendant(condition);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Builds a FlaUI condition from a <see cref="UiLocator"/>.
    /// All supplied properties are combined with AND logic.
    /// </summary>
    private static ConditionBase BuildCondition(AutomationSession session, UiLocator locator)
    {
        var cf = session.Automation.ConditionFactory;
        var conditions = new List<ConditionBase>();

        if (!string.IsNullOrWhiteSpace(locator.Name))
            conditions.Add(cf.ByName(locator.Name));
        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            conditions.Add(cf.ByAutomationId(locator.AutomationId));
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
            conditions.Add(cf.ByClassName(locator.ClassName));
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
            conditions.Add(cf.ByControlType(ParseControlType(locator.ControlType)));

        if (conditions.Count == 0)
            throw new ArgumentException(
                "Locator must specify at least one property " +
                "(name, automationId, className, or controlType).");

        return conditions.Count == 1
            ? conditions[0]
            : new AndCondition(conditions.ToArray());
    }

    private static ControlType ParseControlType(string value)
    {
        if (Enum.TryParse<ControlType>(value, ignoreCase: true, out var ct))
            return ct;
        throw new ArgumentException(
            $"Unknown controlType '{value}'. " +
            "Valid values: Button, CheckBox, ComboBox, DataGrid, DataItem, Edit, " +
            "Group, List, ListItem, Menu, MenuItem, Pane, RadioButton, Slider, " +
            "Spinner, Tab, TabItem, Table, Text, ToolBar, Tree, TreeItem, Window.");
    }

    private static string DescribeLocator(UiLocator l)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(l.Name))           parts.Add($"name='{l.Name}'");
        if (!string.IsNullOrWhiteSpace(l.AutomationId))   parts.Add($"automationId='{l.AutomationId}'");
        if (!string.IsNullOrWhiteSpace(l.ClassName))      parts.Add($"className='{l.ClassName}'");
        if (!string.IsNullOrWhiteSpace(l.ControlType))    parts.Add($"controlType='{l.ControlType}'");
        return string.Join(", ", parts);
    }

    private (Rectangle r1, Rectangle r2) GetTwoRects(UiRequest req)
    {
        if (req.Locator == null)
            throw new ArgumentException("'locator' is required for position operations.");
        if (req.Locator2 == null)
            throw new ArgumentException("'locator2' is required for position operations.");

        var session = RequireSession();
        var root = GetWindowRoot(session);

        var e1 = FindWithRetry(req);
        // Find the second element directly by its locator (no UiRequest wrapper needed).
        var e2 = FindLocatorWithRetry(session, root, req.Locator2);

        return (e1.BoundingRectangle, e2.BoundingRectangle);
    }

    private static object RectObj(Rectangle r) => new
    {
        left   = r.Left,
        top    = r.Top,
        right  = r.Right,
        bottom = r.Bottom,
        width  = r.Width,
        height = r.Height
    };

    private void SetToggle(UiRequest req, bool wantChecked)
    {
        var element = FindWithRetry(req);

        if (element.Patterns.Toggle.IsSupported)
        {
            var toggle = element.Patterns.Toggle.Pattern;
            var current = toggle.ToggleState;

            // ToggleState cycles: Off → On → Indeterminate (for tri-state).
            // Toggle until we reach the desired state (max 2 presses).
            for (int i = 0; i < 2; i++)
            {
                if (wantChecked && current == ToggleState.On)   return;
                if (!wantChecked && current == ToggleState.Off) return;
                toggle.Toggle();
                current = toggle.ToggleState;
            }
            return;
        }

        throw new InvalidOperationException(
            "Element does not support the Toggle pattern (check/uncheck).");
    }

    /// <summary>
    /// Reads table/grid data from an element via the Grid and Table patterns.
    /// </summary>
    private static object ReadTableData(AutomationElement element, bool headersOnly)
    {
        var headers = new List<string>();
        var rows    = new List<List<string>>();

        if (element.Patterns.Table.IsSupported)
        {
            var table = element.Patterns.Table.Pattern;
            foreach (var h in table.ColumnHeaders.Value)
                headers.Add(h.Name);
        }

        if (!headersOnly && element.Patterns.Grid.IsSupported)
        {
            var grid = element.Patterns.Grid.Pattern;
            int rowCount = grid.RowCount;
            int colCount = grid.ColumnCount;

            for (int r = 0; r < rowCount; r++)
            {
                var row = new List<string>();
                for (int c = 0; c < colCount; c++)
                {
                    var cell = grid.GetItem(r, c);
                    row.Add(cell?.Name ?? string.Empty);
                }
                rows.Add(row);
            }
        }

        if (headersOnly)
            return new { headers };

        return new { headers, rows };
    }

    /// <summary>
    /// Sends a key sequence string using AutoIt/keyboard-shorthand notation:
    /// {ENTER}, {TAB}, {F5}, etc. for named keys;
    /// ^x for Ctrl+X, +x for Shift+X, %x for Alt+X;
    /// any other character is typed literally.
    /// </summary>
    private static void SendKeysString(string keys)
    {
        int i = 0;
        while (i < keys.Length)
        {
            char c = keys[i];

            // {KEYNAME} — named key
            if (c == '{')
            {
                int end = keys.IndexOf('}', i + 1);
                if (end > i)
                {
                    var keyName = keys[(i + 1)..end];
                    if (NamedKeys.TryGetValue(keyName, out var vk))
                    {
                        Keyboard.Press(vk);
                        Keyboard.Release(vk);
                    }
                    i = end + 1;
                    continue;
                }
            }

            // ^x — Ctrl+char
            if (c == '^' && i + 1 < keys.Length)
            {
                var vk = CharToVirtualKey(keys[i + 1]);
                if (vk.HasValue)
                {
                    Keyboard.Press(VirtualKeyShort.LCONTROL);
                    Keyboard.Press(vk.Value);
                    Keyboard.Release(vk.Value);
                    Keyboard.Release(VirtualKeyShort.LCONTROL);
                    i += 2;
                    continue;
                }
            }

            // +x — Shift+char
            if (c == '+' && i + 1 < keys.Length)
            {
                var vk = CharToVirtualKey(keys[i + 1]);
                if (vk.HasValue)
                {
                    Keyboard.Press(VirtualKeyShort.LSHIFT);
                    Keyboard.Press(vk.Value);
                    Keyboard.Release(vk.Value);
                    Keyboard.Release(VirtualKeyShort.LSHIFT);
                    i += 2;
                    continue;
                }
            }

            // %x — Alt+char
            if (c == '%' && i + 1 < keys.Length)
            {
                var vk = CharToVirtualKey(keys[i + 1]);
                if (vk.HasValue)
                {
                    Keyboard.Press(VirtualKeyShort.ALT);
                    Keyboard.Press(vk.Value);
                    Keyboard.Release(vk.Value);
                    Keyboard.Release(VirtualKeyShort.ALT);
                    i += 2;
                    continue;
                }
            }

            // Literal character
            Keyboard.Type(c.ToString());
            i++;
        }
    }

    private static VirtualKeyShort? CharToVirtualKey(char c)
    {
        c = char.ToUpperInvariant(c);
        if (c >= 'A' && c <= 'Z')
            return (VirtualKeyShort)('A' + (c - 'A'));
        if (c >= '0' && c <= '9')
            return (VirtualKeyShort)('0' + (c - '0'));
        return null;
    }

    /// <summary>
    /// Strips control characters from a user-supplied string before it is
    /// written to a log message to prevent log-injection attacks.
    /// </summary>
    private static string SanitizeValue(string? value) =>
        System.Text.RegularExpressions.Regex.Replace(
            value ?? string.Empty, @"[\r\n\t]", "_");
}
