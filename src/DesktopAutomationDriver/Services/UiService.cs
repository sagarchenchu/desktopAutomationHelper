using System.Drawing;
using System.Drawing.Imaging;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

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
            "quit"         => Close(),
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
            "gettabledata"   => GetTable(request),
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
            "clickgridcell" => ClickGridCell(request),

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

        var session = _ctx.ActiveSession;
        var deadline = DateTime.UtcNow + DefaultRetry;

        if (session != null)
        {
            // Retry loop: the target window may not yet be visible immediately after
            // an action that opens it (e.g. a button click), so we poll until it appears.
            AutomationElement? match;
            while (true)
            {
                match = FindWindowByTitle(session, req.Value);
                if (match != null) break;
                if (DateTime.UtcNow >= deadline)
                    throw new InvalidOperationException(
                        $"No window with title containing '{req.Value}' was found within {DefaultRetry.TotalSeconds}s.");
                Thread.Sleep(RetryInterval);
            }

            session.ActiveWindow = match;
            // Seed the handle so the auto-follow logic in GetWindowRoot does not
            // immediately override this explicit switch on the very next operation.
            var switchedHandle = SafeWindowHandle(match);
            if (switchedHandle != IntPtr.Zero)
                session.SeedWindowHandles([switchedHandle]);
            match.SetForeground();
            return new { title = match.Name };
        }
        else
        {
            // No active session: search all desktop windows using a temporary automation,
            // then attach to the found window's process to establish a session.
            int processId = FindWindowProcessId(req.Value, deadline);
            session = _ctx.Attach(processId);

            // Re-find the window using the freshly created session's automation.
            var match = FindWindowByTitle(session, req.Value);
            if (match == null)
                throw new InvalidOperationException(
                    $"No window with title containing '{req.Value}' was found.");

            session.ActiveWindow = match;
            // Seed the handle so the auto-follow logic in GetWindowRoot does not
            // immediately override this explicit switch on the very next operation.
            var switchedHandle = SafeWindowHandle(match);
            if (switchedHandle != IntPtr.Zero)
                session.SeedWindowHandles([switchedHandle]);
            match.SetForeground();
            return new { title = match.Name };
        }
    }

    /// <summary>
    /// Searches the session's application windows and then all desktop windows for one
    /// whose title contains <paramref name="titleFragment"/>.
    /// Returns null when no match is found.
    /// </summary>
    private static AutomationElement? FindWindowByTitle(AutomationSession session, string titleFragment)
    {
        // Try the application's own top-level windows first (fast path).
        var match = session.Application.GetAllTopLevelWindows(session.Automation)
            .FirstOrDefault(w => w.Name.Contains(titleFragment, StringComparison.OrdinalIgnoreCase));

        if (match != null) return match;

        // Fall back to ALL desktop windows so callers can switch to any open window,
        // including windows from other processes.
        var cf = session.Automation.ConditionFactory;
        return session.Automation.GetDesktop()
            .FindAllChildren(cf.ByControlType(ControlType.Window))
            .FirstOrDefault(w => w.Name.Contains(titleFragment, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Searches all top-level desktop windows for one whose title contains
    /// <paramref name="titleFragment"/> and returns its process ID.
    /// Retries until <paramref name="deadline"/> to handle windows that open with a delay.
    /// Throws <see cref="InvalidOperationException"/> if no match is found before the deadline.
    /// </summary>
    private static int FindWindowProcessId(string titleFragment, DateTime deadline)
    {
        using var tempAutomation = new UIA3Automation();
        var cf = tempAutomation.ConditionFactory;

        while (true)
        {
            var match = tempAutomation.GetDesktop()
                .FindAllChildren(cf.ByControlType(ControlType.Window))
                .FirstOrDefault(w =>
                    w.Name.Contains(titleFragment, StringComparison.OrdinalIgnoreCase));

            if (match != null)
                return match.Properties.ProcessId.Value;

            if (DateTime.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"No window with title containing '{titleFragment}' was found within {DefaultRetry.TotalSeconds}s.");

            Thread.Sleep(RetryInterval);
        }
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
        var element = FindWithRetry(req);
        // Prefer the UIA Invoke pattern: it triggers the element's primary action
        // synchronously on the application's UI thread with no mouse-movement
        // overhead, which is significantly faster than a simulated mouse click.
        if (element.Patterns.Invoke.IsSupported)
            element.Patterns.Invoke.Pattern.Invoke();
        else
            element.Click();
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

        // Strategy 1: For editable combo boxes the Value pattern lets us set the text
        // directly without opening the dropdown, which is the most reliable approach.
        if (req.Value != null
            && element.Patterns.Value.IsSupported
            && element.Patterns.Value.PatternOrDefault?.IsReadOnly == false)
        {
            try
            {
                element.Patterns.Value.Pattern.SetValue(req.Value);
                return null;
            }
            catch
            {
                // The Value pattern sometimes reports IsReadOnly=false but still throws
                // (e.g. a non-editable combo with a text renderer). Fall through to the
                // expand-and-select strategy so the caller's request is still fulfilled.
            }
        }

        // Strategy 2: Expand the combo/list so that items become available.
        // 300 ms allows the dropdown animation to complete and UIA child elements to
        // materialise in the accessibility tree on typical hardware; faster machines
        // will still pass since FindAllDescendants is called only after the sleep.
        element.Patterns.ExpandCollapse.PatternOrDefault?.Expand();
        Thread.Sleep(300);

        // Collect list items – some combo boxes nest items inside a List child rather
        // than exposing them as direct descendants of the ComboBox element.
        var items = element.FindAllDescendants(cf.ByControlType(ControlType.ListItem));
        if (items.Length == 0)
        {
            var listChild = element.FindFirstDescendant(cf.ByControlType(ControlType.List));
            if (listChild != null)
                items = listChild.FindAllChildren();
        }

        AutomationElement? target;
        if (req.Value != null)
        {
            target = items.FirstOrDefault(i =>
                i.Name.Equals(req.Value, StringComparison.OrdinalIgnoreCase));
            if (target == null)
                throw new InvalidOperationException(
                    $"ComboBox item '{req.Value}' not found.");
        }
        else
        {
            var idx = req.Index!.Value;
            if (idx < 0 || idx >= items.Length)
                throw new ArgumentException(
                    $"Index {idx} is out of range. ComboBox has {items.Length} item(s).");
            target = items[idx];
        }

        // Scroll the item into view so it is reachable, then select it.
        target.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();

        var selectionItemPattern = target.Patterns.SelectionItem.PatternOrDefault;
        if (selectionItemPattern != null)
        {
            selectionItemPattern.Select();
        }
        else
        {
            // Fallback: physically click the item when SelectionItem is not supported.
            target.Click();
        }

        // Brief pause to allow the application to commit the selection before we
        // collapse the dropdown; some implementations only persist the selected value
        // once the list item's handler has run.
        Thread.Sleep(100);

        // Collapse to commit the selection (some native ComboBox implementations only
        // persist the selected value once the dropdown is dismissed).
        element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();

        return null;
    }

    private object? SelectByAid(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' (item AutomationId) is required for 'selectaid'.");

        var element = FindWithRetry(req);
        var cf = RequireSession().Automation.ConditionFactory;

        // Expand the combo/list so that items become available.
        // 300 ms allows the dropdown animation to complete and UIA child elements to
        // materialise in the accessibility tree on typical hardware.
        element.Patterns.ExpandCollapse.PatternOrDefault?.Expand();
        Thread.Sleep(300);

        var items = element.FindAllDescendants(cf.ByControlType(ControlType.ListItem));
        if (items.Length == 0)
        {
            var listChild = element.FindFirstDescendant(cf.ByControlType(ControlType.List));
            if (listChild != null)
                items = listChild.FindAllChildren();
        }

        var match = items.FirstOrDefault(i =>
            i.AutomationId.Equals(req.Value, StringComparison.OrdinalIgnoreCase));

        if (match == null)
            throw new InvalidOperationException(
                $"ComboBox item with AutomationId '{req.Value}' not found.");

        match.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();

        var selectionItemPattern = match.Patterns.SelectionItem.PatternOrDefault;
        if (selectionItemPattern != null)
            selectionItemPattern.Select();
        else
            match.Click();

        Thread.Sleep(100);

        // Collapse to commit the selection.
        element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();

        return null;
    }

    private object? ClickGridCell(UiRequest req)
    {
        if (req.Index == null)
            throw new ArgumentException("'index' (row index) is required for 'clickgridcell'.");
        if (req.ColumnIndex == null)
            throw new ArgumentException("'columnIndex' is required for 'clickgridcell'.");

        var element = FindWithRetry(req);

        // "Grid pattern not supported" is a bad-request condition (wrong element type),
        // not a "not found", so ArgumentException gives the correct 400 status code.
        if (!element.Patterns.Grid.IsSupported)
            throw new ArgumentException(
                "The target element does not support the Grid pattern. " +
                "Verify the locator targets the grid/table control itself, not a child row or cell.");

        var grid = element.Patterns.Grid.Pattern;
        int row = req.Index.Value;
        int col = req.ColumnIndex.Value;

        // The explicit < 0 guards are intentionally kept separate from the upper-bound
        // checks below: when RowCount/ColumnCount == 0 (virtualised grid) the upper-bound
        // check is skipped, so negative indices would otherwise reach grid.GetItem unchecked.
        if (row < 0)
            throw new ArgumentException($"Row index {row} must be >= 0.");
        if (col < 0)
            throw new ArgumentException($"Column index {col} must be >= 0.");

        // RowCount/ColumnCount report 0 for virtualised grids even when data is present;
        // in that case skip the upper-bound check and let GetItem surface the error if
        // the coordinates are truly out of range.
        int rowCount = grid.RowCount;
        int colCount = grid.ColumnCount;

        if (rowCount > 0 && row >= rowCount)
            throw new ArgumentException(
                $"Row index {row} is out of range. Grid has {rowCount} row(s).");
        if (colCount > 0 && col >= colCount)
            throw new ArgumentException(
                $"Column index {col} is out of range. Grid has {colCount} column(s).");

        var cell = grid.GetItem(row, col);
        if (cell == null)
            throw new InvalidOperationException(
                $"Grid cell at row {row}, column {col} could not be retrieved.");

        if (cell.Patterns.Invoke.IsSupported)
            cell.Patterns.Invoke.Pattern.Invoke();
        else
            cell.Click();

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
    /// When <see cref="AutomationSession.AutoFollowNewWindows"/> is true (the default),
    /// automatically switches to any top-level window that has opened in the application
    /// since the session started (or since the last explicit window switch), and clears
    /// a stale <see cref="AutomationSession.ActiveWindow"/> if the window has closed.
    /// Falls back to the application's main window when no active window is set.
    /// </summary>
    private AutomationElement GetWindowRoot(AutomationSession session)
    {
        if (session.AutoFollowNewWindows)
        {
            try
            {
                var allWindows = session.Application.GetAllTopLevelWindows(session.Automation);

                // Auto-follow: switch to the first window that has opened since last check.
                var newWindow = session.ClaimFirstNewWindow(allWindows);
                if (newWindow != null)
                {
                    session.ActiveWindow = newWindow;
                    _logger.LogInformation(
                        "Auto-followed new window: '{Title}'", SanitizeValue(newWindow.Name));
                }

                // Validate that the current ActiveWindow is still open; clear it if closed.
                if (session.ActiveWindow != null)
                {
                    var activeHandle = SafeWindowHandle(session.ActiveWindow);
                    if (activeHandle == IntPtr.Zero)
                    {
                        // Element handle is zero — element is gone.
                        _logger.LogInformation(
                            "Active window closed; reverting to main application window.");
                        session.ActiveWindow = null;
                    }
                    else if (!allWindows.Any(w => SafeWindowHandle(w) == activeHandle))
                    {
                        // Handle not in this session's process — could be a cross-process window
                        // set by SwitchWindow. Only clear if the element is truly inaccessible.
                        if (!IsElementAlive(session.ActiveWindow))
                        {
                            _logger.LogInformation(
                                "Active window closed; reverting to main application window.");
                            session.ActiveWindow = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex,
                    "Auto-follow window check failed; continuing with current active window.");
            }
        }

        return session.ActiveWindow
            ?? session.Application.GetMainWindow(session.Automation)
            ?? throw new InvalidOperationException("Could not find the main window of the application.");
    }

    /// <summary>
    /// Returns true when the element is still accessible via UI Automation.
    /// Accessing a stale or destroyed element raises a COMException; this helper
    /// treats any such exception as "element is gone".
    /// NativeWindowHandle is tried first; cross-process window elements may not
    /// expose that property, so ProcessId is used as a reliable fallback.
    /// </summary>
    private static bool IsElementAlive(AutomationElement element)
    {
        try
        {
            _ = element.Properties.NativeWindowHandle.Value;
            return true;
        }
        catch
        {
            // NativeWindowHandle is not available for all element types (e.g. cross-process
            // windows may not expose it). Fall through to the ProcessId check below.
        }
        try
        {
            _ = element.Properties.ProcessId.Value;
            return true;
        }
        catch { return false; }
    }

    /// <summary>Returns the native window handle of an element, or IntPtr.Zero on failure.</summary>
    private static IntPtr SafeWindowHandle(AutomationElement element)
    {
        try { return element.Properties.NativeWindowHandle.Value; }
        catch { return IntPtr.Zero; }
    }

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
    /// When <see cref="UiLocator.XPath"/> is set it is evaluated via <see cref="FindByXPath"/>;
    /// otherwise the standard attribute-condition approach is used.
    /// </summary>
    private static AutomationElement? TryFindElement(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(locator.XPath))
                return FindByXPath(root, session, locator.XPath);

            var condition = BuildCondition(session, locator);
            return root.FindFirstDescendant(condition);
        }
        catch
        {
            return null;
        }
    }

    // =========================================================================
    // XPath element finding
    // =========================================================================

    /// <summary>Represents one step in a parsed XPath expression.</summary>
    private sealed class XPathStep
    {
        public bool IsDescendant { get; set; }
        public string NodeName { get; set; } = "*";
        public Dictionary<string, string> Attrs { get; } =
            new(StringComparer.OrdinalIgnoreCase);
        public int? Index { get; set; }
    }

    /// <summary>
    /// Evaluates a simple XPath expression against the UIA tree rooted at
    /// <paramref name="root"/> and returns the first matching element, or
    /// <see langword="null"/> when no element matches.
    /// </summary>
    private static AutomationElement? FindByXPath(
        AutomationElement root, AutomationSession session, string xpath)
    {
        var steps = ParseXPath(xpath);
        var cf = session.Automation.ConditionFactory;

        IReadOnlyList<AutomationElement> current = [root];

        foreach (var step in steps)
        {
            var next = new List<AutomationElement>();
            bool hasFilter = step.NodeName != "*" || step.Attrs.Count > 0;

            foreach (var parent in current)
            {
                AutomationElement[] found;
                if (hasFilter)
                {
                    var condition = BuildStepCondition(cf, step);
                    found = step.IsDescendant
                        ? parent.FindAllDescendants(condition)
                        : parent.FindAllChildren(condition);
                }
                else
                {
                    found = step.IsDescendant
                        ? parent.FindAllDescendants()
                        : parent.FindAllChildren();
                }

                if (step.Index.HasValue)
                {
                    int idx = step.Index.Value - 1; // 1-based → 0-based
                    if (idx >= 0 && idx < found.Length)
                        next.Add(found[idx]);
                }
                else
                {
                    next.AddRange(found);
                }
            }

            current = next;
            if (current.Count == 0) return null;
        }

        return current.Count > 0 ? current[0] : null;
    }

    /// <summary>
    /// Parses an XPath string into a list of <see cref="XPathStep"/> objects.
    /// Supports <c>//</c> (descendant) and <c>/</c> (child) separators, node-name
    /// filters, <c>[@Attr='value']</c> predicates (joined by <c>and</c>), and
    /// <c>[n]</c> 1-based index selection.
    /// </summary>
    private static List<XPathStep> ParseXPath(string xpath)
    {
        xpath = xpath.Trim();
        if (string.IsNullOrEmpty(xpath))
            throw new ArgumentException("XPath expression is empty.");

        var steps = new List<XPathStep>();
        int i = 0;
        bool firstToken = true;

        while (i < xpath.Length)
        {
            bool isDescendant;

            if (firstToken)
            {
                firstToken = false;
                if (i + 1 < xpath.Length && xpath[i] == '/' && xpath[i + 1] == '/')
                { isDescendant = true; i += 2; }
                else if (xpath[i] == '/')
                { isDescendant = false; i += 1; }
                else
                { isDescendant = true; } // no leading slash → treat as //
            }
            else
            {
                if (i + 1 < xpath.Length && xpath[i] == '/' && xpath[i + 1] == '/')
                { isDescendant = true; i += 2; }
                else if (xpath[i] == '/')
                { isDescendant = false; i += 1; }
                else
                    break; // trailing slash or parse done
            }

            if (i >= xpath.Length) break;

            var step = new XPathStep { IsDescendant = isDescendant };

            // --- parse node name (up to '[' or '/') ---
            int nameStart = i;
            while (i < xpath.Length && xpath[i] != '[' && xpath[i] != '/')
                i++;
            step.NodeName = xpath[nameStart..i].Trim();
            if (string.IsNullOrEmpty(step.NodeName))
                step.NodeName = "*";

            // --- parse predicates: [@Attr='val'] or [n] ---
            while (i < xpath.Length && xpath[i] == '[')
            {
                i++; // skip '['
                int start = i;
                int depth = 1;
                bool inSingle = false, inDouble = false;

                while (i < xpath.Length && depth > 0)
                {
                    char c = xpath[i];
                    if (!inSingle && !inDouble)
                    {
                        if (c == '\'') inSingle = true;
                        else if (c == '"') inDouble = true;
                        else if (c == '[') depth++;
                        else if (c == ']') { depth--; if (depth == 0) break; }
                    }
                    else if (inSingle && c == '\'') inSingle = false;
                    else if (inDouble && c == '"') inDouble = false;
                    i++;
                }

                var content = xpath[start..i].Trim();
                i++; // skip ']'

                if (int.TryParse(content, out int idx))
                    step.Index = idx;
                else
                    ParsePredicateContent(content, step.Attrs);
            }

            steps.Add(step);
        }

        if (steps.Count == 0)
            throw new ArgumentException($"No valid steps found in XPath expression: '{xpath}'");

        return steps;
    }

    /// <summary>
    /// Parses the content of a <c>[@…]</c> predicate (possibly joined by <c>and</c>)
    /// and adds the resulting attribute key/value pairs to <paramref name="attrs"/>.
    /// </summary>
    private static void ParsePredicateContent(
        string content, Dictionary<string, string> attrs)
    {
        foreach (var raw in SplitByAnd(content))
        {
            var part = raw.Trim();
            if (!part.StartsWith('@'))
                throw new ArgumentException(
                    $"Predicate '{part}' is not a valid attribute predicate. " +
                    "Use @AttributeName='value' (e.g. @Name='OK' or @AutomationId='btn1').");

            int eq = part.IndexOf('=');
            if (eq < 0)
                throw new ArgumentException(
                    $"Predicate '{part}' is missing '='.");

            var name  = part[1..eq].Trim();
            var value = part[(eq + 1)..].Trim();

            if ((value.StartsWith('\'') && value.EndsWith('\'')) ||
                (value.StartsWith('"')  && value.EndsWith('"')))
                value = value[1..^1];

            attrs[name] = value;
        }
    }

    /// <summary>
    /// Splits a predicate string on the keyword <c> and </c> (case-insensitive),
    /// respecting single- and double-quoted substrings.
    /// </summary>
    private static List<string> SplitByAnd(string content)
    {
        var result = new List<string>();
        int start = 0;
        bool inSingle = false, inDouble = false;

        for (int i = 0; i < content.Length; i++)
        {
            char c = content[i];

            if (!inSingle && !inDouble)
            {
                if (c == '\'') { inSingle = true; continue; }
                if (c == '"')  { inDouble = true; continue; }

                if (i + 5 < content.Length &&
                    content[i..].StartsWith(" and ", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(content[start..i]);
                    i += 4; // advance past " and " (the for-loop will add 1 more, landing on the next predicate)
                    start = i + 1;
                }
            }
            else if (inSingle && c == '\'') inSingle = false;
            else if (inDouble && c == '"')  inDouble = false;
        }

        result.Add(content[start..]);
        return result;
    }

    /// <summary>
    /// Builds a FlaUI condition from a single <see cref="XPathStep"/>.
    /// Only called when the step has at least one filter (node name or attribute);
    /// wildcard steps with no attributes use the no-arg <c>FindAllDescendants</c>/
    /// <c>FindAllChildren</c> overload in <see cref="FindByXPath"/> instead.
    /// </summary>
    private static ConditionBase BuildStepCondition(
        ConditionFactory cf, XPathStep step)
    {
        var conditions = new List<ConditionBase>();

        if (step.NodeName != "*")
            conditions.Add(cf.ByControlType(ParseControlType(step.NodeName)));

        foreach (var (key, value) in step.Attrs)
        {
            conditions.Add(key.ToLowerInvariant() switch
            {
                "name"         => cf.ByName(value),
                "automationid" => cf.ByAutomationId(value),
                "classname"    => cf.ByClassName(value),
                "controltype"  => cf.ByControlType(ParseControlType(value)),
                _ => throw new ArgumentException(
                    $"Unsupported XPath attribute '@{key}'. " +
                    "Supported: @Name, @AutomationId, @ClassName, @ControlType.")
            });
        }

        if (conditions.Count == 0)
            throw new InvalidOperationException(
                "BuildStepCondition requires at least one filter; " +
                "call FindAllDescendants()/FindAllChildren() directly for wildcard steps.");

        return conditions.Count == 1
                ? conditions[0]
                : new AndCondition(conditions.ToArray());
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
                "(name, automationId, className, controlType, or xpath).");

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
        if (!string.IsNullOrWhiteSpace(l.XPath))         parts.Add($"xpath='{l.XPath}'");
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
