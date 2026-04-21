using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core;
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
            "closewindow"  => CloseWindowByTitle(request),
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
            "select"          => Select(request),
            "selectaid"       => SelectByAid(request),
            "typeandselect"   => TypeAndSelect(request),
            "clickgridcell"   => ClickGridCell(request),
            "doubleclickgridcell" => DoubleClickGridCell(request),
            "draganddrop"     => DragAndDrop(request),

            // ----- Alert / Dialog Handling -----
            "alertok"     => AlertOk(request),
            "alertcancel" => AlertCancel(request),
            "alertclose"  => AlertClose(request),

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

    /// <summary>
    /// Finds the first window (at any level in the UIA tree) whose title contains the
    /// value of <c>req.Value</c> (case-insensitive) and closes it.  Works whether or not a
    /// session is active: when a session exists its automation instance is reused;
    /// otherwise a temporary <see cref="UIA3Automation"/> is created for the single lookup.
    /// Top-level desktop children are checked first for performance; nested child/owned
    /// windows (e.g. dialogs opened from a child window) are found via a descendant search.
    ///
    /// Close strategy (in priority order):
    /// 1. Search for a Button with <c>AutomationId="Close"</c> and <c>Name="Close"</c>
    ///    inside the matched window and invoke/click it.
    /// 2. Fall back to the UIA Window pattern <c>Close()</c> method when no such button
    ///    is present.
    /// </summary>
    private object? CloseWindowByTitle(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' must be a partial window title for 'closewindow'.");

        // Prefer the active session's automation so we avoid spinning up a second COM object.
        // Use the session's application windows first so that windows opened through the driver
        // are found before falling back to a full desktop scan.
        var session = _ctx.ActiveSession;
        if (session != null)
        {
            var cf = session.Automation.ConditionFactory;
            var match = FindWindowByTitle(session, req.Value);

            if (match == null)
            {
                // Not among the session's own windows; widen the search to all desktop windows.
                var desktop = session.Automation.GetDesktop();
                match = FindWindowByTitle(desktop, cf, req.Value);
            }

            if (match != null)
            {
                CloseWindowElement(match, cf);
                return null;
            }
        }

        // No active session, or the window was not found: fall back to a temporary automation.
        using var tempAutomation = new UIA3Automation();
        var tempCf = tempAutomation.ConditionFactory;
        var tempDesktop = tempAutomation.GetDesktop();
        var tempMatch = FindWindowByTitle(tempDesktop, tempCf, req.Value);

        if (tempMatch == null)
            throw new InvalidOperationException(
                $"No window with title containing '{SanitizeValue(req.Value)}' was found.");

        // Close is called inside the using block so the automation instance is still live.
        CloseWindowElement(tempMatch, tempCf);
        return null;
    }

    /// <summary>
    /// Finds the first Window element whose Name contains <paramref name="title"/>
    /// (case-insensitive) by first checking direct children of <paramref name="root"/>
    /// (fast path for top-level windows) and then falling back to a full descendant
    /// search that also finds nested child/owned windows in multi-level hierarchies.
    /// </summary>
    private static AutomationElement? FindWindowByTitle(
        AutomationElement root, ConditionFactory cf, string title)
    {
        var windowCondition = cf.ByControlType(ControlType.Window);

        // Fast path: top-level windows are direct children of the Desktop.
        var topLevel = root.FindAllChildren(windowCondition)
            .FirstOrDefault(w => w.Name.Contains(title, StringComparison.OrdinalIgnoreCase));
        if (topLevel != null) return topLevel;

        // Fallback: owned/child windows may be nested beneath their parent in the UIA tree
        // (e.g. a dialog opened from a child window).
        return root.FindAllDescendants(windowCondition)
            .FirstOrDefault(w => w.Name.Contains(title, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Closes <paramref name="window"/> by first searching for a Button element with
    /// <c>AutomationId="Close"</c> and <c>Name="Close"</c> inside the window and
    /// invoking/clicking it.  Falls back to <see cref="CloseElement"/> (UIA Window
    /// pattern) when no such close button is found.
    /// </summary>
    private static void CloseWindowElement(AutomationElement window, ConditionFactory cf)
    {
        // Primary strategy: find the dedicated close button by its known AutomationId and Name.
        var closeButton = window.FindFirstDescendant(
            new AndCondition(
                cf.ByControlType(ControlType.Button),
                cf.ByAutomationId("Close"),
                cf.ByName("Close")));

        if (closeButton != null)
        {
            if (closeButton.Patterns.Invoke.IsSupported)
                closeButton.Patterns.Invoke.Pattern.Invoke();
            else
                closeButton.Click();
            return;
        }

        // Fallback: use the UIA Window pattern (or simulated title-bar click).
        CloseElement(window);
    }

    /// <summary>
    /// Closes <paramref name="element"/> via the UIA Window pattern's Close method.
    /// Falls back to simulating a click on the title-bar close button when the pattern
    /// is not supported.
    /// </summary>
    private static void CloseElement(AutomationElement element)
    {
        if (element.Patterns.Window.IsSupported)
            element.Patterns.Window.Pattern.Close();
        else
            element.AsWindow().Close();
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

    /// <inheritdoc/>
    public string? TakeFailureScreenshot(string directory)
    {
        try
        {
            Directory.CreateDirectory(directory);

            var filePath = Path.Combine(directory,
                $"failure_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");

            // When a session is active, capture only the application window.
            // Fall back to the full primary screen when no session exists.
            var session = _ctx.ActiveSession;
            if (session != null)
            {
                try
                {
                    var root = GetWindowRoot(session);
                    var rect = root.BoundingRectangle;
                    if (rect.Width > 0 && rect.Height > 0)
                    {
                        using var bmp = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
                        using var g = Graphics.FromImage(bmp);
                        g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(rect.Width, rect.Height));
                        bmp.Save(filePath, ImageFormat.Png);
                        return filePath;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "TakeFailureScreenshot: window capture failed, falling back to primary screen");
                }
            }

            // Full primary-screen fallback.
            // SystemInformation.VirtualScreen covers all monitors; fall back to PrimaryScreen
            // bounds and finally to the virtual-screen extent reported by System.Windows.Forms.
            var screen = System.Windows.Forms.Screen.PrimaryScreen?.Bounds
                ?? System.Windows.Forms.SystemInformation.VirtualScreen;
            using var screenBmp = new Bitmap(screen.Width, screen.Height, PixelFormat.Format32bppArgb);
            using var screenG = Graphics.FromImage(screenBmp);
            screenG.CopyFromScreen(screen.Left, screen.Top, 0, 0, new Size(screen.Width, screen.Height));
            screenBmp.Save(filePath, ImageFormat.Png);
            return filePath;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "TakeFailureScreenshot failed; screenshot was not saved");
            return null;
        }
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
        var session = RequireSession();
        var element = FindWithRetry(req);
        return ReadTableData(element, headersOnly: false, session);
    }

    private object? GetTableHeaders(UiRequest req)
    {
        var session = RequireSession();
        var element = FindWithRetry(req);
        return ReadTableData(element, headersOnly: true, session);
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

    private object? TypeAndSelect(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' is required for 'typeandselect'.");

        var element = FindWithRetry(req);
        var session = RequireSession();
        var cf = session.Automation.ConditionFactory;

        // For the structure Group → ComboBox → Edit, the user may target either
        // the ComboBox or its inner Edit child. Typing must go into the Edit, but
        // the ListItem descendants and the ExpandCollapse pattern live on the ComboBox.
        // If the located element is an Edit, walk up to find the parent ComboBox.
        var comboElement = element;
        if (element.ControlType == ControlType.Edit)
        {
            try
            {
                var walker = session.Automation.TreeWalkerFactory.GetControlViewWalker();
                var parent = walker.GetParent(element);
                if (parent?.ControlType == ControlType.ComboBox)
                    comboElement = parent;
            }
            catch { /* unable to walk tree; use the element as-is */ }
        }

        // Focus the Edit child (or the ComboBox itself if no Edit was found) so that
        // keyboard input lands in the correct text-entry area.
        element.Focus();

        // Type the filter text. Prefer the Value pattern (instant, no side effects);
        // fall back to simulated keystrokes so that the control fires its change events.
        bool typed = false;
        if (element.Patterns.Value.IsSupported
            && element.Patterns.Value.PatternOrDefault?.IsReadOnly == false)
        {
            try
            {
                element.Patterns.Value.Pattern.SetValue(req.Value);
                typed = true;
            }
            catch
            {
                // Value pattern may report writable but still throw (e.g. some native
                // combo hosts). Fall through to simulated keyboard input.
            }
        }

        if (!typed)
        {
            // Clear any existing content first so the filter starts fresh.
            try { element.Patterns.Value.PatternOrDefault?.SetValue(string.Empty); }
            catch { }
            Keyboard.Type(req.Value);
        }

        // Some editable comboboxes do not auto-expand when text is set programmatically.
        // If no items appear within the first poll, explicitly try to expand the dropdown.
        bool expandAttempted = false;

        // Wait up to 5 s for filtered ListItems to materialise in the dropdown.
        AutomationElement? target = null;
        var deadline = DateTime.UtcNow + DefaultRetry;
        while (true)
        {
            var items = comboElement.FindAllDescendants(cf.ByControlType(ControlType.ListItem));
            if (items.Length == 0)
            {
                // Some implementations nest items inside a List child.
                var listChild = comboElement.FindFirstDescendant(cf.ByControlType(ControlType.List));
                if (listChild != null)
                    items = listChild.FindAllChildren();
            }

            if (items.Length > 0)
            {
                // Prefer an exact-name match; otherwise fall back to the first visible item.
                target = items.FirstOrDefault(i =>
                    i.Name.Equals(req.Value, StringComparison.OrdinalIgnoreCase))
                    ?? items[0];
                break;
            }

            if (!expandAttempted)
            {
                // Try to force the dropdown open; some combo implementations need this
                // when the value is set programmatically rather than by keyboard events.
                try { comboElement.Patterns.ExpandCollapse.PatternOrDefault?.Expand(); }
                catch { /* best effort */ }
                expandAttempted = true;
            }

            if (DateTime.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"No dropdown items appeared after typing '{SanitizeValue(req.Value)}' " +
                    $"within {DefaultRetry.TotalSeconds}s.");

            Thread.Sleep(RetryInterval);
        }

        target.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();

        var selectionItemPattern = target.Patterns.SelectionItem.PatternOrDefault;
        if (selectionItemPattern != null)
            selectionItemPattern.Select();
        else
            target.Click();

        Thread.Sleep(100);

        // Collapse to commit the selection.
        comboElement.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();

        return null;
    }

    private object? DragAndDrop(UiRequest req)
    {
        if (req.Locator == null)
            throw new ArgumentException("'locator' (source element) is required for 'dragAndDrop'.");
        if (req.Locator2 == null)
            throw new ArgumentException("'locator2' (target element) is required for 'dragAndDrop'.");

        var session = RequireSession();
        var root = GetWindowRoot(session);

        var source = FindLocatorWithRetry(session, root, req.Locator);
        var target = FindLocatorWithRetry(session, root, req.Locator2);

        var srcPt = source.GetClickablePoint();
        var dstPt = target.GetClickablePoint();

        Mouse.Drag(srcPt, dstPt);

        return null;
    }

    // =========================================================================
    // Alert / Dialog Handling
    // =========================================================================

    /// <summary>
    /// Finds the topmost modal dialog and clicks its OK/Yes/Save button.
    /// Returns success without error when no dialog is present.
    /// </summary>
    private object? AlertOk(UiRequest _)
    {
        HandleAlert(["OK", "Yes", "&OK", "&Yes", "Save"], closeOnly: false);
        return new { success = true };
    }

    /// <summary>
    /// Finds the topmost modal dialog and clicks its Cancel/No button.
    /// Returns success without error when no dialog is present.
    /// </summary>
    private object? AlertCancel(UiRequest _)
    {
        HandleAlert(["Cancel", "No", "&Cancel", "&No"], closeOnly: false);
        return new { success = true };
    }

    /// <summary>
    /// Finds the topmost modal dialog and closes it via the Window pattern.
    /// Returns success without error when no dialog is present.
    /// </summary>
    private object? AlertClose(UiRequest _)
    {
        HandleAlert([], closeOnly: true);
        return new { success = true };
    }

    /// <summary>
    /// Locates the first modal dialog window visible on the desktop and either
    /// clicks a button whose name matches one of <paramref name="buttonNames"/> or
    /// closes the window when <paramref name="closeOnly"/> is true.
    /// Returns silently (without throwing) when no dialog is found so that
    /// callers can treat a missing alert as a no-op.
    /// </summary>
    private void HandleAlert(string[] buttonNames, bool closeOnly)
    {
        try
        {
            var session = _ctx.ActiveSession;
            if (session != null)
                PerformHandleAlert(session.Automation, buttonNames, closeOnly);
            else
            {
                using var tempAutomation = new UIA3Automation();
                PerformHandleAlert(tempAutomation, buttonNames, closeOnly);
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "HandleAlert: no alert found or action failed (treated as no-op)");
        }
    }

    /// <summary>
    /// Core alert-handling logic that operates on the supplied <paramref name="automation"/>
    /// instance so it works both with an active session and with a temporary automation.
    /// </summary>
    private static void PerformHandleAlert(AutomationBase automation, string[] buttonNames, bool closeOnly)
    {
        var cf = automation.ConditionFactory;
        var desktop = automation.GetDesktop();

        // Find the first modal dialog among all top-level windows on the desktop.
        var dialog = desktop
            .FindAllChildren(cf.ByControlType(ControlType.Window))
            .FirstOrDefault(IsModalDialog);

        if (dialog == null) return;

        if (!closeOnly && buttonNames.Length > 0)
        {
            // Look for a button whose name matches one of the candidate names.
            var btn = dialog
                .FindAllDescendants(cf.ByControlType(ControlType.Button))
                .FirstOrDefault(b => buttonNames.Any(n =>
                    string.Equals(b.Name, n, StringComparison.OrdinalIgnoreCase)));

            if (btn != null)
            {
                if (btn.Patterns.Invoke.IsSupported)
                    btn.Patterns.Invoke.Pattern.Invoke();
                else
                    btn.Click();
                return;
            }
        }

        // Close the dialog (fallback or alertclose).
        CloseElement(dialog);
    }

    /// <summary>
    /// Returns true when <paramref name="w"/> is a modal dialog window.
    /// Any exception from accessing UIA properties (e.g. stale element reference or
    /// COM error) is treated as "not a modal dialog" so the caller can continue safely.
    /// </summary>
    private static bool IsModalDialog(AutomationElement w)
    {
        try { return w.Patterns.Window.IsSupported && w.Patterns.Window.Pattern.IsModal; }
        catch { return false; }
    }

    private object? ClickGridCell(UiRequest req)
    {
        PerformGridCellAction(req, doubleClick: false);
        return null;
    }

    private object? DoubleClickGridCell(UiRequest req)
    {
        PerformGridCellAction(req, doubleClick: true);
        return null;
    }

    /// <summary>
    /// Locates the cell at <c>req.Index</c> (row) / <c>req.ColumnIndex</c> (column) and
    /// performs either a single click or a double-click depending on <paramref name="doubleClick"/>.
    ///
    /// Two grid layouts are supported:
    /// <list type="bullet">
    ///   <item>
    ///     <description>
    ///       <b>Grid pattern</b> – the element exposes <c>IGridProvider</c> (e.g. WinForms
    ///       DataGridView).  <c>IGridProvider.GetItem(row, col)</c> retrieves the cell directly.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       <b>DataItem rows</b> – the element does not expose the Grid pattern but its
    ///       children are <c>ControlType.DataItem</c> (e.g. WPF DataGrid without Grid
    ///       virtualization enabled).  The row is selected by index and the cell by child
    ///       position within that row.
    ///     </description>
    ///   </item>
    /// </list>
    /// </summary>
    private void PerformGridCellAction(UiRequest req, bool doubleClick)
    {
        if (req.Index == null)
            throw new ArgumentException("'index' (row index) is required for this operation.");
        if (req.ColumnIndex == null)
            throw new ArgumentException("'columnIndex' is required for this operation.");

        var element = FindWithRetry(req);

        int row = req.Index.Value;
        int col = req.ColumnIndex.Value;

        if (row < 0)
            throw new ArgumentException($"Row index {row} must be >= 0.");
        if (col < 0)
            throw new ArgumentException($"Column index {col} must be >= 0.");

        // ── Strategy 1: Grid pattern (WinForms DataGridView, etc.) ──────────────
        if (element.Patterns.Grid.IsSupported)
        {
            var grid = element.Patterns.Grid.Pattern;

            // RowCount/ColumnCount report 0 for virtualised grids even when data is present;
            // skip the upper-bound check in that case.
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

            ActOnCell(cell, doubleClick);
            return;
        }

        // ── Strategy 2: DataItem rows (WPF DataGrid, etc.) ──────────────────────
        // When the Grid pattern is absent the rows are often DataItem elements.
        // Navigate to the row by index, then to the cell by child position.
        var session = RequireSession();
        var cf = session.Automation.ConditionFactory;
        var dataItems = element.FindAllChildren(cf.ByControlType(ControlType.DataItem));

        if (dataItems.Length > 0)
        {
            if (row >= dataItems.Length)
                throw new ArgumentException(
                    $"Row index {row} is out of range. Grid has {dataItems.Length} row(s).");

            ActOnDataItemCell(dataItems[row], col, doubleClick);
            return;
        }

        // ── Strategy 2b: locator resolved to a DataItem row instead of the container ──
        // When the locator targets controlType=DataItem, FindFirstDescendant returns the
        // first DataItem row (not the grid container).  Navigate up to the parent and
        // collect all sibling DataItem rows from there.
        if (element.ControlType == ControlType.DataItem && element.Parent != null)
        {
            var siblingItems = element.Parent.FindAllChildren(cf.ByControlType(ControlType.DataItem));
            if (row >= siblingItems.Length)
                throw new ArgumentException(
                    $"Row index {row} is out of range. Grid has {siblingItems.Length} row(s).");

            ActOnDataItemCell(siblingItems[row], col, doubleClick);
            return;
        }

        throw new ArgumentException(
            "The target element does not support the Grid pattern and has no DataItem rows. " +
            "Verify the locator targets the grid/table control itself, not a child row or cell.");
    }

    /// <summary>Performs a click or double-click on a grid cell element.</summary>
    private static void ActOnCell(AutomationElement cell, bool doubleClick)
    {
        if (doubleClick)
        {
            cell.DoubleClick();
            return;
        }

        if (cell.Patterns.Invoke.IsSupported)
            cell.Patterns.Invoke.Pattern.Invoke();
        else
            cell.Click();
    }

    /// <summary>
    /// Locates the child of <paramref name="dataItem"/> at <paramref name="colIndex"/>
    /// and performs a click or double-click.
    /// </summary>
    private static void ActOnDataItemCell(AutomationElement dataItem, int colIndex, bool doubleClick)
    {
        var children = dataItem.FindAllChildren();
        if (colIndex >= children.Length)
            throw new ArgumentException(
                $"Column index {colIndex} is out of range. Row has {children.Length} cell(s).");

        ActOnCell(children[colIndex], doubleClick);
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
            "Valid values: Button, CheckBox, ComboBox, Custom, DataGrid, DataItem, Edit, " +
            "Group, Header, HeaderItem, List, ListItem, Menu, MenuItem, Pane, RadioButton, " +
            "Slider, Spinner, Tab, TabItem, Table, Text, ToolBar, Tree, TreeItem, Window.");
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
    /// Reads table/grid data from an element using multiple fallback strategies.
    ///
    /// Strategy 1 – Table/Grid UIA patterns (standard controls).
    /// Strategy 2 – Direct descendant search for Header and DataItem control types.
    /// Strategy 3 – Search Custom control descendants of <paramref name="element"/>
    ///              and dig inside each one for Header/DataItem.
    /// Strategy 4 – Window-level fallback: search the entire window for Custom
    ///              control containers when the element itself yields no data.
    /// </summary>
    private object ReadTableData(AutomationElement element, bool headersOnly, AutomationSession session)
    {
        var cf = session.Automation.ConditionFactory;
        var headers = new List<string>();
        var rows    = new List<List<string>>();

        // ── Strategy 1: Standard Table / Grid UIA patterns ───────────────────
        TableViaPatterns(element, headersOnly, headers, rows);

        // ── Strategy 2: Direct Header / DataItem descendant search ───────────
        if (headers.Count == 0)
            HeadersFromDescendants(element, cf, headers);

        if (!headersOnly && rows.Count == 0)
            DataItemsFromDescendants(element, cf, rows);

        // ── Strategy 3: Deep-dive into Custom control descendants ─────────────
        if (headers.Count == 0 && (headersOnly || rows.Count == 0))
            CustomControlDeepDive(element, headersOnly, cf, headers, rows);

        // ── Strategy 4: Window-level Custom control fallback ─────────────────
        if (headers.Count == 0 && (headersOnly || rows.Count == 0))
        {
            var windowRoot = GetWindowRoot(session);
            if (!ReferenceEquals(element, windowRoot))
                CustomControlDeepDive(windowRoot, headersOnly, cf, headers, rows);
        }

        if (headersOnly)
            return new { headers };

        return new { headers, rows };
    }

    /// <summary>
    /// Strategy 1: extract column headers via the Table UIA pattern and rows via
    /// the Grid UIA pattern. Only populates <paramref name="headers"/> and
    /// <paramref name="rows"/> when the respective pattern is supported.
    /// </summary>
    private static void TableViaPatterns(
        AutomationElement element, bool headersOnly,
        List<string> headers, List<List<string>> rows)
    {
        if (element.Patterns.Table.IsSupported)
        {
            var table = element.Patterns.Table.Pattern;
            foreach (var h in table.ColumnHeaders.Value)
                headers.Add(h.Name ?? string.Empty);
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
                    row.Add(cell != null ? GetCellText(cell) : string.Empty);
                }
                rows.Add(row);
            }
        }
    }

    /// <summary>
    /// Strategy 2: locate a Header control type among element's descendants and
    /// add the names of its children (HeaderItem elements) to <paramref name="headers"/>.
    /// </summary>
    private static void HeadersFromDescendants(
        AutomationElement element, ConditionFactory cf, List<string> headers)
    {
        var headerEl = element.FindFirstDescendant(cf.ByControlType(ControlType.Header));
        if (headerEl == null) return;

        foreach (var hi in headerEl.FindAllChildren())
            headers.Add(hi.Name ?? string.Empty);
    }

    /// <summary>
    /// Strategy 2: locate all DataItem descendants and add each one's children
    /// (the cells) as a row. Cell content is read via the Value UIA pattern first
    /// (preferred for editable or text-displaying cell implementations), then falls
    /// back to the element's Name property.
    /// </summary>
    private static void DataItemsFromDescendants(
        AutomationElement element, ConditionFactory cf, List<List<string>> rows)
    {
        var dataItems = element.FindAllDescendants(cf.ByControlType(ControlType.DataItem));
        foreach (var di in dataItems)
        {
            var cells = di.FindAllChildren();
            rows.Add(cells.Select(GetCellText).ToList());
        }
    }

    /// <summary>
    /// Returns the best available text for a table cell element: tries the Value
    /// UIA pattern first (covers Edit/text-display cells), then falls back to the
    /// element's accessible Name, then AutomationId as a last resort.
    /// </summary>
    private static string GetCellText(AutomationElement cell)
    {
        try
        {
            if (cell.Patterns.Value.IsSupported)
            {
                var v = cell.Patterns.Value.Pattern.Value;
                if (!string.IsNullOrEmpty(v)) return v;
            }
        }
        catch { /* pattern unavailable or COM error */ }

        return cell.Name ?? string.Empty;
    }

    /// <summary>
    /// Strategy 3/4: find all Custom control type descendants of <paramref name="root"/>
    /// and, for each, search inside for Header (column headers) and DataItem (rows).
    /// Stops as soon as enough data has been gathered.
    /// </summary>
    private static void CustomControlDeepDive(
        AutomationElement root, bool headersOnly, ConditionFactory cf,
        List<string> headers, List<List<string>> rows)
    {
        AutomationElement[] customEls;
        try
        {
            customEls = root.FindAllDescendants(cf.ByControlType(ControlType.Custom));
        }
        catch
        {
            return;
        }

        foreach (var custom in customEls)
        {
            try
            {
                if (headers.Count == 0)
                    HeadersFromDescendants(custom, cf, headers);

                if (!headersOnly && rows.Count == 0)
                    DataItemsFromDescendants(custom, cf, rows);
            }
            catch { /* best effort */ }

            // Stop early when we have what we need.
            if (headers.Count > 0 && (headersOnly || rows.Count > 0))
                return;
        }
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
