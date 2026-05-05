using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
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

    /// <summary>
    /// Milliseconds to wait after bringing a window to the foreground before sending
    /// keyboard or mouse input, to allow the OS to finish the window-activation sequence.
    /// </summary>
    private const int WindowActivationDelayMs = 100;
    /// <summary>
    /// Brief pause after <c>SetCursorPos</c> so Windows finishes moving the cursor before
    /// the left-button input events are sent.
    /// </summary>
    private const int CursorPositionStabilityDelayMs = 30;
    private const int MenuExpandDelayMs = 250;
    private const int MenuActionDelayMs = 150;
    private const int MenuFocusDelayMs = 75;

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
            "switchwindow"  => SwitchWindow(request),
            "switchwinodw"  => SwitchWindow(request),
            "switch_window" => SwitchWindow(request),
            "switchto"      => SwitchWindow(request),
            "refresh"        => Refresh(),
            "screenshot"     => Screenshot(request),
            "listelements"   => ListElements(request),
            "listwindows"    => ListWindows(request),
            "getcurrentroot" => GetCurrentRoot(request),
            "findlocator"    => FindLocatorDebug(request),

            // ----- Element Query -----
            "exists"         => Exists(request),
            "waitfor"        => WaitFor(request),
            "isenabled"      => IsEnabled(request),
            "isvisible"      => IsVisible(request),
            "isclickable"    => IsClickable(request),
            "iseditable"     => IsEditable(request),
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
            "click"            => Click(request),
            "clickmenu"        => ClickMenuPath(request),
            "clickmenupath"    => IsLogicalMenuMode(request) ? ClickLogicalMenuPath(request) : ClickMenuPath(request),
            "clicklogicalmenupath" => ClickLogicalMenuPath(request),
            "clickmenulogical" => ClickLogicalMenuPath(request),
            "menupath"         => ClickLogicalMenuPath(request),
            "inspectlogicalmenu" => InspectLogicalMenu(request),
            "dumpmenus"        => DumpLogicalMenus(request),
            "dumplogicalmenus" => DumpLogicalMenus(request),
            "doubleclick"      => DoubleClick(request),
            "rightclick"       => RightClick(request),
            "hover"            => Hover(request),
            "focus"            => Focus(request),
            "type"             => TypeText(request),
            "clear"            => Clear(request),
            "sendkeys"         => SendKeys(request),
            "scroll"           => Scroll(request),
            "check"            => Check(request),
            "uncheck"          => Uncheck(request),
            "select"           => Select(request),
            "selectaid"        => SelectByAid(request),
            "typeandselect"    => TypeAndSelect(request),
            "clickgridcell"    => ClickGridCell(request),
            "doubleclickgridcell" => DoubleClickGridCell(request),
            "draganddrop"     => DragAndDrop(request),

            // ----- Alert / Dialog Handling -----
            "alertok"     => AlertOk(request),
            "alertcancel" => AlertCancel(request),
            "alertclose"  => AlertClose(request),
            "popupok"     => PopUpOk(request),

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
            .FirstOrDefault(w => TitleContains(w, title));
        if (topLevel != null) return topLevel;

        // Fallback: owned/child windows may be nested beneath their parent in the UIA tree
        // (e.g. a dialog opened from a child window).
        return root.FindAllDescendants(windowCondition)
            .FirstOrDefault(w => TitleContains(w, title));
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
            {
                try
                {
                    closeButton.Patterns.Invoke.Pattern.Invoke();
                    return;
                }
                catch (Exception ex) when (ex is FlaUI.Core.Exceptions.ElementNotAvailableException
                                        || ex is COMException)
                {
                    // Element became unavailable or invoke failed; fall through to mouse click.
                }
            }

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
                {
                    LogWindowSearchFailure(session.Automation, req.Value);
                    throw new InvalidOperationException(
                        $"No window with title containing '{SanitizeValue(req.Value)}' was found within {DefaultRetry.TotalSeconds}s.");
                }
                Thread.Sleep(RetryInterval);
            }

            return SwitchToWindow(session, match);
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
            {
                LogWindowSearchFailure(session.Automation, req.Value);
                throw new InvalidOperationException(
                    $"No window with title containing '{SanitizeValue(req.Value)}' was found.");
            }

            return SwitchToWindow(session, match);
        }
    }

    /// <summary>
    /// Searches the active window, application windows, and desktop window tree for one
    /// whose title contains <paramref name="titleFragment"/>.
    /// Returns null when no match is found.
    /// </summary>
    private static AutomationElement? FindWindowByTitle(AutomationSession session, string titleFragment)
    {
        var cf = session.Automation.ConditionFactory;
        var windowCondition = cf.ByControlType(ControlType.Window);

        if (session.ActiveWindow != null && TitleContains(session.ActiveWindow, titleFragment))
            return session.ActiveWindow;

        if (session.ActiveWindow != null)
        {
            try
            {
                var activeDescendant = session.ActiveWindow
                    .FindAllDescendants(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
                if (activeDescendant != null)
                    return activeDescendant;
            }
            catch
            {
                // Continue to broader search scopes.
            }
        }

        try
        {
            var mainWindow = session.Application.GetMainWindow(session.Automation);
            if (mainWindow != null)
            {
                if (TitleContains(mainWindow, titleFragment))
                    return mainWindow;

                var mainDescendant = mainWindow
                    .FindAllDescendants(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
                if (mainDescendant != null)
                    return mainDescendant;
            }
        }
        catch
        {
            // Continue to broader search scopes.
        }

        try
        {
            var topLevel = session.Application.GetAllTopLevelWindows(session.Automation)
                .FirstOrDefault(w => TitleContains(w, titleFragment));
            if (topLevel != null)
                return topLevel;
        }
        catch
        {
            // Continue to broader search scopes.
        }

        try
        {
            foreach (var appWindow in session.Application.GetAllTopLevelWindows(session.Automation))
            {
                var nested = appWindow
                    .FindAllDescendants(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
                if (nested != null)
                    return nested;
            }
        }
        catch
        {
            // Continue to broader search scopes.
        }

        AutomationElement? desktop = null;
        try
        {
            desktop = session.Automation.GetDesktop();
            var desktopChild = desktop
                .FindAllChildren(windowCondition)
                .FirstOrDefault(w => TitleContains(w, titleFragment));
            if (desktopChild != null)
                return desktopChild;
        }
        catch
        {
            // Continue to broader search scopes.
        }

        try
        {
            return desktop == null
                ? null
                : desktop
                .FindAllDescendants(windowCondition)
                .FirstOrDefault(w => TitleContains(w, titleFragment));
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Searches the desktop window tree for one whose title contains
    /// <paramref name="titleFragment"/> and returns its process ID.
    /// Retries until <paramref name="deadline"/> to handle windows that open with a delay.
    /// Throws <see cref="InvalidOperationException"/> if no match is found before the deadline.
    /// </summary>
    private static int FindWindowProcessId(string titleFragment, DateTime deadline)
    {
        using var tempAutomation = new UIA3Automation();
        var cf = tempAutomation.ConditionFactory;
        var windowCondition = cf.ByControlType(ControlType.Window);

        while (true)
        {
            var desktop = tempAutomation.GetDesktop();
            AutomationElement? match = null;

            try
            {
                match = desktop
                    .FindAllChildren(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
            }
            catch
            {
                // Continue to descendant search.
            }

            if (match == null)
            {
                try
                {
                    match = desktop
                        .FindAllDescendants(windowCondition)
                        .FirstOrDefault(w => TitleContains(w, titleFragment));
                }
                catch
                {
                    // Ignore unstable UIA trees during retry.
                }
            }

            if (match != null)
                return match.Properties.ProcessId.Value;

            if (DateTime.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"No window with title containing '{SanitizeValue(titleFragment)}' was found within {DefaultRetry.TotalSeconds}s.");

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

    private object SwitchToWindow(AutomationSession session, AutomationElement window)
    {
        var asWindow = window.AsWindow();
        session.ActiveWindow = asWindow;

        var switchedHandle = SafeWindowHandle(asWindow);
        if (switchedHandle != IntPtr.Zero)
            session.SeedWindowHandles([switchedHandle]);

        asWindow.SetForeground();
        Thread.Sleep(WindowActivationDelayMs);

        _logger.LogInformation(
            "Switched active window. title={Title}, automationId={AutomationId}, hwnd=0x{Hwnd:X}",
            SafeElementName(asWindow),
            SafeElementAutomationId(asWindow),
            SafeWindowHandle(asWindow).ToInt64());

        return new
        {
            switched = true,
            title = asWindow.Name,
            automationId = asWindow.AutomationId,
            hwnd = SafeWindowHandle(asWindow).ToInt64()
        };
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
            // Prefer PrimaryScreen bounds; fall back to SystemInformation.VirtualScreen
            // (which covers all monitors) when PrimaryScreen is unavailable.
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
        var cf = session.Automation.ConditionFactory;
        var windows = new List<AutomationElement>();
        var seenHandles = new HashSet<long>();

        try
        {
            foreach (var window in session.Application.GetAllTopLevelWindows(session.Automation))
            {
                windows.Add(window);

                var hwnd = SafeWindowHandle(window).ToInt64();
                if (hwnd != 0)
                    seenHandles.Add(hwnd);
            }
        }
        catch
        {
            // Ignore unstable application window enumerations.
        }

        try
        {
            var desktopDescendants = session.Automation.GetDesktop()
                .FindAllDescendants(cf.ByControlType(ControlType.Window));

            foreach (var window in desktopDescendants)
            {
                var hwnd = SafeWindowHandle(window).ToInt64();
                if (hwnd != 0 && seenHandles.Add(hwnd))
                {
                    windows.Add(window);
                }
            }
        }
        catch
        {
            // Ignore unstable desktop window enumerations.
        }

        IEnumerable<AutomationElement> filtered = windows;
        if (!string.IsNullOrWhiteSpace(req.Value))
            filtered = windows.Where(w => TitleContains(w, req.Value));

        return filtered.Select(w => new
        {
            title = w.Name,
            automationId = w.AutomationId,
            className = w.ClassName,
            processId = w.Properties.ProcessId.ValueOrDefault,
            hwnd = SafeWindowHandle(w).ToInt64(),
            isOffscreen = w.IsOffscreen
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

    private object? IsEditable(UiRequest req)
    {
        var element = FindWithRetry(req);
        return new { editable = IsElementEditable(element) };
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

        if (element.ControlType == ControlType.MenuItem)
            return ClickMenuItem(element, req);

        if (TryPhysicalClick(element, "Click"))
            return null;

        if (TryElementClick(element, "Click"))
            return null;

        if (TryInvokePattern(element, "Click"))
            return null;

        throw new InvalidOperationException(
            $"Click failed after trying physical click, FlaUI click, and InvokePattern for " +
            $"name='{SafeElementName(element)}' controlType={element.ControlType}");
    }

    private object? ClickMenuItem(AutomationElement menuItem, UiRequest _)
    {
        _logger.LogInformation(
            "ClickMenuItem requested. name={Name}, automationId={AutomationId}, controlType={ControlType}",
            SafeElementName(menuItem),
            SafeElementAutomationId(menuItem),
            menuItem.ControlType);

        BringElementWindowToForeground(menuItem);

        try
        {
            if (menuItem.Patterns.ExpandCollapse.IsSupported)
            {
                var state = menuItem.Patterns.ExpandCollapse.Pattern.ExpandCollapseState;
                if (state != ExpandCollapseState.Expanded)
                {
                    menuItem.Patterns.ExpandCollapse.Pattern.Expand();
                    Thread.Sleep(MenuExpandDelayMs);

                    _logger.LogInformation("MenuItem expanded: {Name}", SafeElementName(menuItem));
                    return null;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "MenuItem ExpandCollapse failed; trying click strategies");
        }

        try
        {
            if (menuItem.Patterns.SelectionItem.IsSupported)
            {
                menuItem.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(MenuActionDelayMs);
                return null;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "MenuItem SelectionItem.Select failed");
        }

        try
        {
            menuItem.Click();
            Thread.Sleep(MenuActionDelayMs);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "MenuItem element.Click failed");
        }

        try
        {
            menuItem.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(MenuActionDelayMs);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "MenuItem Enter fallback failed");
        }

        if (TryInstantPhysicalClick(menuItem, "MenuItemClick"))
            return null;

        throw new InvalidOperationException(
            $"MenuItem click failed for '{SafeElementName(menuItem)}'. " +
            "For submenu navigation use operation='clicklogicalmenupath' " +
            "or operation='clickmenupath' with locator.mode='logical'.");
    }

    private object? ClickMenuPath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' is required for clickmenupath. Example: File>Open>Recent");

        var session = RequireSession();
        var root = GetWindowRoot(session);
        var cf = session.Automation.ConditionFactory;
        var normalizedValue = System.Net.WebUtility.HtmlDecode(req.Value);
        var parts = normalizedValue
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (parts.Count == 0)
            throw new ArgumentException("clickmenupath value did not contain menu items.");

        AutomationElement searchRoot = root;

        for (int i = 0; i < parts.Count; i++)
        {
            var name = parts[i];
            var isLast = i == parts.Count - 1;

            var item = FindMenuItemByName(searchRoot, cf, name);
            if (item == null)
            {
                var desktop = session.Automation.GetDesktop();
                item = FindMenuItemByName(desktop, cf, name);
            }

            if (item == null)
            {
                throw new InvalidOperationException(
                    $"Menu path failed. Menu item '{name}' was not found. Path='{SanitizeValue(normalizedValue)}'");
            }

            _logger.LogInformation(
                "Menu path step {Index}/{Count}: clicking/expanding '{Name}'",
                i + 1,
                parts.Count,
                name);

            if (!isLast)
            {
                OpenMenuItem(item);
                Thread.Sleep(MenuExpandDelayMs);
                searchRoot = session.Automation.GetDesktop();
            }
            else
            {
                ActivateMenuItem(item);
                Thread.Sleep(MenuActionDelayMs);
            }
        }

        return null;
    }

    private object? ClickLogicalMenuPath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
        {
            throw new ArgumentException(
                "'value' is required for clicklogicalmenupath. Example: tag3>tag3Child1");
        }

        var session = RequireSession();
        var searchRoot = GetSearchRootForMenuOperation(req, session);
        if (IsDesktopRoot(searchRoot))
        {
            throw new InvalidOperationException(
                "menupath search root resolved to Desktop1. This is invalid. " +
                "Pass a MenuBar locator or switch to the application window first.");
        }

        var cf = session.Automation.ConditionFactory;
        var normalizedValue = System.Net.WebUtility.HtmlDecode(req.Value);
        var parts = SplitMenuPath(normalizedValue);

        if (parts.Count == 0)
            throw new ArgumentException("Menu path is empty.");

        _logger.LogInformation(
            "ClickLogicalMenuPath requested. path={Path}, searchRootName={RootName}, searchRootAutomationId={RootAutomationId}, searchRootControlType={RootControlType}",
            string.Join(" > ", parts),
            SafeElementName(searchRoot),
            SafeElementAutomationId(searchRoot),
            searchRoot.ControlType);

        var target = FindLogicalMenuPath(searchRoot, cf, parts);

        if (target == null)
        {
            throw new InvalidOperationException(
                $"Logical menu path not found: '{req.Value}'. " +
                "Do not expand parent first; child must be resolved from the logical MenuItem tree.");
        }

        _logger.LogInformation(
            "Logical menu target found. name={Name}, automationId={AutomationId}, controlType={ControlType}",
            SafeElementName(target),
            SafeElementAutomationId(target),
            target.ControlType);

        if (TryActivateLogicalMenuItem(target, "ClickLogicalMenuPath"))
            return null;

        throw new InvalidOperationException(
            $"Failed to activate logical menu item '{SafeElementName(target)}' for path '{req.Value}'.");
    }

    private object? GetCurrentRoot(UiRequest req)
    {
        _ = req;
        var session = RequireSession();
        var root = GetWindowRoot(session);

        return new
        {
            root = CreateElementSnapshot(root),
            activeWindow = session.ActiveWindow == null ? null : CreateElementSnapshot(session.ActiveWindow)
        };
    }

    private object? FindLocatorDebug(UiRequest req)
    {
        var session = RequireSession();
        var root = GetWindowRoot(session);
        var element = FindLocatorWithRetry(session, root, RequireLocator(req));

        return new
        {
            found = true,
            root = CreateElementSnapshot(root),
            element = CreateElementSnapshot(element)
        };
    }

    private object? InspectLogicalMenu(UiRequest req)
    {
        var session = RequireSession();
        var root = GetSearchRootForMenuOperation(req, session);
        var cf = session.Automation.ConditionFactory;
        var parentName = string.IsNullOrWhiteSpace(req.Value)
            ? null
            : System.Net.WebUtility.HtmlDecode(req.Value);

        var parents = root
            .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
            .Where(e => string.IsNullOrWhiteSpace(parentName) || MenuTextMatches(e, parentName))
            .ToList();

        if (parents.Count == 0)
            throw new InvalidOperationException($"MenuItem '{parentName}' not found.");

        return parents.Select(parent => new
        {
            parent = new
            {
                name = SafeElementName(parent),
                automationId = SafeElementAutomationId(parent),
                parentChain = BuildParentChain(parent)
            },
            children = parent
                .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                .Select(e => new
                {
                    name = SafeElementName(e),
                    automationId = SafeElementAutomationId(e),
                    controlType = e.ControlType.ToString(),
                    parentChain = BuildParentChain(e),
                    patterns = new
                    {
                        invoke = e.Patterns.Invoke.IsSupported,
                        expandCollapse = e.Patterns.ExpandCollapse.IsSupported,
                        selectionItem = e.Patterns.SelectionItem.IsSupported
                    }
                })
                .ToList()
        }).ToList();
    }

    private object? DumpLogicalMenus(UiRequest req)
    {
        var session = RequireSession();
        var root = GetSearchRootForMenuOperation(req, session);
        var cf = session.Automation.ConditionFactory;

        var menuItems = root
            .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
            .Select(e =>
            {
                try
                {
                    return new
                    {
                        name = SafeElementName(e),
                        automationId = SafeElementAutomationId(e),
                        controlType = e.ControlType.ToString(),
                        className = e.ClassName,
                        bounds = e.BoundingRectangle.ToString(),
                        parentChain = BuildParentChain(e)
                    };
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "DumpLogicalMenus skipped unstable menu item.");
                    return null;
                }
            })
            .Where(x => x != null)
            .ToList();

        return new
        {
            root = new
            {
                name = SafeElementName(root),
                automationId = SafeElementAutomationId(root),
                controlType = root.ControlType.ToString(),
                hwnd = SafeWindowHandle(root).ToInt64()
            },
            menuItems
        };
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

    private static AutomationElement? FindMenuItemByName(
        AutomationElement root,
        ConditionFactory cf,
        string name)
    {
        try
        {
            if (root.ControlType == ControlType.MenuItem &&
                string.Equals(root.Name?.Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return root;
            }

            return root.FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                .FirstOrDefault(e =>
                    string.Equals(e.Name?.Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return null;
        }
    }

    private static List<string> SplitMenuPath(string path)
    {
        // Canonical separator is '>', but logical menu paths may also arrive from
        // external tools or recordings using '/' or '|', so accept all three.
        return path
            .Split(['>', '|', '/'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToList();
    }

    private AutomationElement? FindLogicalMenuPath(
        AutomationElement root,
        ConditionFactory cf,
        IReadOnlyList<string> parts)
    {
        if (parts.Count == 0)
            return null;

        try
        {
            var allMenuItems = root
                .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                .ToList();

            _logger.LogInformation(
                "FindLogicalMenuPath: path={Path}, rootName={RootName}, totalMenuItems={Count}",
                string.Join(" > ", parts),
                SafeElementName(root),
                allMenuItems.Count);

            var parentCandidates = allMenuItems
                .Where(e => MenuTextMatches(e, parts[0]))
                .ToList();

            _logger.LogInformation(
                "FindLogicalMenuPath: parent candidates for '{Parent}' = {Candidates}",
                parts[0],
                parentCandidates.Select(x => new
                {
                    name = SafeElementName(x),
                    automationId = SafeElementAutomationId(x),
                    controlType = x.ControlType.ToString(),
                    parentChain = BuildParentChain(x)
                }).ToList());

            foreach (var candidate in parentCandidates)
            {
                var resolved = TryResolveMenuPathFromParent(candidate, cf, parts, 1);

                if (resolved != null)
                {
                    _logger.LogInformation(
                        "FindLogicalMenuPath: resolved path={Path} using parent name={ParentName}, automationId={ParentAutomationId}",
                        string.Join(" > ", parts),
                        SafeElementName(candidate),
                        SafeElementAutomationId(candidate));

                    return resolved;
                }
            }

            if (parts.Count >= 2)
            {
                var leaf = FindMenuItemByLeafAndAncestor(root, cf, parts[^1], parts[^2]);

                if (leaf != null)
                    return leaf;

                var leafName = parts[^1];
                var leafMatches = allMenuItems
                    .Where(e => MenuTextMatches(e, leafName))
                    .ToList();

                _logger.LogWarning(
                    "FindLogicalMenuPath: direct path failed. Leaf matches for '{Leaf}' = {Matches}",
                    leafName,
                    leafMatches.Select(x => new
                    {
                        name = SafeElementName(x),
                        automationId = SafeElementAutomationId(x),
                        parentChain = BuildParentChain(x)
                    }).ToList());

                if (leafMatches.Count == 1)
                {
                    _logger.LogInformation(
                        "FindLogicalMenuPath: using unique leaf fallback for '{Leaf}'",
                        leafName);

                    return leafMatches[0];
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "FindLogicalMenuPath failed for path={Path}",
                string.Join(" > ", parts));

            return null;
        }
    }

    private AutomationElement? TryResolveMenuPathFromParent(
        AutomationElement parent,
        ConditionFactory cf,
        IReadOnlyList<string> parts,
        int nextIndex)
    {
        var current = parent;

        for (var i = nextIndex; i < parts.Count; i++)
        {
            var nextName = parts[i];
            List<AutomationElement> descendants;

            try
            {
                descendants = current
                    .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                    .ToList();
            }
            catch (Exception ex)
            {
                _logger.LogDebug(
                    ex,
                    "TryResolveMenuPathFromParent failed while enumerating descendants for {Current}",
                    SafeElementName(current));
                return null;
            }

            _logger.LogDebug(
                "TryResolveMenuPathFromParent: current={Current}, lookingFor={Child}, descendants={Descendants}",
                SafeElementName(current),
                nextName,
                descendants.Select(x => new
                {
                    name = SafeElementName(x),
                    automationId = SafeElementAutomationId(x)
                }).ToList());

            var next = descendants.FirstOrDefault(e => MenuTextMatches(e, nextName));

            if (next == null)
                return null;

            current = next;
        }

        return current;
    }

    private AutomationElement? FindMenuItemByLeafAndAncestor(
        AutomationElement root,
        ConditionFactory cf,
        string leafName,
        string ancestorName)
    {
        try
        {
            var all = root.FindAllDescendants(cf.ByControlType(ControlType.MenuItem));

            foreach (var item in all)
            {
                if (!MenuTextMatches(item, leafName))
                    continue;

                var current = item.Parent;

                while (current != null)
                {
                    if (current.ControlType == ControlType.MenuItem &&
                        MenuTextMatches(current, ancestorName))
                    {
                        _logger.LogInformation(
                            "FindMenuItemByLeafAndAncestor matched leaf={Leaf}, ancestor={Ancestor}",
                            SafeElementName(item),
                            SafeElementName(current));

                        return item;
                    }

                    current = current.Parent;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindMenuItemByLeafAndAncestor failed");
        }

        return null;
    }

    private List<string> BuildParentChain(AutomationElement element)
    {
        var chain = new List<string>();

        try
        {
            var current = element.Parent;

            while (current != null)
            {
                var name = string.IsNullOrWhiteSpace(SafeElementName(current))
                    ? "<empty>"
                    : SafeElementName(current);
                var automationId = string.IsNullOrWhiteSpace(SafeElementAutomationId(current))
                    ? "<empty>"
                    : SafeElementAutomationId(current);

                chain.Add($"{current.ControlType}:{name}:{automationId}");
                current = current.Parent;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "BuildParentChain failed for {Element}", SafeElementName(element));
        }

        return chain;
    }

    private static bool MenuTextMatches(AutomationElement element, string expected)
    {
        try
        {
            var expectedNorm = NormalizeMenuText(expected);
            var nameNorm = NormalizeMenuText(element.Name);
            var automationIdNorm = NormalizeMenuText(element.AutomationId);

            if (string.IsNullOrWhiteSpace(expectedNorm))
                return false;

            var isExactMatch =
                string.Equals(nameNorm, expectedNorm, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(automationIdNorm, expectedNorm, StringComparison.OrdinalIgnoreCase);

            if (isExactMatch)
                return true;

            return nameNorm.Contains(expectedNorm, StringComparison.OrdinalIgnoreCase) ||
                   automationIdNorm.Contains(expectedNorm, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizeMenuText(string? value)
    {
        var normalized = (value ?? string.Empty)
            .Replace("&amp;", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("&", string.Empty)
            .Replace("_", string.Empty)
            .Replace("\u00A0", " ")
            .Trim();

        return System.Text.RegularExpressions.Regex.Replace(normalized, @"\s{2,}", " ");
    }

    private bool TryActivateLogicalMenuItem(AutomationElement item, string actionName)
    {
        BringElementWindowToForeground(item);

        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();

                _logger.LogInformation(
                    "{ActionName}: InvokePattern succeeded for logical menu item {Name}",
                    actionName,
                    SafeElementName(item));

                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: InvokePattern failed for logical menu item {Name}",
                actionName,
                SafeElementName(item));
        }

        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();

                _logger.LogInformation(
                    "{ActionName}: SelectionItem.Select succeeded for logical menu item {Name}",
                    actionName,
                    SafeElementName(item));

                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: SelectionItem failed for logical menu item {Name}",
                actionName,
                SafeElementName(item));
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);

            _logger.LogInformation(
                "{ActionName}: Focus+Enter succeeded for logical menu item {Name}",
                actionName,
                SafeElementName(item));

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: Focus+Enter failed for logical menu item {Name}",
                actionName,
                SafeElementName(item));
        }

        try
        {
            return TryInstantPhysicalClick(item, actionName);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: physical fallback failed for logical menu item {Name}",
                actionName,
                SafeElementName(item));

            return false;
        }
    }

    private void OpenMenuItem(AutomationElement item)
    {
        BringElementWindowToForeground(item);

        try
        {
            if (item.Patterns.ExpandCollapse.IsSupported)
            {
                item.Patterns.ExpandCollapse.Pattern.Expand();
                return;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenMenuItem Expand failed for {Name}", SafeElementName(item));
        }

        try
        {
            item.Click();
            return;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenMenuItem Click failed for {Name}", SafeElementName(item));
        }

        TryInstantPhysicalClick(item, "OpenMenuItem");
    }

    private static bool IsLogicalMenuMode(UiRequest req) =>
        string.Equals(req.Locator?.Mode, "logical", StringComparison.OrdinalIgnoreCase);

    private void ActivateMenuItem(AutomationElement item)
    {
        BringElementWindowToForeground(item);

        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                return;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ActivateMenuItem Invoke failed for {Name}", SafeElementName(item));
        }

        try
        {
            item.Click();
            return;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ActivateMenuItem Click failed for {Name}", SafeElementName(item));
        }

        try
        {
            item.Focus();
            Keyboard.Press(VirtualKeyShort.RETURN);
            return;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ActivateMenuItem Enter failed for {Name}", SafeElementName(item));
        }

        if (!TryInstantPhysicalClick(item, "ActivateMenuItem"))
        {
            throw new InvalidOperationException(
                $"Failed to activate menu item '{SafeElementName(item)}'");
        }
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
            {
                var mainWindow = session.Application.GetMainWindow(session.Automation);
                var mainHandle = mainWindow != null ? SafeWindowHandle(mainWindow) : (IntPtr?)null;
                PerformHandleAlert(session.Automation, buttonNames, closeOnly,
                    session.Application.ProcessId, mainHandle);
            }
            else
            {
                using var tempAutomation = new UIA3Automation();
                PerformHandleAlert(tempAutomation, buttonNames, closeOnly, null, null);
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
    private static void PerformHandleAlert(
        AutomationBase automation,
        string[] buttonNames,
        bool closeOnly,
        int? processId,
        IntPtr? mainWindowHandle)
    {
        var dialog = FindPopupDialog(automation, processId, mainWindowHandle);
        if (dialog == null) return;

        if (!closeOnly && buttonNames.Length > 0)
        {
            var cf = automation.ConditionFactory;

            // Look for a button whose name matches one of the candidate names.
            var btn = dialog
                .FindAllDescendants(cf.ByControlType(ControlType.Button))
                .FirstOrDefault(b => buttonNames.Any(n =>
                    string.Equals(b.Name, n, StringComparison.OrdinalIgnoreCase)));

            if (btn != null)
            {
                if (btn.Patterns.Invoke.IsSupported)
                {
                    try
                    {
                        btn.Patterns.Invoke.Pattern.Invoke();
                        return;
                    }
                    catch (Exception ex) when (ex is FlaUI.Core.Exceptions.ElementNotAvailableException
                                            || ex is COMException)
                    {
                        // Element became unavailable or invoke failed; fall through to mouse click.
                    }
                }

                btn.Click();
                return;
            }
        }

        // Close the dialog (fallback or alertclose).
        CloseElement(dialog);
    }

    /// <summary>
    /// Locates the popup/confirmation dialog using a prioritised search strategy:
    /// <list type="number">
    ///   <item>
    ///     <description>
    ///       Collect all top-level desktop windows belonging to <paramref name="processId"/>
    ///       (when provided), falling back to UIA-tree descendants for owned dialogs that
    ///       are nested rather than top-level in the OS window hierarchy.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       Require the Window UIA pattern (<c>WindowPattern</c>) to confirm the candidate
    ///       is a real dialog window rather than an arbitrary UIA element.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       Prefer <c>IsModal == true</c> — the strongest signal that the window is a
    ///       confirmation popup.  Also recognises Win32 standard dialog class <c>#32770</c>
    ///       which is functionally modal even when <c>IsModal</c> is reported as false.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       Exclude the main application window by comparing
    ///       <paramref name="mainWindowHandle"/> against each candidate's
    ///       <c>NativeWindowHandle</c> to avoid acting on the wrong window.
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <description>
    ///       Fall back to the OS foreground window — confirmation popups typically steal
    ///       focus when they open.
    ///     </description>
    ///   </item>
    /// </list>
    /// </summary>
    private static AutomationElement? FindPopupDialog(
        AutomationBase automation,
        int? processId,
        IntPtr? mainWindowHandle)
    {
        var cf = automation.ConditionFactory;
        var desktop = automation.GetDesktop();

        // Filters Window-type elements and appends qualifying ones to `candidates`,
        // deduplicating by NativeWindowHandle to avoid processing the same window twice.
        var seenHandles = new HashSet<IntPtr>();
        var candidates  = new List<AutomationElement>();

        void CollectWindowCandidates(AutomationElement[] windows)
        {
            foreach (var w in windows)
            {
                // Priority 2: must support WindowPattern — confirms it is a real window/dialog.
                try { if (!w.Patterns.Window.IsSupported) continue; }
                catch { continue; }

                // Priority 1: same ProcessId when known — avoids picking up unrelated dialogs.
                if (processId.HasValue)
                {
                    try { if (w.Properties.ProcessId.Value != processId.Value) continue; }
                    catch { continue; }
                }

                // Priority 4: exclude the main application window by NativeWindowHandle.
                if (mainWindowHandle.HasValue && mainWindowHandle.Value != IntPtr.Zero)
                {
                    try
                    {
                        if (w.Properties.NativeWindowHandle.Value == mainWindowHandle.Value)
                            continue;
                    }
                    catch { /* can't read handle — keep as candidate */ }
                }

                // Require a valid (non-zero) HWND and deduplicate by it.
                // Elements without a readable HWND are not real top-level windows
                // and are skipped to avoid false positives.
                IntPtr hwnd;
                try { hwnd = w.Properties.NativeWindowHandle.Value; }
                catch { continue; }

                if (hwnd != IntPtr.Zero && seenHandles.Add(hwnd))
                    candidates.Add(w);
            }
        }

        // Priority 1 — top-level windows first (direct desktop children).
        CollectWindowCandidates(desktop.FindAllChildren(cf.ByControlType(ControlType.Window)));

        // If no qualifying top-level window was found, descend into the UIA tree: owned
        // dialogs are sometimes nested under their owning window rather than appearing as
        // direct desktop children.  This scan is more expensive so it is only run as a
        // fallback when the fast path yields nothing.
        if (candidates.Count == 0)
            CollectWindowCandidates(desktop.FindAllDescendants(cf.ByControlType(ControlType.Window)));

        if (candidates.Count == 0) return null;

        // Priority 3a: IsModal == true is the strongest confirmation-popup signal.
        var modal = candidates.FirstOrDefault(w =>
        {
            try { return w.Patterns.Window.Pattern.IsModal; }
            catch { return false; }
        });
        if (modal != null) return modal;

        // Priority 3b: Win32 standard dialog class (#32770) is functionally modal even
        // when IsModal is reported as false by UIA.
        var win32Dialog = candidates.FirstOrDefault(w =>
        {
            try { return string.Equals(w.ClassName, "#32770", StringComparison.Ordinal); }
            catch { return false; }
        });
        if (win32Dialog != null) return win32Dialog;

        // Priority 5: foreground window fallback — confirmation popups usually take focus.
        var foregroundHwnd = GetForegroundWindow();
        if (foregroundHwnd != IntPtr.Zero)
        {
            var foreground = candidates.FirstOrDefault(w =>
            {
                try { return w.Properties.NativeWindowHandle.Value == foregroundHwnd; }
                catch { return false; }
            });
            if (foreground != null) return foreground;
        }

        // Last resort: return the first remaining candidate.
        return candidates.FirstOrDefault();
    }

    /// <summary>
    /// Finds the popup / non-modal window whose title contains <c>req.Value</c>,
    /// brings it to the foreground, sends a single Enter key-press to dismiss it
    /// (accepting the default button), and then waits up to <see cref="DefaultRetry"/>
    /// for the window to disappear.
    ///
    /// Use this operation for popup dialogs where <c>alertok</c> cannot dismiss the
    /// window because it is not detected as a modal dialog (e.g. <c>IsModal</c> is false).
    /// The <c>value</c> field must contain a partial window title (case-insensitive).
    /// </summary>
    private object? PopUpOk(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' must be a partial window title for 'popupok'.");

        // Resolve the automation instance to use.
        var session = _ctx.ActiveSession;
        AutomationBase automation;
        UIA3Automation? tempAutomation = null;

        if (session != null)
            automation = session.Automation;
        else
        {
            tempAutomation = new UIA3Automation();
            automation = tempAutomation;
        }

        try
        {
            var cf = automation.ConditionFactory;
            var desktop = automation.GetDesktop();

            // Find the target popup window.
            var window = FindWindowByTitle(desktop, cf, req.Value);
            if (window == null)
                throw new InvalidOperationException(
                    $"No window with title containing '{SanitizeValue(req.Value)}' was found.");

            // Bring the popup to the foreground and send Enter to dismiss it.
            window.SetForeground();
            Thread.Sleep(WindowActivationDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);

            // Poll until the window is gone or we time out.
            var deadline = DateTime.UtcNow + DefaultRetry;
            while (DateTime.UtcNow < deadline)
            {
                Thread.Sleep(RetryInterval);
                if (FindWindowByTitle(desktop, cf, req.Value) == null)
                    return new { success = true };
            }

            // Final check after the retry loop.
            if (FindWindowByTitle(desktop, cf, req.Value) == null)
                return new { success = true };

            throw new InvalidOperationException(
                $"Window '{SanitizeValue(req.Value)}' did not close after pressing Enter " +
                $"within {DefaultRetry.TotalSeconds}s.");
        }
        finally
        {
            tempAutomation?.Dispose();
        }
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
        {
            try
            {
                cell.Patterns.Invoke.Pattern.Invoke();
                return;
            }
            catch (Exception ex) when (ex is FlaUI.Core.Exceptions.ElementNotAvailableException
                                    || ex is COMException)
            {
                // Element became unavailable or invoke failed; fall through to mouse click.
            }
        }

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
    /// Also detects popup/dialog windows from other processes (e.g. system authentication
    /// dialogs, OS security prompts) that appear as a result of application actions.
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
                else
                {
                    // GetAllTopLevelWindows only scans direct children of the desktop UIA
                    // element filtered by the application's PID.  Owned dialog windows
                    // (ControlType=Window, LocalizedControlType="dialog") that are spawned
                    // by the application are often nested as descendants of the owning window
                    // in the UIA virtual tree, not as direct desktop children, so they are
                    // invisible to GetAllTopLevelWindows.
                    //
                    // Additionally, some popup windows run in a different process entirely
                    // (e.g. Windows credential dialogs, COM-hosted security prompts).
                    //
                    // Both cases are handled by scanning all Window-type descendants of the
                    // desktop, filtering to visible windows only, and letting ClaimFirstNewWindow
                    // identify any that have not been seen before.
                    //
                    // The scan is throttled to run at most once per DesktopScanThrottle interval
                    // (default 2 s) so that rapid successive operations (e.g. typing into a form)
                    // do not pay the cost of a full desktop traversal on every call.
                    if (DateTime.UtcNow - session.LastDesktopScan >= AutomationSession.DesktopScanThrottle)
                    {
                        session.LastDesktopScan = DateTime.UtcNow;
                        try
                        {
                            var cf = session.Automation.ConditionFactory;
                            var allDesktopDescendants = session.Automation.GetDesktop()
                                .FindAllDescendants(cf.ByControlType(ControlType.Window));
                            var newVisibleWindows = allDesktopDescendants
                                .Where(w =>
                                {
                                    try { return !w.IsOffscreen; }
                                    catch { return false; }
                                })
                                .ToArray();
                            var newPopup = session.ClaimFirstNewWindow(newVisibleWindows);
                            if (newPopup != null)
                            {
                                session.ActiveWindow = newPopup;
                                _logger.LogInformation(
                                    "Auto-followed popup/dialog window: '{Title}'",
                                    SanitizeValue(newPopup.Name));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogDebug(ex,
                                "Popup/dialog window check failed; continuing with current active window.");
                        }
                    }
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
                        // set by SwitchWindow or the cross-process popup detection above.
                        // Only clear if the element is truly inaccessible.
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

    private AutomationElement GetSearchRootForMenuOperation(UiRequest req, AutomationSession session)
    {
        if (req.Locator != null && !IsEmptyLocator(req.Locator))
        {
            var currentRoot = GetWindowRoot(session);

            _logger.LogInformation(
                "Menu locator provided. currentRoot name={Name}, automationId={AutomationId}, controlType={ControlType}, hwnd=0x{Hwnd:X}",
                SafeElementName(currentRoot),
                SafeElementAutomationId(currentRoot),
                currentRoot.ControlType,
                SafeWindowHandle(currentRoot).ToInt64());

            try
            {
                var locatorRoot = FindLocatorWithRetry(session, currentRoot, req.Locator);

                _logger.LogInformation(
                    "Menu operation using locator root under current root. name={Name}, automationId={AutomationId}, controlType={ControlType}, hwnd=0x{Hwnd:X}",
                    SafeElementName(locatorRoot),
                    SafeElementAutomationId(locatorRoot),
                    locatorRoot.ControlType,
                    SafeWindowHandle(locatorRoot).ToInt64());

                return locatorRoot;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(
                    ex,
                    "Menu locator was not found under current root. Will try app windows before desktop.");
            }

            foreach (var appWindow in session.Application.GetAllTopLevelWindows(session.Automation))
            {
                try
                {
                    var locatorRoot = FindLocatorWithRetry(session, appWindow, req.Locator);

                    _logger.LogInformation(
                        "Menu operation using locator root under app window. appWindow={AppWindow}, locatorRoot={Root}, automationId={AutomationId}, controlType={ControlType}",
                        SafeElementName(appWindow),
                        SafeElementName(locatorRoot),
                        SafeElementAutomationId(locatorRoot),
                        locatorRoot.ControlType);

                    return locatorRoot;
                }
                catch
                {
                    // Best effort: continue scanning other application windows.
                }
            }

            try
            {
                var locatorRoot = FindLocatorWithRetry(session, session.Automation.GetDesktop(), req.Locator);

                _logger.LogInformation(
                    "Menu operation using locator root under desktop fallback. locatorRoot={Root}, automationId={AutomationId}, controlType={ControlType}",
                    SafeElementName(locatorRoot),
                    SafeElementAutomationId(locatorRoot),
                    locatorRoot.ControlType);

                if (IsDesktopRoot(locatorRoot))
                {
                    throw new InvalidOperationException(
                        "menupath search root resolved to Desktop1. This is invalid. " +
                        "Pass a MenuBar locator or switch to the application window first.");
                }

                return locatorRoot;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Menu locator was provided but could not be found. Locator={DescribeLocator(req.Locator)}",
                    ex);
            }
        }

        var root = GetWindowRoot(session);

        _logger.LogInformation(
            "Menu operation using window root. name={Name}, automationId={AutomationId}, controlType={ControlType}, hwnd=0x{Hwnd:X}",
            SafeElementName(root),
            SafeElementAutomationId(root),
            root.ControlType,
            SafeWindowHandle(root).ToInt64());

        return root;
    }

    private static bool IsDesktopRoot(AutomationElement element) =>
        element.ControlType == ControlType.Pane &&
        (
            string.Equals(SafeElementName(element), "Desktop1", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(SafeElementName(element), "Desktop", StringComparison.OrdinalIgnoreCase)
        );

    private static object CreateElementSnapshot(AutomationElement element) => new
    {
        name = SafeElementName(element),
        automationId = SafeElementAutomationId(element),
        controlType = element.ControlType.ToString(),
        hwnd = SafeWindowHandle(element).ToInt64()
    };

    private static bool IsEmptyLocator(UiLocator? locator)
    {
        return locator == null || (
            string.IsNullOrWhiteSpace(locator.Name) &&
            string.IsNullOrWhiteSpace(locator.AutomationId) &&
            string.IsNullOrWhiteSpace(locator.ClassName) &&
            string.IsNullOrWhiteSpace(locator.XPath) &&
            string.IsNullOrWhiteSpace(locator.ControlType));
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
    /// Normalizes window titles for matching by converting non-breaking spaces to
    /// regular spaces and trimming surrounding whitespace.
    /// </summary>
    private static string NormalizeTitle(string? value) =>
        (value ?? string.Empty)
            .Replace("\u00A0", " ")
            .Trim();

    private static bool TitleContains(AutomationElement element, string titleFragment)
    {
        try
        {
            var name = NormalizeTitle(element.Name);
            var fragment = NormalizeTitle(titleFragment);
            return !string.IsNullOrWhiteSpace(name) &&
                   !string.IsNullOrWhiteSpace(fragment) &&
                   name.Contains(fragment, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static List<object> SnapshotWindowTitles(AutomationBase automation)
    {
        var results = new List<object>();

        try
        {
            var cf = automation.ConditionFactory;
            var windows = automation.GetDesktop()
                .FindAllDescendants(cf.ByControlType(ControlType.Window));

            foreach (var window in windows)
            {
                try
                {
                    results.Add(new
                    {
                        title = window.Name,
                        automationId = window.AutomationId,
                        className = window.ClassName,
                        processId = window.Properties.ProcessId.ValueOrDefault,
                        hwnd = SafeWindowHandle(window).ToInt64(),
                        isOffscreen = window.IsOffscreen
                    });
                }
                catch
                {
                    // Ignore unstable window snapshots.
                }
            }
        }
        catch
        {
            // Ignore snapshot failures.
        }

        return results;
    }

    private void LogWindowSearchFailure(AutomationBase automation, string titleFragment)
    {
        var snapshot = SnapshotWindowTitles(automation);
        _logger.LogWarning(
            "switchwindow failed. Requested title fragment='{TitleFragment}'. Visible window snapshot={Snapshot}",
            SanitizeValue(titleFragment),
            snapshot);
    }

    /// <summary>
    /// Finds an element using a locator with up to 5 s retry (500 ms interval).
    /// <para>
    /// The window root is re-evaluated on every retry iteration so that newly
    /// opened dialogs (e.g. a modal confirmation dialog that appears after
    /// clicking OK) are picked up by the auto-follow logic inside
    /// <see cref="GetWindowRoot"/> and subsequent retries search the correct
    /// window rather than the stale previous root.
    /// </para>
    /// </summary>
    private AutomationElement FindWithRetry(UiRequest req)
    {
        var locator = RequireLocator(req);
        var session = RequireSession();

        var deadline = DateTime.UtcNow + DefaultRetry;
        while (true)
        {
            // Re-query the root on every iteration: if a new dialog opened since
            // the last attempt, GetWindowRoot will auto-follow it and return the
            // dialog as the new root, allowing the element search to succeed.
            var root = GetWindowRoot(session);
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

    private static bool IsElementEditable(AutomationElement element)
    {
        if (!element.IsEnabled)
            return false;

        var valuePattern = element.Patterns.Value.PatternOrDefault;
        if (valuePattern != null)
            return !valuePattern.IsReadOnly;

        return element.ControlType == ControlType.Edit || element.ControlType == ControlType.Document;
    }

    private bool TryPhysicalClick(AutomationElement element, string actionName) =>
        TryInstantPhysicalClick(element, actionName);

    private bool TryElementClick(AutomationElement element, string actionName)
    {
        try
        {
            BringElementWindowToForeground(element);
            element.Click();
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "{ActionName}: element.Click failed", actionName);
            return false;
        }
    }

    private bool TryInvokePattern(AutomationElement element, string actionName)
    {
        try
        {
            if (!element.Patterns.Invoke.IsSupported)
                return false;

            element.Patterns.Invoke.Pattern.Invoke();
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "{ActionName}: InvokePattern failed", actionName);
            return false;
        }
    }

    private bool TryInstantPhysicalClick(AutomationElement element, string actionName)
    {
        try
        {
            BringElementWindowToForeground(element);

            var rect = element.BoundingRectangle;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            {
                _logger.LogWarning(
                    "{ActionName}: invalid bounding rectangle for name={Name}, controlType={ControlType}",
                    actionName,
                    SafeElementName(element),
                    element.ControlType);
                return false;
            }

            var point = new Point(
                (int)Math.Round(rect.Left + (rect.Width / 2.0)),
                (int)Math.Round(rect.Top + (rect.Height / 2.0)));

            _logger.LogInformation(
                "{ActionName}: instant physical click at {Point}, name={Name}, automationId={AutomationId}, controlType={ControlType}",
                actionName,
                point,
                SafeElementName(element),
                SafeElementAutomationId(element),
                element.ControlType);

            return SendInstantLeftClick(point, actionName);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "{ActionName}: instant physical click failed", actionName);
            return false;
        }
    }

    private bool SendInstantLeftClick(Point point, string actionName)
    {
        try
        {
            if (!SetCursorPos((int)point.X, (int)point.Y))
            {
                _logger.LogWarning(
                    "{ActionName}: SetCursorPos failed. LastError={Error}",
                    actionName,
                    Marshal.GetLastWin32Error());
                return false;
            }

            Thread.Sleep(CursorPositionStabilityDelayMs);

            var inputs = new[]
            {
                new INPUT
                {
                    type = INPUT_MOUSE,
                    U = new InputUnion
                    {
                        mi = new MOUSEINPUT
                        {
                            dx = 0,
                            dy = 0,
                            mouseData = 0,
                            dwFlags = MOUSEEVENTF_LEFTDOWN,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_MOUSE,
                    U = new InputUnion
                    {
                        mi = new MOUSEINPUT
                        {
                            dx = 0,
                            dy = 0,
                            mouseData = 0,
                            dwFlags = MOUSEEVENTF_LEFTUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                }
            };

            var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
            if (sent != inputs.Length)
            {
                _logger.LogWarning(
                    "{ActionName}: SendInput failed. sent={Sent}, expected={Expected}, LastError={Error}",
                    actionName,
                    sent,
                    inputs.Length,
                    Marshal.GetLastWin32Error());
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "{ActionName}: SendInstantLeftClick failed", actionName);
            return false;
        }
    }

    private void BringElementWindowToForeground(AutomationElement? element)
    {
        if (element == null)
            return;

        bool activated = false;
        try
        {
            var hwnd = element.Properties.NativeWindowHandle.Value;
            if (hwnd != IntPtr.Zero)
            {
                var root = GetAncestor(hwnd, GA_ROOT);
                if (root != IntPtr.Zero)
                    activated = SetForegroundWindow(root);
            }
        }
        catch
        {
            // Best effort only.
        }

        if (!activated)
        {
            try
            {
                var pid = element.Properties.ProcessId.Value;
                var proc = System.Diagnostics.Process.GetProcessById(pid);
                var mainHwnd = proc.MainWindowHandle;
                if (mainHwnd != IntPtr.Zero)
                    SetForegroundWindow(mainHwnd);
            }
            catch
            {
                // Best effort only.
            }
        }

        Thread.Sleep(WindowActivationDelayMs);
    }

    private static string SafeElementName(AutomationElement element)
    {
        try
        {
            return SanitizeValue(element.Name);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeElementAutomationId(AutomationElement element)
    {
        try
        {
            return SanitizeValue(element.AutomationId);
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Strips control characters from a user-supplied string before it is
    /// written to a log message to prevent log-injection attacks.
    /// </summary>
    private static string SanitizeValue(string? value) =>
        System.Text.RegularExpressions.Regex.Replace(
            value ?? string.Empty, @"[\r\n\t]", "_");

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private const uint INPUT_MOUSE = 0;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint GA_ROOT = 2;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetForegroundWindow();
}
