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
    private const string ClearSelectAllBackspaceKeys = "^a{BACKSPACE}";
    private const string ClearSelectAllDeleteKeys = "^a{DELETE}";
    private const int MenuExpandDelayMs = 250;
    private const int MenuActionDelayMs = 150;
    private const int MenuFocusDelayMs = 75;
    private const int KeyboardInputReadyDelayMs = 100;
    private const int ComboBoxSelectionCommitDelayMs = 150;
    // Menu parent chains in supported desktop apps are shallow; 20 gives ample room for deeply
    // nested menus while preventing unbounded traversal of unstable UIA ancestors. If exceeded,
    // strict full-path matching fails safely instead of selecting a wrong duplicate leaf.
    private const int MaxMenuParentChainDepth = 20;
    // Submenu popups generally appear near the parent item's right edge; these offsets click the
    // arrow area without hugging the border and allow small overlap/jitter in native menu popups.
    private const double SubmenuArrowMinOffsetPx = 8.0;
    private const double SubmenuArrowMaxOffsetPx = 20.0;
    private const double SubmenuArrowWidthDivisor = 8.0;
    private const int SubmenuHorizontalProximityPx = 20;
    private const int SubmenuVerticalProximityPx = 40;
    private const int DropdownItemPhysicalClickSettleMs = 250;
    private const int DropdownItemFallbackDelayMs = 150;
    private const int DropdownItemMinPadX = 6;
    private const int DropdownItemMaxPadX = 18;
    private const int DropdownItemPadXDivisor = 10;
    private const int DropdownItemMinPadY = 3;
    private const int DropdownItemMaxPadY = 8;
    private const int DropdownItemPadYDivisor = 4;
    private const int TreeItemExpanderMinOffsetPx = 6;
    private const int TreeItemExpanderMaxOffsetPx = 18;
    private const int TreeItemExpanderWidthDivisor = 10;
    private const int ComboBoxRightEdgeMinOffsetPx = 8;
    private const int ComboBoxRightEdgeMaxOffsetPx = 20;
    private const int ComboBoxRightEdgeOffsetDivisor = 8;
    private const int ComboBoxLeftEdgeMinOffsetPx = 8;
    private const int ComboBoxLeftEdgeMaxOffsetPx = 20;
    private const int ComboBoxLeftEdgeOffsetDivisor = 10;
    private const int ComboBoxDropdownVerticalTolerancePx = 30;
    private const int ComboBoxScrollSearchMaxAttempts = 30;
    private const int ComboBoxScrollPageWheelClicks = -3;
    private const int ComboBoxScrollSettleDelayMs = 150;
    private const int ComboBoxVisibleItemSearchLimit = 200;
    private const int ComboBoxPagedSearchMaxVisibleItems = 20;
    private const int ComboBoxPagedSearchMaxPages = 300;
    private const int ComboBoxPagedSearchSettleDelayMs = 150;
    private const int ComboBoxPagedSearchWheelClicks = -5;
    private const int ComboBoxAnchorWindowSearchMaxWindows = 300;
    private const int ComboBoxAnchorMoveDelayMs = 40;
    private const int ComboBoxAnchorReadDelayMs = 120;
    private const int ComboBoxKeyboardStepSearchMaxSteps = 2000;
    private const int ComboBoxKeyboardStepDelayMs = 40;
    private const int ComboBoxKeyboardStepReadDelayMs = 80;
    private const int ComboBoxKeyboardStepLogEvery = 25;
    // Detection limit and huge-list threshold are separate knobs even though they
    // currently share the same value: one caps sampling, the other classifies size.
    private const int ComboBoxLargeListDetectionLimit = 100;
    private const int ComboBoxHugeListTypeAheadThreshold = 100;
    private const int ComboBoxTypeAheadDelayMs = 150;
    private const int ComboBoxTypeAheadCommitDelayMs = 250;
    private const int ComboBoxTypeAheadFocusDelayMs = 150;
    private const int MaxComboBoxDropdownListCandidates = 100;
    private const int MaxWindowSearchDepth = 5;
    private const int MaxAssistiveDropdownItemsToDisplay = 25;
    private const int MaxDynamicDropdownCandidates = 80;
    private const int MaxDynamicPopupSearchDepth = 4;
    private const int MaxApplicationContextMenuCandidates = 80;
    private const int MaxApplicationContextMenuSearchDepth = 4;
    private const int MaxApplicationContextMenuPlaybackItems = 100;
    private const int DefaultListResponseLimit = 500;
    private const int MaxListResponseLimit = 5000;
    private const string DesktopRootName = "Desktop";
    private const string DesktopRootNameWithSuffix = "Desktop1";
    private const string InvalidMenuRootMessage =
        "menupath search root resolved to Desktop. This is invalid. " +
        "Pass a MenuBar locator or switch to the application window first.";
    private static readonly TimeSpan FindWindowCacheDuration = TimeSpan.FromSeconds(5);
    private static readonly object FindWindowCacheLock = new();
    private static readonly Dictionary<string, CachedWindowMatch> FindWindowCache = new();

    private sealed record CachedWindowMatch(AutomationElement Element, DateTime ExpiresAt);

    private enum DropdownItemClickRegion
    {
        LeftCenter,
        Center,
        RightCenter,
        UpperLeft,
        LowerLeft,
        ProbeAll
    }

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
            ["WIN"] = VirtualKeyShort.LWIN,
            ["LWIN"] = VirtualKeyShort.LWIN,
            ["RWIN"] = VirtualKeyShort.RWIN,
            ["PRINTSCREEN"] = VirtualKeyShort.SNAPSHOT,
            ["PRTSC"] = VirtualKeyShort.SNAPSHOT,
            ["CAPSLOCK"] = VirtualKeyShort.CAPITAL,
            ["NUMLOCK"] = VirtualKeyShort.NUMLOCK,
            ["SCROLLLOCK"] = VirtualKeyShort.SCROLL,
            ["PAUSE"] = VirtualKeyShort.PAUSE,
            ["APPS"] = VirtualKeyShort.APPS,
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
            "contextmenupath"  => ContextMenuPath(request),
            "inspectlogicalmenu" => InspectLogicalMenu(request),
            "inspectmenupathcandidates" => InspectMenuPathCandidates(request),
            "dumpmenus"        => DumpLogicalMenus(request),
            "dumplogicalmenus" => DumpLogicalMenus(request),
            "doubleclick"      => DoubleClick(request),
            "rightclick"       => RightClick(request),
            "hover"            => Hover(request),
            "focus"            => Focus(request),
            "type"             => TypeText(request),
            "typedate"         => TypeDate(request),
            "clear"            => Clear(request),
            "sendkeys"         => SendKeys(request),
            "expandtreeitem"   => ExpandTreeItem(request),
            "collapsetreeitem" => CollapseTreeItem(request),
            "selecttreeitem"   => SelectTreeItem(request),
            "expandtreepath"   => ExpandTreePath(request),
            "selecttreepath"   => SelectTreePath(request),
            "scroll"           => Scroll(request),
            "mousescroll"      => MouseScroll(request),
            "wheelscroll"      => MouseScroll(request),
            "check"            => Check(request),
            "uncheck"          => Uncheck(request),
            "select"           => Select(request),
            "selectaid"        => SelectByAid(request),
            "typeandselect"    => TypeAndSelect(request),
            "clickgridcell"    => ClickGridCell(request),
            "doubleclickgridcell" => DoubleClickGridCell(request),
            "openheaderdropdown" => OpenHeaderDropdown(request),
            "selectheaderdropdownitem" => SelectHeaderDropdownItem(request),
            // Dynamic menu playback operations use a root MenuItem locator.
            // selectdynamicmenuitem is kept for one-level compatibility and delegates
            // to path traversal; selectdynamicmenupath accepts Root>Child>Leaf or Child>Leaf.
            "selectdynamicmenuitem" => SelectDynamicMenuItem(request),
            "selectdynamicmenupath" => SelectDynamicMenuPath(request),
            "selectcomboboxitem" => SelectComboBoxItem(request),
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
        var cacheKey = BuildFindWindowCacheKey("root", root, title);
        if (TryGetCachedWindow(cacheKey, title, out var cached))
            return cached;

        AutomationElement? Cache(AutomationElement? match)
        {
            if (match != null)
                StoreCachedWindow(cacheKey, match);
            return match;
        }

        var windowCondition = cf.ByControlType(ControlType.Window);

        // Fast path: top-level windows are direct children of the Desktop.
        var topLevel = root.FindAllChildren(windowCondition)
            .FirstOrDefault(w => TitleContains(w, title));
        if (topLevel != null) return Cache(topLevel);

        // Fallback: owned/child windows may be nested beneath their parent in the UIA tree
        // (e.g. a dialog opened from a child window).
        return Cache(root.FindAllDescendants(windowCondition)
            .FirstOrDefault(w => TitleContains(w, title)));
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
        var cacheKey = BuildFindWindowCacheKey(session, titleFragment);
        if (TryGetCachedWindow(cacheKey, titleFragment, out var cached))
            return cached;

        AutomationElement? Cache(AutomationElement? match)
        {
            if (match != null)
                StoreCachedWindow(cacheKey, match);
            return match;
        }

        var cf = session.Automation.ConditionFactory;
        var windowCondition = cf.ByControlType(ControlType.Window);

        if (session.ActiveWindow != null && TitleContains(session.ActiveWindow, titleFragment))
            return Cache(session.ActiveWindow);

        if (session.ActiveWindow != null)
        {
            try
            {
                var activeDescendant = session.ActiveWindow
                    .FindAllDescendants(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
                if (activeDescendant != null)
                    return Cache(activeDescendant);
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
                    return Cache(mainWindow);

                var mainDescendant = mainWindow
                    .FindAllDescendants(windowCondition)
                    .FirstOrDefault(w => TitleContains(w, titleFragment));
                if (mainDescendant != null)
                    return Cache(mainDescendant);
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
                return Cache(topLevel);
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
                    return Cache(nested);
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
                return Cache(desktopChild);
        }
        catch
        {
            // Continue to broader search scopes.
        }

        try
        {
            return Cache(desktop == null
                ? null
                : desktop
                .FindAllDescendants(windowCondition)
                .FirstOrDefault(w => TitleContains(w, titleFragment)));
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
            title = SafeElementName(asWindow),
            automationId = SafeElementAutomationId(asWindow),
            controlType = SafeElementControlType(asWindow),
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
        var limit = GetListResponseLimit(req);

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

        var total = elements.Length;
        var limited = elements.Take(limit).Select(e => new
        {
            name = SafeElementName(e),
            automationId = SafeElementAutomationId(e),
            className = SafeElementClassName(e),
            controlType = SafeElementControlType(e),
            enabled = SafeIsEnabled(e),
            visible = SafeIsOffscreen(e) is false
        }).ToList();

        return new
        {
            elements = limited,
            total,
            returned = limited.Count,
            limit,
            truncated = total > limit
        };
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

        if (req.IncludeDesktopDescendants == true)
        {
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
        }

        IEnumerable<AutomationElement> filtered = windows;
        if (!string.IsNullOrWhiteSpace(req.Value))
            filtered = windows.Where(w => TitleContains(w, req.Value));

        return filtered.Select(w => new
        {
            title = SafeElementName(w),
            automationId = SafeElementAutomationId(w),
            className = SafeElementClassName(w),
            processId = SafeProcessId(w),
            hwnd = SafeWindowHandle(w).ToInt64(),
            isOffscreen = SafeIsOffscreen(w)
        }).ToList();
    }

    private static int GetListResponseLimit(UiRequest req)
    {
        if (req.Limit is not { } requested || requested <= 0)
            return DefaultListResponseLimit;

        return Math.Min(requested, MaxListResponseLimit);
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
                throw new InvalidOperationException(
                    $"Menu path failed. Menu item '{name}' was not found in the current menu scope. Path='{SanitizeValue(normalizedValue)}'");
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
                searchRoot = GetDropdownForMenuItem(session, item);
            }
            else
            {
                ActivateMenuItem(item);
                Thread.Sleep(MenuActionDelayMs);
            }
        }

        return null;
    }

    private AutomationElement GetDropdownForMenuItem(
        AutomationSession session,
        AutomationElement menuItem)
    {
        var dropdown = FindDynamicMenuDropdown(session, menuItem)
            ?? FindDynamicSubMenuDropdown(session, menuItem);

        if (dropdown == null)
        {
            throw new InvalidOperationException(
                $"Dropdown was not found after opening menu item '{SafeElementName(menuItem)}'.");
        }

        return dropdown;
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
            throw new InvalidOperationException(InvalidMenuRootMessage);

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
                    controlType = SafeElementControlType(e),
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
        var limit = GetListResponseLimit(req);

        var allMenuItems = root
            .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
            .ToList();

        var menuItems = allMenuItems
            .Take(limit)
            .Select(e =>
            {
                try
                {
                    return new
                    {
                        name = SafeElementName(e),
                        automationId = SafeElementAutomationId(e),
                        controlType = SafeElementControlType(e),
                        className = SafeElementClassName(e),
                        bounds = SafeBoundingRectangle(e),
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

        var total = allMenuItems.Count;
        return new
        {
            root = new
            {
                name = SafeElementName(root),
                automationId = SafeElementAutomationId(root),
                controlType = SafeElementControlType(root),
                hwnd = SafeWindowHandle(root).ToInt64()
            },
            menuItems,
            total,
            returned = menuItems.Count,
            limit,
            truncated = total > limit
        };
    }

    private object? InspectMenuPathCandidates(UiRequest req)
    {
        var session = RequireSession();
        var searchRoot = GetSearchRootForMenuOperation(req, session);
        var cf = session.Automation.ConditionFactory;
        var parts = SplitMenuPath(req.Value ?? string.Empty);

        if (parts.Count == 0)
            throw new ArgumentException("value is required, example: DQA>Level 17");

        var allMenuItems = searchRoot
            .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
            .ToList();

        var topLevel = searchRoot.ControlType == ControlType.MenuBar
            ? FindDirectChildrenByControlType(searchRoot, ControlType.MenuItem)
            : allMenuItems;

        var parentCandidates = topLevel
            .Where(e => MenuTextMatches(e, parts[0]))
            .ToList();

        return new
        {
            root = new
            {
                name = SafeElementName(searchRoot),
                automationId = SafeElementAutomationId(searchRoot),
                controlType = SafeElementControlType(searchRoot),
                totalMenuItems = allMenuItems.Count,
                topLevelCount = topLevel.Count
            },
            path = parts,
            topLevelMenuItems = topLevel.Select(x => new
            {
                name = SafeElementName(x),
                automationId = SafeElementAutomationId(x),
                childCount = SafeMenuItemChildCount(x, cf)
            }).ToList(),
            parentCandidates = parentCandidates.Select(x => new
            {
                name = SafeElementName(x),
                automationId = SafeElementAutomationId(x),
                bounds = SafeBoundingRectangle(x),
                children = GetMenuItemDescendantSummary(x, cf)
            }).ToList()
        };
    }

    private object? DoubleClick(UiRequest req)
    {
        FindWithRetry(req).DoubleClick();
        return null;
    }

    private object? RightClick(UiRequest req)
    {
        var element = FindWithRetry(req);

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        try
        {
            element.RightClick();

            return new
            {
                rightClicked = true,
                strategy = "FlaUI.RightClick",
                element = CreateElementSnapshot(element)
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "FlaUI RightClick failed for {Name}; falling back to physical right click",
                SafeElementName(element));
        }

        if (TryInstantPhysicalRightClick(element, "RightClick"))
        {
            return new
            {
                rightClicked = true,
                strategy = "physical-right-click",
                element = CreateElementSnapshot(element)
            };
        }

        throw new InvalidOperationException($"Failed to right-click element '{SafeElementName(element)}'.");
    }

    private object? ContextMenuPath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for contextmenupath.");

        var element = FindWithRetry(req);
        var rawValue = System.Net.WebUtility.HtmlDecode(req.Value).Trim();
        var pathParts = rawValue
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (pathParts.Count == 0)
            throw new ArgumentException("contextmenupath requires at least one menu item in value.");

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        if (!TryInstantPhysicalRightClick(element, "ContextMenuPath RightClick"))
            throw new InvalidOperationException($"Failed to right-click element '{SafeElementName(element)}'.");

        Thread.Sleep(MenuExpandDelayMs);

        var session = RequireSession();
        var currentRoot = FindActiveContextMenuPopup(session)
            ?? throw new InvalidOperationException("No application context menu was detected after right-click.");

        for (var i = 0; i < pathParts.Count; i++)
        {
            var part = pathParts[i];
            var item = FindContextMenuItemByText(session, currentRoot, part);

            if (item == null)
            {
                var available = GetContextMenuItems(session, currentRoot, MaxAssistiveDropdownItemsToDisplay)
                    .Select(SafeElementName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                throw new InvalidOperationException(
                    $"Context menu item '{part}' was not found for path '{rawValue}'. Available: {string.Join(", ", available)}");
            }

            var isLast = i == pathParts.Count - 1;

            if (isLast)
            {
                if (!ActivateContextMenuItem(item, part))
                    throw new InvalidOperationException($"Failed to activate context menu item '{part}'.");

                return new
                {
                    selected = rawValue,
                    strategy = "context-menu-path",
                    element = CreateElementSnapshot(element)
                };
            }

            if (!OpenContextSubMenu(item, part))
                throw new InvalidOperationException($"Failed to open context submenu '{part}' for path '{rawValue}'.");

            Thread.Sleep(MenuExpandDelayMs);

            currentRoot = FindContextSubMenuPopup(session, item) ?? FindActiveContextMenuPopup(session);

            if (currentRoot == null)
                throw new InvalidOperationException($"Context submenu popup was not found after opening '{part}'.");
        }

        throw new InvalidOperationException($"Context menu path '{rawValue}' was not completed.");
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
        if (WinFormsDateTimePickerHelper.IsDateTimePicker(element))
            return TypeDate(element, req.Value);

        if (!FocusElementForKeyboardInput(element, "TypeText"))
        {
            throw new InvalidOperationException(
                $"Keyboard focus could not be confirmed on target before typing. target='{SafeElementName(element)}'");
        }

        Thread.Sleep(KeyboardInputReadyDelayMs);

        Keyboard.Type(req.Value);
        return null;
    }

    private object? TypeDate(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("Date value is required.");

        var element = FindWithRetry(req);

        if (!WinFormsDateTimePickerHelper.IsDateTimePicker(element))
        {
            _logger.LogWarning(
                "typedate called on non-DateTimePicker. name={Name}, className={ClassName}, controlType={ControlType}",
                SafeElementName(element),
                SafeElementClassName(element),
                element.ControlType);
        }

        return TypeDate(element, req.Value);
    }

    private object? TypeDate(AutomationElement element, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Date value is required.");

        if (!WinFormsDateTimePickerHelper.TryParseDateParts(value, out var month, out var day, out var year))
            throw new ArgumentException("Invalid date value. Use MM/DD/YYYY or MM-DD-YYYY.");

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        ClickDatePickerMonthSection(element);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerClickDelayMs);

        Keyboard.Press(VirtualKeyShort.HOME);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(month);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Press(VirtualKeyShort.RIGHT);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(day);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Press(VirtualKeyShort.RIGHT);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(year);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Press(VirtualKeyShort.RETURN);

        return new
        {
            typed = true,
            strategy = "date-segments",
            value = $"{month}/{day}/{year}",
            element = new
            {
                name = SafeElementName(element),
                className = SafeElementClassName(element),
                controlType = element.ControlType.ToString()
            }
        };
    }

    private object? Clear(UiRequest req)
    {
        var element = FindWithRetry(req);

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        try
        {
            if (element.Patterns.Value.IsSupported)
            {
                var pattern = element.Patterns.Value.Pattern;
                if (!pattern.IsReadOnly)
                {
                    pattern.SetValue(string.Empty);
                    return new
                    {
                        cleared = true,
                        strategy = "ValuePattern.SetValue",
                        element = CreateElementSnapshot(element)
                    };
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Clear via ValuePattern failed for {Name}", SafeElementName(element));
        }

        try
        {
            element.AsTextBox().Text = string.Empty;
            return new
            {
                cleared = true,
                strategy = "TextBox.Text",
                element = CreateElementSnapshot(element)
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Clear via AsTextBox failed for {Name}", SafeElementName(element));
        }

        try
        {
            element.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            SendKeysString(ClearSelectAllBackspaceKeys);
            return new
            {
                cleared = true,
                strategy = "CtrlA_Backspace",
                element = CreateElementSnapshot(element)
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Clear via Ctrl+A Backspace failed for {Name}", SafeElementName(element));
        }

        try
        {
            element.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            SendKeysString(ClearSelectAllDeleteKeys);
            return new
            {
                cleared = true,
                strategy = "CtrlA_Delete",
                element = CreateElementSnapshot(element)
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Clear via Ctrl+A Delete failed for {Name}", SafeElementName(element));
        }

        throw new InvalidOperationException($"Failed to clear element '{SafeElementName(element)}'.");
    }

    private object? SendKeys(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("Parameter 'value' is required for 'sendkeys' operation.");

        AutomationElement? element = null;

        if (req.Locator != null && !IsEmptyLocator(req.Locator))
        {
            element = FindWithRetry(req);
            if (!FocusElementForKeyboardInput(element, "SendKeys"))
            {
                throw new InvalidOperationException(
                    $"Keyboard focus could not be confirmed on target before sendkeys. target='{SafeElementName(element)}'");
            }

            Thread.Sleep(KeyboardInputReadyDelayMs);
        }
        else
        {
            BringActiveWindowToForeground();
            Thread.Sleep(WindowActivationDelayMs);
        }

        var normalizedKeys = NormalizeSendKeysValue(req.Value);
        SendKeysString(normalizedKeys);

        return new
        {
            sent = true,
            original = req.Value,
            normalized = normalizedKeys,
            target = element == null ? null : CreateElementSnapshot(element)
        };
    }

    private object? ExpandTreeItem(UiRequest req)
    {
        var item = FindWithRetry(req);

        if (item.ControlType != ControlType.TreeItem)
        {
            _logger.LogWarning(
                "expandtreeitem called on non-TreeItem. name={Name}, controlType={ControlType}",
                SafeElementName(item),
                item.ControlType);
        }

        if (!ExpandTreeItemElement(item))
            throw new InvalidOperationException($"Failed to expand TreeItem '{SafeElementName(item)}'.");

        return new
        {
            expanded = true,
            name = SafeElementName(item)
        };
    }

    private object? CollapseTreeItem(UiRequest req)
    {
        var item = FindWithRetry(req);

        if (item.ControlType != ControlType.TreeItem)
        {
            _logger.LogWarning(
                "collapsetreeitem called on non-TreeItem. name={Name}, controlType={ControlType}",
                SafeElementName(item),
                item.ControlType);
        }

        if (!CollapseTreeItemElement(item))
            throw new InvalidOperationException($"Failed to collapse TreeItem '{SafeElementName(item)}'.");

        return new
        {
            collapsed = true,
            name = SafeElementName(item)
        };
    }

    private object? SelectTreeItem(UiRequest req)
    {
        var item = FindWithRetry(req);

        if (item.ControlType != ControlType.TreeItem)
        {
            _logger.LogWarning(
                "selecttreeitem called on non-TreeItem. name={Name}, controlType={ControlType}",
                SafeElementName(item),
                item.ControlType);
        }

        if (!SelectTreeItemElement(item))
            throw new InvalidOperationException($"Failed to select TreeItem '{SafeElementName(item)}'.");

        return new
        {
            selected = true,
            name = SafeElementName(item)
        };
    }

    private object? SelectTreePath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for selecttreepath.");

        var root = FindWithRetry(req);

        var parts = req.Value
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        var current = root;

        foreach (var part in parts)
        {
            if (!ExpandTreeItemElement(current))
                throw new InvalidOperationException($"Failed to expand TreeItem '{SafeElementName(current)}'.");

            Thread.Sleep(MenuExpandDelayMs);

            var child = FindChildTreeItemByName(current, part);

            if (child == null)
            {
                throw new InvalidOperationException(
                    $"Tree path part '{part}' not found under '{SafeElementName(current)}'. Available (up to 25): {GetAvailableChildTreeItemNames(current)}");
            }

            current = child;
        }

        if (!SelectTreeItemElement(current))
            throw new InvalidOperationException($"Failed to select TreeItem '{SafeElementName(current)}'.");

        return new
        {
            selected = true,
            path = req.Value,
            final = SafeElementName(current)
        };
    }

    private object? ExpandTreePath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for expandtreepath.");

        var root = FindWithRetry(req);

        var parts = req.Value
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        var current = root;

        foreach (var part in parts)
        {
            if (!ExpandTreeItemElement(current))
                throw new InvalidOperationException($"Failed to expand TreeItem '{SafeElementName(current)}'.");

            Thread.Sleep(MenuExpandDelayMs);

            var child = FindChildTreeItemByName(current, part);

            if (child == null)
            {
                throw new InvalidOperationException(
                    $"Tree path part '{part}' not found under '{SafeElementName(current)}'. Available (up to 25): {GetAvailableChildTreeItemNames(current)}");
            }

            current = child;
        }

        if (!ExpandTreeItemElement(current))
            throw new InvalidOperationException($"Failed to expand final TreeItem '{SafeElementName(current)}'.");

        return new
        {
            expanded = true,
            path = req.Value,
            final = SafeElementName(current)
        };
    }

    private bool ExpandTreeItemElement(AutomationElement treeItem)
    {
        try
        {
            if (treeItem.Patterns.ExpandCollapse.IsSupported)
            {
                var pattern = treeItem.Patterns.ExpandCollapse.Pattern;

                if (pattern.ExpandCollapseState != ExpandCollapseState.Expanded)
                {
                    pattern.Expand();
                    Thread.Sleep(MenuExpandDelayMs);
                }

                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ExpandCollapsePattern.Expand failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        try
        {
            treeItem.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RIGHT);
            Thread.Sleep(MenuExpandDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Keyboard RIGHT expand failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        try
        {
            var rect = treeItem.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round((double)(rect.Left + Math.Max(
                        TreeItemExpanderMinOffsetPx,
                        Math.Min(TreeItemExpanderMaxOffsetPx, rect.Width / TreeItemExpanderWidthDivisor)))),
                    (int)Math.Round((double)(rect.Top + rect.Height / 2)));

                if (SendInstantLeftClick(point, $"Expand TreeItem {SafeElementName(treeItem)}"))
                {
                    Thread.Sleep(MenuExpandDelayMs);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Physical expander click failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        return false;
    }

    private bool CollapseTreeItemElement(AutomationElement treeItem)
    {
        try
        {
            if (treeItem.Patterns.ExpandCollapse.IsSupported)
            {
                var pattern = treeItem.Patterns.ExpandCollapse.Pattern;

                if (pattern.ExpandCollapseState != ExpandCollapseState.Collapsed)
                {
                    pattern.Collapse();
                    Thread.Sleep(MenuExpandDelayMs);
                }

                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ExpandCollapsePattern.Collapse failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        try
        {
            treeItem.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.LEFT);
            Thread.Sleep(MenuExpandDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Keyboard LEFT collapse failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        return false;
    }

    private bool SelectTreeItemElement(AutomationElement treeItem)
    {
        try
        {
            if (treeItem.Patterns.SelectionItem.IsSupported)
            {
                treeItem.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SelectionItem.Select failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        try
        {
            var rect = treeItem.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round((double)(rect.Left + rect.Width / 2)),
                    (int)Math.Round((double)(rect.Top + rect.Height / 2)));

                if (SendInstantLeftClick(point, $"Select TreeItem {SafeElementName(treeItem)}"))
                {
                    Thread.Sleep(MenuActionDelayMs);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Physical select failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        try
        {
            treeItem.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(MenuActionDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Focus+Enter select failed for TreeItem {Name}", SafeElementName(treeItem));
        }

        return false;
    }

    private List<AutomationElement> GetChildTreeItems(
        AutomationElement parentTreeItem,
        int maxItems = 200)
    {
        var results = new List<AutomationElement>();

        try
        {
            var session = RequireSession();
            var cf = session.Automation.ConditionFactory;

            foreach (var child in parentTreeItem.FindAllChildren(cf.ByControlType(ControlType.TreeItem)))
            {
                results.Add(child);

                if (results.Count >= maxItems)
                    return results;
            }

            return results;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetChildTreeItems failed for {Name}", SafeElementName(parentTreeItem));
            return results;
        }
    }

    private AutomationElement? FindChildTreeItemByName(
        AutomationElement parentTreeItem,
        string childName)
    {
        var children = GetChildTreeItems(parentTreeItem);

        return children.FirstOrDefault(x =>
            string.Equals(
                NormalizeMenuText(SafeElementName(x)),
                NormalizeMenuText(childName),
                StringComparison.OrdinalIgnoreCase) ||
            string.Equals(
                NormalizeMenuText(SafeElementAutomationId(x)),
                NormalizeMenuText(childName),
                StringComparison.OrdinalIgnoreCase));
    }

    private string GetAvailableChildTreeItemNames(AutomationElement parentTreeItem, int maxItems = 25)
    {
        var available = GetChildTreeItems(parentTreeItem, maxItems)
            .Select(SafeElementName)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        return string.Join(", ", available);
    }

    private object? Scroll(UiRequest req)
    {
        var element = FindWithRetry(req);
        element.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();
        return null;
    }

    /// <summary>
    /// Scrolls the mouse wheel over an element (or over the current cursor position when
    /// no locator is provided).
    ///
    /// <list type="bullet">
    ///   <item><b>locator</b> – (optional) The element to scroll over.  When provided the
    ///     cursor is moved to the element's centre before scrolling.</item>
    ///   <item><b>value</b> – (optional) Scroll amount as a number of wheel clicks.
    ///     Positive values scroll up; negative values scroll down.
    ///     The strings <c>"up"</c> and <c>"down"</c> also work and map to ±3 clicks.
    ///     Omitting <c>value</c> defaults to 3 clicks down.</item>
    /// </list>
    /// </summary>
    private object? MouseScroll(UiRequest req)
    {
        AutomationElement? element = null;

        if (req.Locator != null && !IsEmptyLocator(req.Locator))
        {
            element = FindWithRetry(req);
            BringElementWindowToForeground(element);
            Thread.Sleep(WindowActivationDelayMs);

            var rect = element.BoundingRectangle;
            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round(rect.Left + rect.Width / 2.0),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                SetCursorPos(point.X, point.Y);
                Thread.Sleep(CursorPositionStabilityDelayMs);
            }
        }

        var wheelClicks = ParseWheelClicks(req.Value);
        if (!SendMouseWheel(wheelClicks))
            throw new InvalidOperationException($"Mouse wheel scroll failed. wheelClicks={wheelClicks}");

        return new
        {
            scrolled = true,
            wheelClicks,
            direction = wheelClicks >= 0 ? "up" : "down",
            target = element == null ? null : CreateElementSnapshot(element)
        };
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
        if (element.ControlType == ControlType.ComboBox && !string.IsNullOrWhiteSpace(req.Value))
        {
            var comboRequest = new UiRequest
            {
                Operation = "selectcomboboxitem",
                Locator = req.Locator,
                Value = req.Value
            };

            return SelectComboBoxItem(comboRequest);
        }

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

        Thread.Sleep(ComboBoxSelectionCommitDelayMs);

        try
        {
            element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Select: Collapse failed for {Name}", SafeElementName(element));
        }

        WaitForComboBoxDropdownToClose(element, timeoutMs: 1000);

        // Important: do not leave pending focus on the dropdown list item.
        // Focus back to the ComboBox itself so the next operation starts from a stable state.
        try
        {
            element.Focus();
            Thread.Sleep(MenuFocusDelayMs);
        }
        catch
        {
            // best effort
        }

        return new
        {
            selected = req.Value,
            index = req.Index,
            comboBox = CreateElementSnapshot(element),
            focusStabilized = true
        };
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

            var topLevelItems = root.ControlType == ControlType.MenuBar
                ? FindDirectChildrenByControlType(root, ControlType.MenuItem)
                : allMenuItems;

            _logger.LogInformation(
                "FindLogicalMenuPath: path={Path}, rootName={RootName}, rootControlType={RootControlType}, totalMenuItems={Total}, topLevelCount={TopLevelCount}",
                string.Join(" > ", parts),
                SafeElementName(root),
                root.ControlType,
                allMenuItems.Count,
                topLevelItems.Count);

            if (_logger.IsEnabled(LogLevel.Debug))
            {
                _logger.LogDebug(
                    "FindLogicalMenuPath top-level details: path={Path}, topLevelItems={TopLevelItems}",
                    string.Join(" > ", parts),
                    topLevelItems.Select(x => new
                    {
                        name = SafeElementName(x),
                        automationId = SafeElementAutomationId(x),
                        nameNorm = NormalizeMenuText(SafeElementName(x)),
                        expectedNorm = NormalizeMenuText(parts[0]),
                        isMatch = MenuTextMatches(x, parts[0]),
                        childCount = SafeMenuItemChildCount(x, cf)
                    }).ToList());
            }

            var parentCandidates = topLevelItems
                .Where(e => MenuTextMatches(e, parts[0]))
                .ToList();

            _logger.LogInformation(
                "FindLogicalMenuPath: parent candidate count for '{Parent}' = {CandidateCount}",
                parts[0],
                parentCandidates.Count);

            if (_logger.IsEnabled(LogLevel.Debug))
            {
                _logger.LogDebug(
                    "FindLogicalMenuPath parent candidate details for '{Parent}' = {Candidates}",
                    parts[0],
                    parentCandidates.Select(x => new
                    {
                        name = SafeElementName(x),
                        automationId = SafeElementAutomationId(x),
                        nameNorm = NormalizeMenuText(SafeElementName(x)),
                        expectedNorm = NormalizeMenuText(parts[0]),
                        isMatch = MenuTextMatches(x, parts[0]),
                        childCount = SafeMenuItemChildCount(x, cf)
                    }).ToList());
            }

            foreach (var candidate in parentCandidates)
            {
                var resolved = TryResolveMenuPathFromParent(candidate, cf, parts, 1);

                if (resolved != null)
                {
                    _logger.LogInformation(
                        "FindLogicalMenuPath: resolved path={Path} using parent={Parent}, target={Target}",
                        string.Join(" > ", parts),
                        SafeElementName(candidate),
                        SafeElementName(resolved));

                    return resolved;
                }
            }

            var chainMatch = FindMenuItemByFullPathChain(root, cf, parts);
            if (chainMatch != null)
                return chainMatch;

            _logger.LogWarning(
                "FindLogicalMenuPath strict resolution failed. No fallback leaf selection will be used because duplicate leaf names can select wrong menu. path={Path}",
                string.Join(" > ", parts));

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

    private AutomationElement? FindMenuItemByFullPathChain(
        AutomationElement root,
        ConditionFactory cf,
        IReadOnlyList<string> parts)
    {
        var leafName = parts[^1];

        var leafCandidates = root
            .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
            .Where(e => MenuTextMatches(e, leafName))
            .ToList();

        foreach (var leaf in leafCandidates)
        {
            var chain = BuildParentChainNames(leaf);

            if (DoesMenuParentChainEndWithPath(chain, parts))
            {
                _logger.LogInformation(
                    "Resolved menu path by full parent chain. path={Path}, chain={Chain}",
                    string.Join(">", parts),
                    string.Join(">", chain));

                return leaf;
            }
        }

        _logger.LogWarning(
            "No leaf matched full parent chain. path={Path}, leafCandidateCount={Count}",
            string.Join(">", parts),
            leafCandidates.Count);

        return null;
    }

    private List<string> BuildParentChainNames(AutomationElement element)
    {
        var result = new List<string>();

        try
        {
            var current = element;

            while (current != null && result.Count < MaxMenuParentChainDepth)
            {
                var name = SafeElementName(current);
                var ct = current.ControlType;

                if (ct == ControlType.MenuItem && !string.IsNullOrWhiteSpace(name))
                    result.Add(NormalizeMenuText(name));

                current = current.Parent;
            }
        }
        catch
        {
            // ignore unstable UIA parent chains
        }

        result.Reverse();
        return result;
    }

    internal static bool DoesMenuParentChainEndWithPath(
        IReadOnlyList<string> chain,
        IReadOnlyList<string> pathParts)
    {
        var path = pathParts
            .Select(NormalizeMenuText)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        if (chain.Count < path.Count)
            return false;

        var offset = chain.Count - path.Count;

        for (var i = 0; i < path.Count; i++)
        {
            if (!string.Equals(chain[offset + i], path[i], StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
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
            var directChildren = FindDirectChildrenByControlType(current, ControlType.MenuItem);

            _logger.LogInformation(
                "Looking for direct child. parent={Parent}, lookingFor={LookingFor}, childCount={ChildCount}",
                SafeElementName(current),
                nextName,
                directChildren.Count);

            if (_logger.IsEnabled(LogLevel.Debug))
            {
                _logger.LogDebug(
                    "Looking for direct child details. parent={Parent}, lookingFor={LookingFor}, children={Children}",
                    SafeElementName(current),
                    nextName,
                    directChildren.Select(x => new
                    {
                        name = SafeElementName(x),
                        automationId = SafeElementAutomationId(x),
                        nameNorm = NormalizeMenuText(SafeElementName(x)),
                        expectedNorm = NormalizeMenuText(nextName),
                        isMatch = MenuTextMatches(x, nextName)
                    }).ToList());
            }

            var next = directChildren.FirstOrDefault(e => MenuTextMatches(e, nextName));

            if (next == null)
            {
                var descendants = current
                    .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                    .ToList();

                _logger.LogInformation(
                    "Looking for descendant child. parent={Parent}, lookingFor={LookingFor}, descendantCount={DescendantCount}",
                    SafeElementName(current),
                    nextName,
                    descendants.Count);

                if (_logger.IsEnabled(LogLevel.Debug))
                {
                    _logger.LogDebug(
                        "Looking for descendant child details. parent={Parent}, lookingFor={LookingFor}, descendants={Descendants}",
                        SafeElementName(current),
                        nextName,
                        descendants.Select(x => new
                        {
                            name = SafeElementName(x),
                            automationId = SafeElementAutomationId(x),
                            nameNorm = NormalizeMenuText(SafeElementName(x)),
                            expectedNorm = NormalizeMenuText(nextName),
                            isMatch = MenuTextMatches(x, nextName)
                        }).ToList());
                }

                next = descendants.FirstOrDefault(e => MenuTextMatches(e, nextName));
            }

            if (next == null)
            {
                _logger.LogWarning(
                    "Menu path step not found. currentParent={Parent}, lookingFor={LookingFor}",
                    SafeElementName(current),
                    nextName);

                return null;
            }

            current = next;
        }

        return current;
    }

    private List<AutomationElement> FindDirectChildrenByControlType(
        AutomationElement parent,
        ControlType controlType)
    {
        try
        {
            return parent
                .FindAllChildren()
                .Where(e =>
                {
                    try { return e.ControlType == controlType; }
                    catch { return false; }
                })
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "FindDirectChildrenByControlType failed for {Parent} and {ControlType}",
                SafeElementName(parent),
                controlType);
            return [];
        }
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

                chain.Add($"{SafeElementControlType(current)}:{name}:{automationId}");
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
        var expectedNorm = NormalizeMenuText(expected);

        if (string.IsNullOrWhiteSpace(expectedNorm))
            return false;

        var nameNorm = NormalizeMenuText(SafeElementName(element));
        var automationIdNorm = NormalizeMenuText(SafeElementAutomationId(element));

        return string.Equals(nameNorm, expectedNorm, StringComparison.OrdinalIgnoreCase) ||
               string.Equals(automationIdNorm, expectedNorm, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeMenuText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var normalized = value
            .Replace("&amp;", string.Empty)
            .Replace("_", string.Empty)
            .Replace("\u00A0", " ")
            .Trim();

        while (normalized.Contains("  "))
            normalized = normalized.Replace("  ", " ");

        return normalized;
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

    private object? OpenHeaderDropdown(UiRequest req)
    {
        var header = FindWithRetry(req);
        var region = GridHeaderDropdownHelper.ParseRegion(req.Value ?? req.ClickRegion);

        if (!GridHeaderDropdownHelper.IsGridHeaderElement(header))
        {
            _logger.LogWarning(
                "openheaderdropdown called on non-header. name={Name}, controlType={ControlType}, className={ClassName}",
                SafeElementName(header),
                header.ControlType,
                SafeElementClassName(header));
        }

        var list = OpenHeaderDropdownAndFindList(header, region);

        if (list == null)
        {
            return new
            {
                opened = true,
                listFound = false,
                header = SafeElementName(header),
                region = region.ToString(),
                clickRegion = region.ToString()
            };
        }

        var items = GetListItems(list)
            .Select(x => SafeElementName(x))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        return new
        {
            opened = true,
            listFound = true,
            header = SafeElementName(header),
            region = region.ToString(),
            clickRegion = region.ToString(),
            items
        };
    }

    private object? SelectHeaderDropdownItem(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for selectheaderdropdownitem.");

        var header = FindWithRetry(req);

        if (!GridHeaderDropdownHelper.IsGridHeaderElement(header))
        {
            _logger.LogWarning(
                "selectheaderdropdownitem called on non-header. name={Name}, controlType={ControlType}, className={ClassName}",
                SafeElementName(header),
                header.ControlType,
                SafeElementClassName(header));
        }

        var region = GridHeaderDropdownHelper.ParseRegion(req.ClickRegion);
        var itemRegion = ParseDropdownItemRegion(req.ItemRegion);
        var list = OpenHeaderDropdownAndFindList(header, region);
        if (list == null)
            throw new InvalidOperationException("Header dropdown list was not found after opening.");

        var item = GetListItems(list)
            .FirstOrDefault(x =>
                string.Equals(
                    NormalizeMenuText(SafeElementName(x)),
                    NormalizeMenuText(req.Value),
                    StringComparison.OrdinalIgnoreCase));

        if (item == null)
        {
            var available = GetListItems(list)
                .Select(x => SafeElementName(x))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            throw new InvalidOperationException(
                $"Dropdown item '{req.Value}' was not found. Available: {string.Join(", ", available)}");
        }

        if (!ActivateDropdownListItem(item, req.Value, itemRegion))
            throw new InvalidOperationException($"Failed to click dropdown item '{req.Value}' with region {itemRegion}.");

        return new
        {
            selected = req.Value,
            header = SafeElementName(header),
            headerRegion = region.ToString(),
            itemRegion = itemRegion.ToString(),
            headerClickRegion = region.ToString(),
            itemClickRegion = itemRegion.ToString()
        };
    }

    private object? SelectDynamicMenuItem(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for selectdynamicmenuitem.");

        return SelectDynamicMenuPath(req);
    }

    private object? SelectDynamicMenuPath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for selectdynamicmenupath.");

        var parentMenuItem = FindWithRetry(req);
        if (parentMenuItem.ControlType != ControlType.MenuItem)
        {
            _logger.LogWarning(
                "selectdynamicmenupath called on non-MenuItem. name={Name}, controlType={ControlType}, className={ClassName}",
                SafeElementName(parentMenuItem),
                parentMenuItem.ControlType,
                SafeElementClassName(parentMenuItem));
        }

        var rawValue = System.Net.WebUtility.HtmlDecode(req.Value).Trim();
        var pathParts = rawValue
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (pathParts.Count == 0)
            throw new ArgumentException("selectdynamicmenupath requires at least one child menu item in value.");

        if (DynamicPathStartsWithParent(pathParts, parentMenuItem))
            pathParts.RemoveAt(0);

        var session = RequireSession();

        BringElementWindowToForeground(parentMenuItem);
        Thread.Sleep(WindowActivationDelayMs);

        if (!OpenDynamicMenuParent(parentMenuItem))
            throw new InvalidOperationException($"Failed to open dynamic menu '{SafeElementName(parentMenuItem)}'.");

        Thread.Sleep(MenuExpandDelayMs);

        var dropdown = FindDynamicMenuDropdown(session, parentMenuItem);
        if (dropdown == null)
        {
            throw new InvalidOperationException(
                $"Dynamic menu dropdown was not found after opening '{SafeElementName(parentMenuItem)}'.");
        }

        AutomationElement? item = null;

        for (var i = 0; i < pathParts.Count; i++)
        {
            var part = pathParts[i];
            var isLast = i == pathParts.Count - 1;

            item = FindDynamicMenuItemByName(session, dropdown, part);
            if (item == null)
            {
                var available = GetDynamicDropdownMenuItems(session, dropdown, MaxAssistiveDropdownItemsToDisplay)
                    .Select(SafeElementName)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                throw new InvalidOperationException(
                    $"Dynamic menu path item '{part}' was not found. Available: {string.Join(", ", available)}");
            }

            if (isLast)
            {
                if (!ActivateDynamicMenuItem(item, part))
                    throw new InvalidOperationException($"Failed to activate dynamic menu item '{part}'.");

                return new
                {
                    selected = string.Join(">", pathParts),
                    parent = SafeElementName(parentMenuItem),
                    dropdown = SafeElementName(dropdown)
                };
            }

            if (!OpenDynamicSubMenuItem(item))
                throw new InvalidOperationException($"Failed to open dynamic submenu '{part}'.");

            Thread.Sleep(MenuExpandDelayMs);

            dropdown = FindDynamicSubMenuDropdown(session, item)
                ?? throw new InvalidOperationException($"Dynamic submenu dropdown was not found after opening '{part}'.");
        }

        return null;
    }

    private object? SelectComboBoxItem(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for selectcomboboxitem.");

        var comboBox = FindWithRetry(req);
        if (!IsComboBoxElement(comboBox))
        {
            _logger.LogWarning(
                "selectcomboboxitem called on non-ComboBox. name={Name}, controlType={ControlType}, className={ClassName}",
                SafeElementName(comboBox),
                comboBox.ControlType,
                SafeElementClassName(comboBox));
        }

        var itemName = System.Net.WebUtility.HtmlDecode(req.Value).Trim();
        if (string.IsNullOrWhiteSpace(itemName))
            throw new ArgumentException("selectcomboboxitem requires a non-empty value.");

        var session = RequireSession();

        if (!OpenComboBoxDropdown(session, comboBox))
            throw new InvalidOperationException($"Failed to open ComboBox '{SafeElementName(comboBox)}'.");

        Thread.Sleep(MenuExpandDelayMs);

        if (IsHugeComboBoxDropdown(session, comboBox))
        {
            _logger.LogInformation(
                "ComboBox detected as huge list. Using paged visible-list search. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                itemName);

            if (TrySelectComboBoxByPagedVisibleSearch(session, comboBox, itemName))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "huge-list-paged-visible-search"
                };
            }

            _logger.LogInformation(
                "Huge ComboBox paged visible-list search failed. Trying visible anchor-window search. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                itemName);

            if (TrySelectComboBoxByVisibleAnchorWindowSearch(session, comboBox, itemName))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "huge-list-visible-anchor-window-search"
                };
            }

            if (req.AllowKeyboardFallback == true)
            {
                _logger.LogInformation(
                    "Huge ComboBox visible anchor-window search failed. allowKeyboardFallback=true, trying type-ahead. combo={Combo}, value={Value}",
                    SafeElementName(comboBox),
                    itemName);

                if (TrySelectComboBoxByKeyboardSafe(session, comboBox, itemName))
                {
                    var actual = GetComboBoxCurrentValue(session, comboBox);

                    return new
                    {
                        selected = itemName,
                        actual,
                        comboBox = SafeElementName(comboBox),
                        verified = true,
                        strategy = "huge-list-explicit-typeahead-fallback"
                    };
                }
            }

            var dropdownList = FindDynamicComboBoxList(session, comboBox);
            var visibleBatch = (dropdownList != null
                    ? GetListItemsBounded(session, dropdownList, ComboBoxPagedSearchMaxVisibleItems)
                    : GetLogicalComboBoxItems(session, comboBox, ComboBoxPagedSearchMaxVisibleItems))
                .Where(IsElementVisibleAndClickableEnough)
                .Take(ComboBoxPagedSearchMaxVisibleItems)
                .Select(SafeElementName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            throw new InvalidOperationException(
                $"Huge ComboBox item '{itemName}' was not found/verified by paged visible search or visible anchor-window search. " +
                $"dropdownListDetected={dropdownList != null}, " +
                $"expandedState={GetComboBoxExpandState(comboBox)}, " +
                $"currentValue='{GetComboBoxCurrentValue(session, comboBox)}', " +
                $"visibleBatch='{string.Join(", ", visibleBatch)}'");
        }

        var item = FindComboBoxItemByTextWithScroll(
            session,
            comboBox,
            itemName,
            ComboBoxScrollSearchMaxAttempts);

        if (item == null)
        {
            _logger.LogInformation(
                "ComboBox item '{Item}' was not found by visible/scroll search. Trying keyboard type-ahead fallback.",
                itemName);

            if (TrySelectComboBoxByKeyboardSafe(session, comboBox, itemName))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "keyboard-typeahead"
                };
            }

            var available = FindDynamicComboBoxItems(session, comboBox, maxItems: MaxAssistiveDropdownItemsToDisplay)
                .Select(SafeElementName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            throw new InvalidOperationException(
                $"ComboBox item '{itemName}' was not found after scrolling or keyboard type-ahead. " +
                $"dropdownListDetected={FindDynamicComboBoxList(session, comboBox) != null}, " +
                $"expandedState={GetComboBoxExpandState(comboBox)}, " +
                $"currentValue='{GetComboBoxCurrentValue(session, comboBox)}', " +
                $"availableFirst{MaxAssistiveDropdownItemsToDisplay}='{string.Join(", ", available)}'");
        }

        if (!ActivateComboBoxListItem(item, itemName))
            throw new InvalidOperationException($"Failed to activate ComboBox item '{itemName}'.");

        Thread.Sleep(300);

        if (!VerifyComboBoxSelectedValue(session, comboBox, itemName))
        {
            var actualAfterActivation = GetComboBoxCurrentValue(session, comboBox);

            _logger.LogWarning(
                "ComboBox ListItem activation did not select expected value. requested={Requested}, actual={Actual}. Trying keyboard type-ahead fallback.",
                itemName,
                actualAfterActivation);

            if (TrySelectComboBoxByKeyboardSafe(session, comboBox, itemName))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "listitem-click-keyboard-typeahead-fallback",
                    previousActual = actualAfterActivation
                };
            }

            var actualAfterFallback = GetComboBoxCurrentValue(session, comboBox);

            _logger.LogWarning(
                "ComboBox keyboard type-ahead fallback after ListItem activation did not verify. requested={Requested}, actual={Actual}, previousActual={PreviousActual}",
                itemName,
                actualAfterFallback,
                actualAfterActivation);

            throw new InvalidOperationException(
                $"ComboBox selection failed to verify after list item activation and keyboard type-ahead fallback. Requested='{itemName}', Actual='{actualAfterFallback}'");
        }

        try { comboBox.Patterns.ExpandCollapse.PatternOrDefault?.Collapse(); }
        catch { /* best effort */ }

        return new
        {
            selected = itemName,
            actual = GetComboBoxCurrentValue(session, comboBox),
            comboBox = SafeElementName(comboBox),
            verified = true
        };
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
                        "Auto-followed new window: '{Title}'", SafeElementName(newWindow));
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
                                    try { return SafeIsOffscreen(w) is false; }
                                    catch { return false; }
                                })
                                .ToArray();
                            var newPopup = session.ClaimFirstNewWindow(newVisibleWindows);
                            if (newPopup != null)
                            {
                                session.ActiveWindow = newPopup;
                                _logger.LogInformation(
                                    "Auto-followed popup/dialog window: '{Title}'",
                                    SafeElementName(newPopup));
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
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "Menu locator was not found under app window {AppWindow}; trying next window.",
                        SafeElementName(appWindow));
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
                    throw new InvalidOperationException(InvalidMenuRootMessage);

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
            string.Equals(SafeElementName(element), DesktopRootNameWithSuffix, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(SafeElementName(element), DesktopRootName, StringComparison.OrdinalIgnoreCase)
        );

    private static object CreateElementSnapshot(AutomationElement element) => new
    {
        name = SafeElementName(element),
        automationId = SafeElementAutomationId(element),
        controlType = SafeElementControlType(element),
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

    private static string BuildFindWindowCacheKey(AutomationSession session, string titleFragment) =>
        string.Join(
            "|",
            "session",
            session.SessionId,
            session.Application.ProcessId,
            NormalizeTitle(titleFragment).ToUpperInvariant());

    private static string BuildFindWindowCacheKey(string scope, AutomationElement root, string titleFragment) =>
        string.Join(
            "|",
            scope,
            SafeWindowHandle(root).ToInt64(),
            SafeProcessId(root)?.ToString() ?? "unknown",
            NormalizeTitle(titleFragment).ToUpperInvariant());

    private static bool TryGetCachedWindow(
        string cacheKey,
        string titleFragment,
        out AutomationElement? cached)
    {
        lock (FindWindowCacheLock)
        {
            if (FindWindowCache.TryGetValue(cacheKey, out var entry))
            {
                if (entry.ExpiresAt > DateTime.UtcNow && TitleContains(entry.Element, titleFragment))
                {
                    cached = entry.Element;
                    return true;
                }

                FindWindowCache.Remove(cacheKey);
            }
        }

        cached = null;
        return false;
    }

    private static void StoreCachedWindow(string cacheKey, AutomationElement element)
    {
        lock (FindWindowCacheLock)
            FindWindowCache[cacheKey] = new CachedWindowMatch(element, DateTime.UtcNow + FindWindowCacheDuration);
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
                        title = SafeElementName(window),
                        automationId = SafeElementAutomationId(window),
                        className = SafeElementClassName(window),
                        processId = SafeProcessId(window),
                        hwnd = SafeWindowHandle(window).ToInt64(),
                        isOffscreen = SafeIsOffscreen(window)
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
            var element = root.FindFirstDescendant(condition);
            if (element != null)
                return element;

            return TryFindWinFormsDateTimePickerByPartialClassName(root, session, locator);
        }
        catch
        {
            return null;
        }
    }

    private static AutomationElement? TryFindWinFormsDateTimePickerByPartialClassName(
        AutomationElement root,
        AutomationSession session,
        UiLocator locator)
    {
        if (!WinFormsDateTimePickerHelper.IsDateTimePickerClassName(locator.ClassName))
            return null;

        var candidates = !string.IsNullOrWhiteSpace(locator.ControlType)
            ? root.FindAllDescendants(session.Automation.ConditionFactory.ByControlType(ParseControlType(locator.ControlType)))
            : root.FindAllDescendants();

        foreach (var candidate in candidates)
        {
            if (!LocatorMatchesExceptClassName(candidate, locator))
                continue;

            var actualClassName = SafeElementClassName(candidate);
            if (actualClassName.Contains(locator.ClassName!, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }

        return null;
    }

    private static bool LocatorMatchesExceptClassName(AutomationElement element, UiLocator locator)
    {
        if (!string.IsNullOrWhiteSpace(locator.Name) &&
            !string.Equals(SafeElementName(element), locator.Name, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.AutomationId) &&
            !string.Equals(SafeElementAutomationId(element), locator.AutomationId, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ControlType) &&
            !string.Equals(element.ControlType.ToString(), locator.ControlType, StringComparison.OrdinalIgnoreCase))
            return false;

        return true;
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

    private static string NormalizeSendKeysValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return value;

        var v = value.Trim();
        var upper = v.ToUpperInvariant();

        // Canonical modifier+character outputs intentionally use lowercase letters
        // (for example "^a") because SendKeysString treats characters case-insensitively.
        return upper switch
        {
            "CTRL+A" or "CONTROL+A" => "^a",
            "CTRL+C" or "CONTROL+C" => "^c",
            "CTRL+V" or "CONTROL+V" => "^v",
            "CTRL+X" or "CONTROL+X" => "^x",
            "CTRL+S" or "CONTROL+S" => "^s",
            "CTRL+Z" or "CONTROL+Z" => "^z",
            "CTRL+Y" or "CONTROL+Y" => "^y",

            "ALT+F4" => "%{F4}",
            "ALT+TAB" => "%{TAB}",

            "SHIFT+TAB" => "+{TAB}",

            "ENTER" or "RETURN" => "{ENTER}",
            "BACKSPACE" or "BS" => "{BACKSPACE}",
            "DELETE" or "DEL" => "{DELETE}",
            "TAB" => "{TAB}",
            "ESC" or "ESCAPE" => "{ESC}",
            "SPACE" => "{SPACE}",
            "HOME" => "{HOME}",
            "END" => "{END}",
            "UP" or "ARROWUP" => "{UP}",
            "DOWN" or "ARROWDOWN" => "{DOWN}",
            "LEFT" or "ARROWLEFT" => "{LEFT}",
            "RIGHT" or "ARROWRIGHT" => "{RIGHT}",
            "PAGEUP" or "PGUP" => "{PAGEUP}",
            "PAGEDOWN" or "PGDN" => "{PAGEDOWN}",

            _ => value
        };
    }

    /// <summary>
    /// Parses wheel-click input.
    /// Defaults to -3 (scroll down three clicks) when value is omitted.
    /// </summary>
    private static int ParseWheelClicks(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return -3;

        var v = value.Trim();

        if (string.Equals(v, "down", StringComparison.OrdinalIgnoreCase))
            return -3;

        if (string.Equals(v, "up", StringComparison.OrdinalIgnoreCase))
            return 3;

        if (int.TryParse(v, out var clicks))
            return clicks;

        throw new ArgumentException(
            "Parameter 'value' must be 'up', 'down', or an integer. Positive values scroll up; negative values scroll down.");
    }

    /// <summary>
    /// Sends a key sequence string using AutoIt/keyboard-shorthand notation:
    /// <list type="bullet">
    ///   <item><c>{KEYNAME}</c> — named key (see <see cref="NamedKeys"/>)</item>
    ///   <item><c>^x</c> or <c>^{KEYNAME}</c> — Ctrl + key</item>
    ///   <item><c>+x</c> or <c>+{KEYNAME}</c> — Shift + key</item>
    ///   <item><c>%x</c> or <c>%{KEYNAME}</c> — Alt + key</item>
    ///   <item>any other character — typed literally</item>
    /// </list>
    /// Examples: <c>"^a"</c> = Ctrl+A, <c>"+{TAB}"</c> = Shift+Tab, <c>"%{F4}"</c> = Alt+F4.
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

            // ^x or ^{KEYNAME} — Ctrl+key
            if (c == '^' && i + 1 < keys.Length)
            {
                if (TryParseModifiedKey(keys, i + 1, out var vk, out int consumed))
                {
                    Keyboard.Press(VirtualKeyShort.LCONTROL);
                    Keyboard.Press(vk);
                    Keyboard.Release(vk);
                    Keyboard.Release(VirtualKeyShort.LCONTROL);
                    i += 1 + consumed;
                    continue;
                }
            }

            // +x or +{KEYNAME} — Shift+key
            if (c == '+' && i + 1 < keys.Length)
            {
                if (TryParseModifiedKey(keys, i + 1, out var vk, out int consumed))
                {
                    Keyboard.Press(VirtualKeyShort.LSHIFT);
                    Keyboard.Press(vk);
                    Keyboard.Release(vk);
                    Keyboard.Release(VirtualKeyShort.LSHIFT);
                    i += 1 + consumed;
                    continue;
                }
            }

            // %x or %{KEYNAME} — Alt+key
            if (c == '%' && i + 1 < keys.Length)
            {
                if (TryParseModifiedKey(keys, i + 1, out var vk, out int consumed))
                {
                    Keyboard.Press(VirtualKeyShort.ALT);
                    Keyboard.Press(vk);
                    Keyboard.Release(vk);
                    Keyboard.Release(VirtualKeyShort.ALT);
                    i += 1 + consumed;
                    continue;
                }
            }

            // Literal character
            Keyboard.Type(c.ToString());
            i++;
        }
    }

    /// <summary>
    /// Attempts to parse a key at position <paramref name="pos"/> in <paramref name="keys"/>.
    /// Handles both a single alphanumeric character (<c>a</c>–<c>z</c>, <c>0</c>–<c>9</c>)
    /// and a named-key token (<c>{KEYNAME}</c>).
    /// </summary>
    /// <param name="keys">The full key sequence string.</param>
    /// <param name="pos">Starting position to parse from.</param>
    /// <param name="vk">The resolved virtual key, if successful.</param>
    /// <param name="consumed">Number of characters consumed (not including the modifier prefix).</param>
    /// <returns><c>true</c> when a key was successfully resolved.</returns>
    private static bool TryParseModifiedKey(
        string keys, int pos, out VirtualKeyShort vk, out int consumed)
    {
        if (pos < keys.Length && keys[pos] == '{')
        {
            int end = keys.IndexOf('}', pos + 1);
            if (end > pos)
            {
                var keyName = keys[(pos + 1)..end];
                if (NamedKeys.TryGetValue(keyName, out vk))
                {
                    consumed = end - pos + 1; // length of "{KEYNAME}"
                    return true;
                }
            }
        }
        else if (pos < keys.Length)
        {
            var charVk = CharToVirtualKey(keys[pos]);
            if (charVk.HasValue)
            {
                vk = charVk.Value;
                consumed = 1;
                return true;
            }
        }

        vk = default;
        consumed = 0;
        return false;
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

    private bool IsElementEditable(AutomationElement element)
    {
        if (!element.IsEnabled)
            return false;

        if (TypeCapabilityHelper.IsTypeCapableElement(element, _ctx.ActiveSession?.Automation.ConditionFactory))
            return true;

        var valuePattern = element.Patterns.Value.PatternOrDefault;
        if (valuePattern != null)
            return !valuePattern.IsReadOnly;

        return element.ControlType == ControlType.Edit || element.ControlType == ControlType.Document;
    }

    private bool TryPhysicalClick(AutomationElement element, string actionName) =>
        TryInstantPhysicalClick(element, actionName);

    private bool FocusElementForKeyboardInput(
        AutomationElement element,
        string actionName,
        int timeoutMs = 1500)
    {
        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        try
        {
            element.Focus();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: element.Focus() failed for {Name}",
                actionName,
                SafeElementName(element));
        }

        if (WaitForKeyboardFocusOnElement(element, timeoutMs))
            return true;

        try
        {
            var rect = element.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round(rect.Left + rect.Width / 2.0),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                if (SendInstantLeftClick(point, $"{actionName}: focus physical click"))
                {
                    Thread.Sleep(MenuFocusDelayMs);

                    if (WaitForKeyboardFocusOnElement(element, timeoutMs))
                        return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "{ActionName}: physical focus click failed for {Name}",
                actionName,
                SafeElementName(element));
        }

        _logger.LogWarning(
            "{ActionName}: keyboard focus was not confirmed on target. target={Target}, controlType={ControlType}",
            actionName,
            SafeElementName(element),
            element.ControlType);

        return false;
    }

    private bool WaitForKeyboardFocusOnElement(
        AutomationElement target,
        int timeoutMs)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var focused = target.Automation.FocusedElement();

                if (focused != null && IsSameOrDescendantOf(focused, target))
                    return true;

                // Some WinForms/Edit controls expose focus on a child Edit/Text inside the target.
                if (focused != null && IsSameOrDescendantOf(target, focused))
                    return true;
            }
            catch
            {
                // ignore transient focus read failures
            }

            Thread.Sleep(50);
        }

        return false;
    }

    private bool IsSameOrDescendantOf(
        AutomationElement possibleChild,
        AutomationElement possibleAncestor)
    {
        try
        {
            if (AreSameElement(possibleChild, possibleAncestor))
                return true;

            var current = possibleChild.Parent;
            var depth = 0;

            while (current != null && depth < 20)
            {
                if (AreSameElement(current, possibleAncestor))
                    return true;

                current = current.Parent;
                depth++;
            }
        }
        catch
        {
            // ignore unstable parent chain
        }

        return false;
    }

    private bool AreSameElement(
        AutomationElement a,
        AutomationElement b)
    {
        try
        {
            var ah = SafeWindowHandle(a);
            var bh = SafeWindowHandle(b);

            if (ah != IntPtr.Zero && bh != IntPtr.Zero && ah == bh)
                return true;
        }
        catch
        {
            // ignore
        }

        try
        {
            var aidA = SafeElementAutomationId(a);
            var aidB = SafeElementAutomationId(b);
            var nameA = SafeElementName(a);
            var nameB = SafeElementName(b);

            if ((string.IsNullOrWhiteSpace(aidA) || string.IsNullOrWhiteSpace(aidB)) &&
                (string.IsNullOrWhiteSpace(nameA) || string.IsNullOrWhiteSpace(nameB)))
                return false;

            return string.Equals(aidA, aidB, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(nameA, nameB, StringComparison.OrdinalIgnoreCase) &&
                   a.ControlType == b.ControlType;
        }
        catch
        {
            return false;
        }
    }

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

    private bool ClickDatePickerMonthSection(AutomationElement element)
    {
        try
        {
            var rect = element.BoundingRectangle;

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return false;

            var point = WinFormsDateTimePickerHelper.GetMonthSectionPoint(rect);

            return SendInstantLeftClick(point, "Click Date Month Section");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Click Date Month Section failed");
            return false;
        }
    }

    private static bool IsComboBoxElement(AutomationElement element)
    {
        try
        {
            return element.ControlType == ControlType.ComboBox;
        }
        catch
        {
            return false;
        }
    }

    private AutomationElement? GetSearchRootForDynamicPopup(
        AutomationSession session,
        AutomationElement sourceElement)
    {
        try
        {
            var foregroundHwnd = GetForegroundWindow();

            if (foregroundHwnd != IntPtr.Zero)
            {
                var foreground = session.Automation.FromHandle(foregroundHwnd);
                if (foreground != null)
                    return foreground;
            }
        }
        catch
        {
            // continue
        }

        try
        {
            return FindWindowAncestorOrSelf(sourceElement);
        }
        catch
        {
            return null;
        }
    }

    private static AutomationElement? FindWindowAncestorOrSelf(AutomationElement element)
    {
        try
        {
            var current = element;

            for (int i = 0; i < MaxWindowSearchDepth && current != null; i++)
            {
                if (current.ControlType == ControlType.Window)
                    return current;

                current = current.Parent;
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private List<AutomationElement> GetLogicalComboBoxItems(
        AutomationSession session,
        AutomationElement comboBox,
        int maxItems)
    {
        try
        {
            return GetListItemsBounded(session, comboBox, maxItems);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetLogicalComboBoxItems failed for {Combo}", SafeElementName(comboBox));
            return [];
        }
    }

    private bool OpenComboBoxDropdown(AutomationSession session, AutomationElement comboBox)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (comboBox.Patterns.ExpandCollapse.IsSupported)
            {
                var pattern = comboBox.Patterns.ExpandCollapse.Pattern;
                if (pattern.ExpandCollapseState != ExpandCollapseState.Expanded)
                {
                    pattern.Expand();
                    Thread.Sleep(MenuExpandDelayMs);
                }

                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox ExpandCollapse failed for {Combo}", SafeElementName(comboBox));
        }

        try
        {
            var openButton = FindComboBoxOpenButton(session, comboBox);
            if (openButton != null)
            {
                var rect = openButton.BoundingRectangle;
                if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
                {
                    var point = new Point(
                        (int)Math.Round(rect.Left + (rect.Width / 2.0)),
                        (int)Math.Round(rect.Top + (rect.Height / 2.0)));

                    if (SendInstantLeftClick(point, $"Open ComboBox button {SafeElementName(comboBox)}"))
                    {
                        Thread.Sleep(MenuExpandDelayMs);
                        return true;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox open button click failed for {Combo}", SafeElementName(comboBox));
        }

        try
        {
            var rect = comboBox.BoundingRectangle;
            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round((double)rect.Right - Math.Max(
                        ComboBoxRightEdgeMinOffsetPx,
                        Math.Min(ComboBoxRightEdgeMaxOffsetPx, rect.Width / ComboBoxRightEdgeOffsetDivisor))),
                    (int)Math.Round(rect.Top + (rect.Height / 2.0)));

                if (SendInstantLeftClick(point, $"Open ComboBox right edge {SafeElementName(comboBox)}"))
                {
                    Thread.Sleep(MenuExpandDelayMs);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox right-edge click failed for {Combo}", SafeElementName(comboBox));
        }

        try
        {
            comboBox.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.F4);
            Thread.Sleep(MenuExpandDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox F4 open failed for {Combo}", SafeElementName(comboBox));
        }

        try
        {
            comboBox.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.ALT);
            Keyboard.Press(VirtualKeyShort.DOWN);
            Thread.Sleep(MenuExpandDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox Alt+Down open failed for {Combo}", SafeElementName(comboBox));
        }

        return false;
    }

    private bool WaitForComboBoxDropdownToClose(
        AutomationElement comboBox,
        int timeoutMs)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                if (comboBox.Patterns.ExpandCollapse.IsSupported)
                {
                    var state = comboBox.Patterns.ExpandCollapse.Pattern.ExpandCollapseState;

                    if (state == ExpandCollapseState.Collapsed ||
                        state == ExpandCollapseState.LeafNode)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                _logger.LogDebug(
                    "WaitForComboBoxDropdownToClose: treating state read failure as closed. combo={Combo}",
                    SafeElementName(comboBox));
                return true;
            }

            Thread.Sleep(50);
        }

        _logger.LogWarning(
            "ComboBox dropdown did not report Collapsed within timeout. combo={Combo}",
            SafeElementName(comboBox));

        return false;
    }

    private AutomationElement? FindComboBoxOpenButton(AutomationSession session, AutomationElement comboBox)
    {
        try
        {
            var cf = session.Automation.ConditionFactory;

            return comboBox
                .FindAllDescendants(cf.ByControlType(ControlType.Button))
                .FirstOrDefault(x =>
                {
                    var name = SafeElementName(x);
                    var aid = SafeElementAutomationId(x);

                    return (!string.IsNullOrWhiteSpace(name) && name.Contains("Open", StringComparison.OrdinalIgnoreCase))
                           || (!string.IsNullOrWhiteSpace(aid) && aid.Contains("Open", StringComparison.OrdinalIgnoreCase))
                           || string.IsNullOrWhiteSpace(name);
                });
        }
        catch
        {
            return null;
        }
    }

    private List<AutomationElement> FindDynamicComboBoxItems(
        AutomationSession session,
        AutomationElement comboBox,
        int maxItems = ComboBoxVisibleItemSearchLimit)
    {
        try
        {
            var list = FindDynamicComboBoxList(session, comboBox);

            if (list != null)
            {
                var popupItems = GetListItemsBounded(session, list, maxItems);

                if (popupItems.Count > 0)
                {
                    _logger.LogInformation(
                        "ComboBox dynamic items resolved from popup list. combo={Combo}, count={Count}",
                        SafeElementName(comboBox),
                        popupItems.Count);

                    return popupItems;
                }
            }

            var logicalItems = GetLogicalComboBoxItems(session, comboBox, maxItems);

            if (logicalItems.Count > 0)
            {
                _logger.LogInformation(
                    "ComboBox dynamic items resolved from logical children. combo={Combo}, count={Count}",
                    SafeElementName(comboBox),
                    logicalItems.Count);

                return logicalItems;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "FindDynamicComboBoxItems failed for combo={Combo}",
                SafeElementName(comboBox));
        }

        return [];
    }

    private bool TrySelectComboBoxByPagedVisibleSearch(
        AutomationSession session,
        AutomationElement comboBox,
        string itemName)
    {
        var requested = NormalizeMenuText(itemName);
        if (!ResetComboBoxDropdownToTop(session, comboBox))
            return false;

        for (var page = 0; page < ComboBoxPagedSearchMaxPages; page++)
        {
            var items = GetCurrentVisibleComboBoxBatch(session, comboBox, ComboBoxPagedSearchMaxVisibleItems);
            var signature = BuildComboBoxVisibleBatchSignature(items);

            foreach (var item in items)
            {
                var name = NormalizeMenuText(SafeElementName(item));

                if (string.Equals(name, requested, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogInformation(
                        "ComboBox paged visible-list search found current visible item. combo={Combo}, item={Item}, page={Page}, name={Name}",
                        SafeElementName(comboBox),
                        itemName,
                        page,
                        SafeElementName(item));

                    if (!ActivateComboBoxListItem(item, itemName))
                        return false;

                    Thread.Sleep(ComboBoxPagedSearchSettleDelayMs);
                    return VerifyComboBoxSelectedValue(session, comboBox, itemName);
                }
            }

            if (page == ComboBoxPagedSearchMaxPages - 1)
            {
                _logger.LogInformation(
                    "ComboBox paged visible-list search reached max pages. combo={Combo}, item={Item}, pages={Pages}, signature={Signature}",
                    SafeElementName(comboBox),
                    itemName,
                    ComboBoxPagedSearchMaxPages,
                    signature);

                break;
            }

            var scrolled = ScrollComboBoxVisibleWindowDown(session, comboBox);
            Thread.Sleep(ComboBoxPagedSearchSettleDelayMs);

            var afterItems = GetCurrentVisibleComboBoxBatch(session, comboBox, ComboBoxPagedSearchMaxVisibleItems);
            var afterSignature = BuildComboBoxVisibleBatchSignature(afterItems);
            var changed = !string.Equals(
                signature,
                afterSignature,
                StringComparison.OrdinalIgnoreCase);

            if (!scrolled || !changed)
            {
                _logger.LogInformation(
                    "ComboBox paged search stopped because visible window did not change. requested={Requested}, before={Before}, after={After}",
                    itemName,
                    signature,
                    afterSignature);

                break;
            }
        }

        return false;
    }

    private bool ResetComboBoxDropdownToTop(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!OpenComboBoxDropdown(session, comboBox))
            {
                _logger.LogWarning(
                    "ComboBox paged search could not open dropdown for reset. combo={Combo}",
                    SafeElementName(comboBox));

                return false;
            }

            Thread.Sleep(MenuExpandDelayMs);

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxPagedSearchResetTop"))
                return false;

            Thread.Sleep(ComboBoxPagedSearchSettleDelayMs);

            Keyboard.Press(VirtualKeyShort.HOME);
            Keyboard.Release(VirtualKeyShort.HOME);

            Thread.Sleep(ComboBoxPagedSearchSettleDelayMs);

            _logger.LogInformation(
                "ComboBox dropdown reset to top for paged search. combo={Combo}",
                SafeElementName(comboBox));

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox dropdown reset to top failed. combo={Combo}",
                SafeElementName(comboBox));

            return false;
        }
    }

    private List<AutomationElement> GetCurrentVisibleComboBoxBatch(
        AutomationSession session,
        AutomationElement comboBox,
        int batchSize)
    {
        try
        {
            var list = FindDynamicComboBoxList(session, comboBox);

            if (list != null)
            {
                return GetListItemsBounded(session, list, batchSize)
                    .Where(IsElementVisibleAndClickableEnough)
                    .Take(batchSize)
                    .ToList();
            }

            return GetLogicalComboBoxItems(session, comboBox, batchSize)
                .Where(IsElementVisibleAndClickableEnough)
                .Take(batchSize)
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox visible batch read failed. combo={Combo}",
                SafeElementName(comboBox));

            return [];
        }
    }

    private static bool IsElementVisibleAndClickableEnough(AutomationElement element)
    {
        try
        {
            var rect = element.BoundingRectangle;

            return !rect.IsEmpty &&
                   rect.Width > 0 &&
                   rect.Height > 0 &&
                   !element.IsOffscreen;
        }
        catch
        {
            return false;
        }
    }

    private bool ScrollComboBoxVisibleWindowDown(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!OpenComboBoxDropdown(session, comboBox))
            {
                _logger.LogWarning(
                    "ComboBox paged search could not reopen dropdown before visible-window scroll. combo={Combo}",
                    SafeElementName(comboBox));

                return false;
            }

            Thread.Sleep(MenuExpandDelayMs);

            var list = FindDynamicComboBoxList(session, comboBox);
            var target = list ?? comboBox;
            var rect = target.BoundingRectangle;

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            {
                _logger.LogWarning(
                    "ComboBox visible-window scroll skipped because bounds are invalid. combo={Combo}",
                    SafeElementName(comboBox));

                return false;
            }

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            SetCursorPos(point.X, point.Y);
            Thread.Sleep(CursorPositionStabilityDelayMs);

            var scrolled = SendMouseWheel(ComboBoxPagedSearchWheelClicks);
            Thread.Sleep(ComboBoxPagedSearchSettleDelayMs);

            if (scrolled)
            {
                _logger.LogInformation(
                    "ComboBox visible-window mouse-wheel scroll sent. combo={Combo}, point={Point}",
                    SafeElementName(comboBox),
                    point);
            }

            return scrolled;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox visible-window scroll failed. combo={Combo}",
                SafeElementName(comboBox));

            return false;
        }
    }

    private static string BuildComboBoxVisibleBatchSignature(
        IReadOnlyCollection<AutomationElement> items)
    {
        if (items.Count == 0)
            return string.Empty;

        return string.Join(
            "||",
            items.Select(item =>
            {
                var name = NormalizeMenuText(SafeElementName(item));
                var aid = NormalizeMenuText(SafeElementAutomationId(item));
                var rect = SafeBoundingRectangle(item);
                return $"{name}|{aid}|{rect}";
            }));
    }

    private AutomationElement? GetCurrentHighlightedComboBoxItem(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            var list = FindDynamicComboBoxList(session, comboBox);

            if (list != null)
            {
                var items = GetListItemsBounded(session, list, ComboBoxPagedSearchMaxVisibleItems);
                var selectedOrFocused = FindSelectedOrFocusedItem(items);

                if (selectedOrFocused != null)
                    return selectedOrFocused;

                return items.FirstOrDefault();
            }

            var logicalItems = GetLogicalComboBoxItems(session, comboBox, ComboBoxPagedSearchMaxVisibleItems);
            var logicalSelectedOrFocused = FindSelectedOrFocusedItem(logicalItems);

            if (logicalSelectedOrFocused != null)
                return logicalSelectedOrFocused;

            return logicalItems.FirstOrDefault();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox keyboard step search failed reading highlighted item. combo={Combo}",
                SafeElementName(comboBox));

            return null;
        }
    }

    private static AutomationElement? FindSelectedOrFocusedItem(
        IReadOnlyCollection<AutomationElement> items)
    {
        foreach (var item in items)
        {
            if (TryHasKeyboardFocus(item) || TryIsSelectionItemSelected(item))
                return item;
        }

        return null;
    }

    private static bool TryHasKeyboardFocus(AutomationElement item)
    {
        try
        {
            return item.Properties.HasKeyboardFocus.ValueOrDefault;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryIsSelectionItemSelected(AutomationElement item)
    {
        try
        {
            return item.Patterns.SelectionItem.IsSupported &&
                   item.Patterns.SelectionItem.Pattern.IsSelected;
        }
        catch
        {
            return false;
        }
    }

    private string GetCurrentHighlightedComboBoxItemText(
        AutomationSession session,
        AutomationElement comboBox)
    {
        var item = GetCurrentHighlightedComboBoxItem(session, comboBox);

        if (item == null)
            return string.Empty;

        var name = SafeElementName(item);

        if (!string.IsNullOrWhiteSpace(name))
            return name;

        return SafeElementAutomationId(item) ?? string.Empty;
    }

    private static bool ComboBoxTextMatchesExactly(
        string actual,
        string expected)
    {
        return string.Equals(
            NormalizeMenuText(actual),
            NormalizeMenuText(expected),
            StringComparison.OrdinalIgnoreCase);
    }

    private AutomationElement? FindComboBoxItemByTextWithScroll(
        AutomationSession session,
        AutomationElement comboBox,
        string itemName,
        int maxScrollAttempts)
    {
        var requested = NormalizeMenuText(itemName);
        var seenSignatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var attempt = 0; attempt <= maxScrollAttempts; attempt++)
        {
            var items = FindDynamicComboBoxItems(
                session,
                comboBox,
                maxItems: ComboBoxVisibleItemSearchLimit);

            foreach (var item in items)
            {
                var name = NormalizeMenuText(SafeElementName(item));
                var automationId = NormalizeMenuText(SafeElementAutomationId(item));

                if (string.Equals(name, requested, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(automationId, requested, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogInformation(
                        "ComboBox item found. item={Item}, attempt={Attempt}, name={Name}, automationId={AutomationId}",
                        itemName,
                        attempt,
                        SafeElementName(item),
                        SafeElementAutomationId(item));

                    return item;
                }
            }

            var signatureBeforeScroll = BuildComboBoxVisibleItemsSignature(items);

            if (!seenSignatures.Add(signatureBeforeScroll) && attempt > 0)
            {
                _logger.LogInformation(
                    "ComboBox visible item signature repeated before scrolling. Will try fallback scrolling before stopping. item={Item}, attempt={Attempt}, signature={Signature}",
                    itemName,
                    attempt,
                    signatureBeforeScroll);
            }

            if (attempt == maxScrollAttempts)
            {
                _logger.LogInformation(
                    "ComboBox scroll search reached max attempts. item={Item}, attempts={Attempts}",
                    itemName,
                    maxScrollAttempts);

                break;
            }

            var wheelChangedVisibleItems = TryScrollComboBoxAndDetectChange(
                session,
                comboBox,
                signatureBeforeScroll,
                useKeyboardPageDown: false,
                itemName,
                attempt);

            if (wheelChangedVisibleItems)
            {
                continue;
            }

            _logger.LogInformation(
                "ComboBox mouse-wheel scroll did not change visible items. Trying PageDown fallback. item={Item}, attempt={Attempt}",
                itemName,
                attempt);

            var pageDownChangedVisibleItems = TryScrollComboBoxAndDetectChange(
                session,
                comboBox,
                signatureBeforeScroll,
                useKeyboardPageDown: true,
                itemName,
                attempt);

            if (pageDownChangedVisibleItems)
            {
                continue;
            }

            _logger.LogInformation(
                "ComboBox scroll search stopped because neither mouse wheel nor PageDown changed visible items. item={Item}, attempt={Attempt}",
                itemName,
                attempt);

            break;
        }

        return null;
    }

    private bool IsHugeComboBoxDropdown(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            if (!OpenComboBoxDropdown(session, comboBox))
                return false;

            Thread.Sleep(MenuExpandDelayMs);

            var list = FindDynamicComboBoxList(session, comboBox);

            if (list == null)
                return false;

            var items = GetListItemsBounded(session, list, ComboBoxLargeListDetectionLimit);

            return items.Count >= ComboBoxHugeListTypeAheadThreshold;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to detect huge ComboBox dropdown. combo={Combo}", SafeElementName(comboBox));
            return false;
        }
    }

    private bool TryScrollComboBoxAndDetectChange(
        AutomationSession session,
        AutomationElement comboBox,
        string signatureBeforeScroll,
        bool useKeyboardPageDown,
        string itemName,
        int attempt)
    {
        try
        {
            var strategy = useKeyboardPageDown ? "PageDown" : "MouseWheel";

            var scrolled = useKeyboardPageDown
                ? ScrollComboBoxDropdownByKeyboard(session, comboBox)
                : ScrollComboBoxDropdown(session, comboBox, ComboBoxScrollPageWheelClicks);

            Thread.Sleep(ComboBoxScrollSettleDelayMs);

            if (!scrolled)
            {
                _logger.LogInformation(
                    "ComboBox scroll attempt returned false. strategy={Strategy}, item={Item}, attempt={Attempt}",
                    strategy,
                    itemName,
                    attempt);

                return false;
            }

            var itemsAfterScroll = FindDynamicComboBoxItems(
                session,
                comboBox,
                maxItems: ComboBoxVisibleItemSearchLimit);

            var signatureAfterScroll = BuildComboBoxVisibleItemsSignature(itemsAfterScroll);

            var changed = !string.Equals(
                signatureBeforeScroll,
                signatureAfterScroll,
                StringComparison.OrdinalIgnoreCase);

            _logger.LogInformation(
                "ComboBox scroll detect change. strategy={Strategy}, item={Item}, attempt={Attempt}, changed={Changed}, before={Before}, after={After}",
                strategy,
                itemName,
                attempt,
                changed,
                signatureBeforeScroll,
                signatureAfterScroll);

            return changed;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox scroll detect change failed. strategy={Strategy}, item={Item}, attempt={Attempt}, combo={Combo}",
                useKeyboardPageDown ? "PageDown" : "MouseWheel",
                itemName,
                attempt,
                SafeElementName(comboBox));

            return false;
        }
    }

    private static string BuildComboBoxVisibleItemsSignature(
        IReadOnlyCollection<AutomationElement> items)
    {
        if (items.Count == 0)
            return string.Empty;

        return string.Join(
            "||",
            items.Select(item =>
            {
                var name = NormalizeMenuText(SafeElementName(item));
                var aid = NormalizeMenuText(SafeElementAutomationId(item));
                var rect = SafeBoundingRectangle(item);
                return $"{name}|{aid}|{rect}";
            }));
    }

    private bool ScrollComboBoxDropdown(
        AutomationSession session,
        AutomationElement comboBox,
        int wheelClicks)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            var scrollTarget = FindDynamicComboBoxList(session, comboBox) ?? comboBox;
            var rect = scrollTarget.BoundingRectangle;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return false;

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            SetCursorPos(point.X, point.Y);
            Thread.Sleep(CursorPositionStabilityDelayMs);

            return SendMouseWheel(wheelClicks);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ScrollComboBoxDropdown failed for {Combo}", SafeElementName(comboBox));
            return false;
        }
    }

    private bool ScrollComboBoxDropdownByKeyboard(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxKeyboardScroll"))
            {
                _logger.LogWarning(
                    "ComboBox PageDown scroll skipped because focus could not be confirmed. combo={Combo}",
                    SafeElementName(comboBox));

                return false;
            }

            if (!OpenComboBoxDropdown(session, comboBox))
            {
                _logger.LogWarning(
                    "ComboBox PageDown scroll skipped because dropdown could not be opened. combo={Combo}",
                    SafeElementName(comboBox));

                return false;
            }

            Thread.Sleep(MenuExpandDelayMs);

            Keyboard.Press(VirtualKeyShort.NEXT);
            Keyboard.Release(VirtualKeyShort.NEXT);

            Thread.Sleep(ComboBoxScrollSettleDelayMs);

            _logger.LogInformation(
                "ComboBox dropdown PageDown scroll executed. combo={Combo}",
                SafeElementName(comboBox));

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox PageDown scroll fallback failed for combo={Combo}",
                SafeElementName(comboBox));

            return false;
        }
    }

    private bool TrySelectComboBoxByKeyboardSafe(
        AutomationSession session,
        AutomationElement comboBox,
        string itemName)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(itemName))
                return false;

            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxTypeAhead"))
            {
                _logger.LogWarning(
                    "ComboBox keyboard type-ahead skipped because focus could not be confirmed. combo={Combo}, item={Item}",
                    SafeElementName(comboBox),
                    itemName);

                return false;
            }

            Thread.Sleep(ComboBoxTypeAheadFocusDelayMs);

            if (!OpenComboBoxDropdown(session, comboBox))
            {
                _logger.LogWarning(
                    "ComboBox keyboard type-ahead skipped because dropdown could not be opened. combo={Combo}, item={Item}",
                    SafeElementName(comboBox),
                    itemName);

                return false;
            }

            Thread.Sleep(MenuExpandDelayMs);

            var list = FindDynamicComboBoxList(session, comboBox);

            if (list == null)
            {
                var expanded = IsComboBoxExpanded(comboBox);

                if (!expanded)
                {
                    _logger.LogWarning(
                        "ComboBox keyboard type-ahead skipped because dropdown list was not detected and ComboBox is not expanded. combo={Combo}, item={Item}",
                        SafeElementName(comboBox),
                        itemName);

                    return false;
                }

                _logger.LogInformation(
                    "ComboBox keyboard type-ahead continuing without UIA List because ComboBox reports Expanded. combo={Combo}, item={Item}",
                    SafeElementName(comboBox),
                    itemName);
            }

            _logger.LogInformation(
                "ComboBox keyboard type-ahead fallback started. combo={Combo}, item={Item}",
                SafeElementName(comboBox),
                itemName);

            // Use Keyboard.Type for literal character-by-character input so that special
            // characters in item names (e.g. "+", "^", "%") are never misinterpreted as
            // modifier keys the way SendKeysString would treat them.
            Keyboard.Type(itemName);

            Thread.Sleep(ComboBoxTypeAheadCommitDelayMs);

            Keyboard.Press(VirtualKeyShort.RETURN);
            Keyboard.Release(VirtualKeyShort.RETURN);

            Thread.Sleep(ComboBoxTypeAheadCommitDelayMs);

            var verified = VerifyComboBoxSelectedValue(session, comboBox, itemName);

            if (verified)
            {
                _logger.LogInformation(
                    "ComboBox keyboard type-ahead verified. combo={Combo}, item={Item}",
                    SafeElementName(comboBox),
                    itemName);
            }
            else
            {
                _logger.LogWarning(
                    "ComboBox keyboard type-ahead did not verify. requested={Requested}, actual={Actual}, combo={Combo}",
                    itemName,
                    GetComboBoxCurrentValue(session, comboBox),
                    SafeElementName(comboBox));
            }

            return verified;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox keyboard type-ahead failed. combo={Combo}, item={Item}",
                SafeElementName(comboBox),
                itemName);

            return false;
        }
    }

    private bool TrySelectComboBoxByVisibleAnchorWindowSearch(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue)
    {
        try
        {
            if (!ResetComboBoxDropdownToTop(session, comboBox))
                return false;

            string? previousSignature = null;

            for (var window = 0; window < ComboBoxAnchorWindowSearchMaxWindows; window++)
            {
                if (!EnsureComboBoxDropdownOpen(session, comboBox))
                    return false;

                Thread.Sleep(ComboBoxAnchorReadDelayMs);

                var visibleItems = GetCurrentVisibleComboBoxItems(session, comboBox);
                var signature = BuildVisibleComboBoxItemsSignature(visibleItems);

                _logger.LogInformation(
                    "ComboBox visible anchor-window search. combo={Combo}, requested={Requested}, window={Window}, visibleCount={Count}, signature={Signature}",
                    SafeElementName(comboBox),
                    requestedValue,
                    window,
                    visibleItems.Count,
                    signature);

                var exactItem = FindExactVisibleComboBoxItem(visibleItems, requestedValue);
                if (exactItem != null)
                {
                    _logger.LogInformation(
                        "ComboBox visible anchor-window search found exact visible item. combo={Combo}, requested={Requested}, window={Window}, item={Item}",
                        SafeElementName(comboBox),
                        requestedValue,
                        window,
                        SafeElementName(exactItem));

                    if (!ActivateComboBoxListItem(exactItem, requestedValue))
                        return false;

                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);
                    return VerifyComboBoxSelectedValue(session, comboBox, requestedValue);
                }

                if (visibleItems.Count == 0)
                {
                    _logger.LogInformation(
                        "ComboBox visible anchor-window search stopped because no visible items were read. combo={Combo}, requested={Requested}, window={Window}",
                        SafeElementName(comboBox),
                        requestedValue,
                        window);

                    break;
                }

                if (window > 0 &&
                    string.Equals(previousSignature, signature, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogInformation(
                        "ComboBox visible anchor-window search stopped because visible window did not change. combo={Combo}, requested={Requested}, signature={Signature}",
                        SafeElementName(comboBox),
                        requestedValue,
                        signature);

                    break;
                }

                previousSignature = signature;

                var anchorItem = visibleItems[^1];
                var downCount = visibleItems.Count;

                if (!ClickComboBoxAnchorItem(anchorItem, requestedValue))
                    return false;

                Thread.Sleep(ComboBoxAnchorMoveDelayMs);

                if (!EnsureComboBoxDropdownOpen(session, comboBox))
                    return false;

                Thread.Sleep(ComboBoxAnchorMoveDelayMs);

                if (!PressComboBoxDownKeys(comboBox, downCount, requestedValue))
                    return false;
            }

            _logger.LogWarning(
                "ComboBox visible anchor-window search did not find requested value. combo={Combo}, requested={Requested}, maxWindows={MaxWindows}, actual={Actual}",
                SafeElementName(comboBox),
                requestedValue,
                ComboBoxAnchorWindowSearchMaxWindows,
                GetComboBoxCurrentValue(session, comboBox));

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox visible anchor-window search failed. combo={Combo}, requested={Requested}",
                SafeElementName(comboBox),
                requestedValue);

            return false;
        }
    }

    private List<AutomationElement> GetCurrentVisibleComboBoxItems(
        AutomationSession session,
        AutomationElement comboBox)
    {
        return FindDynamicComboBoxItems(
                session,
                comboBox,
                maxItems: ComboBoxPagedSearchMaxVisibleItems)
            .Where(x =>
                !string.IsNullOrWhiteSpace(SafeElementName(x)) ||
                !string.IsNullOrWhiteSpace(SafeElementAutomationId(x)))
            .Take(ComboBoxPagedSearchMaxVisibleItems)
            .ToList();
    }

    private bool EnsureComboBoxDropdownOpen(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            if (IsComboBoxExpanded(comboBox) || FindDynamicComboBoxList(session, comboBox) != null)
                return true;
        }
        catch
        {
            // Fall through to the normal open path.
        }

        return OpenComboBoxDropdown(session, comboBox);
    }

    private AutomationElement? FindExactVisibleComboBoxItem(
        IReadOnlyList<AutomationElement> visibleItems,
        string requestedValue)
    {
        foreach (var item in visibleItems)
        {
            var name = SafeElementName(item);
            var automationId = SafeElementAutomationId(item);

            if (string.Equals(
                    NormalizeMenuText(name),
                    NormalizeMenuText(requestedValue),
                    StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }

            if (string.Equals(
                    NormalizeMenuText(automationId),
                    NormalizeMenuText(requestedValue),
                    StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }
        }

        return null;
    }

    private bool ClickComboBoxAnchorItem(
        AutomationElement anchorItem,
        string requestedValue)
    {
        try
        {
            var rect = anchorItem.BoundingRectangle;

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return false;

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            _logger.LogInformation(
                "ComboBox anchor click. requested={Requested}, anchor={Anchor}, point={Point}",
                requestedValue,
                SafeElementName(anchorItem),
                point);

            return SendInstantLeftClick(point, "ComboBox Anchor Click");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox anchor click failed. requested={Requested}, anchor={Anchor}",
                requestedValue,
                SafeElementName(anchorItem));

            return false;
        }
    }

    private bool PressComboBoxDownKeys(
        AutomationElement comboBox,
        int count,
        string requestedValue)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxAnchorDownKeys"))
                return false;

            Thread.Sleep(ComboBoxAnchorMoveDelayMs);

            for (var i = 0; i < count; i++)
            {
                Keyboard.Press(VirtualKeyShort.DOWN);
                Keyboard.Release(VirtualKeyShort.DOWN);
                Thread.Sleep(ComboBoxAnchorMoveDelayMs);
            }

            Thread.Sleep(ComboBoxAnchorReadDelayMs);

            _logger.LogInformation(
                "ComboBox anchor movement completed. requested={Requested}, downCount={Count}",
                requestedValue,
                count);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox anchor DOWN movement failed. requested={Requested}, count={Count}",
                requestedValue,
                count);

            return false;
        }
    }

    private string BuildVisibleComboBoxItemsSignature(
        IReadOnlyCollection<AutomationElement> items)
    {
        if (items.Count == 0)
            return string.Empty;

        return string.Join(
            "||",
            items.Select(item =>
            {
                var name = NormalizeMenuText(SafeElementName(item));
                var aid = NormalizeMenuText(SafeElementAutomationId(item));
                return $"{name}|{aid}";
            }));
    }

    private bool TrySelectComboBoxByKeyboardStepSearch(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue)
    {
        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxKeyboardStepSearch"))
            {
                _logger.LogWarning(
                    "ComboBox keyboard step search skipped because focus could not be confirmed. combo={Combo}, value={Value}",
                    SafeElementName(comboBox),
                    requestedValue);

                return false;
            }

            if (!OpenComboBoxDropdown(session, comboBox))
            {
                _logger.LogWarning(
                    "ComboBox keyboard step search skipped because dropdown could not be opened. combo={Combo}, value={Value}",
                    SafeElementName(comboBox),
                    requestedValue);

                return false;
            }

            Thread.Sleep(MenuExpandDelayMs);

            Keyboard.Press(VirtualKeyShort.HOME);
            Keyboard.Release(VirtualKeyShort.HOME);

            Thread.Sleep(ComboBoxKeyboardStepReadDelayMs);

            string? previousText = null;
            var repeatedSameTextCount = 0;

            for (var step = 0; step < ComboBoxKeyboardStepSearchMaxSteps; step++)
            {
                var currentText = GetCurrentHighlightedComboBoxItemText(session, comboBox);

                if (step < 10 || step % ComboBoxKeyboardStepLogEvery == 0)
                {
                    _logger.LogInformation(
                        "ComboBox keyboard step search. combo={Combo}, requested={Requested}, step={Step}, current={Current}",
                        SafeElementName(comboBox),
                        requestedValue,
                        step,
                        currentText);
                }

                if (ComboBoxTextMatchesExactly(currentText, requestedValue))
                {
                    Keyboard.Press(VirtualKeyShort.RETURN);
                    Keyboard.Release(VirtualKeyShort.RETURN);

                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxSelectedValue(session, comboBox, requestedValue))
                    {
                        _logger.LogInformation(
                            "ComboBox keyboard step search selected requested value. combo={Combo}, value={Value}, step={Step}",
                            SafeElementName(comboBox),
                            requestedValue,
                            step);

                        return true;
                    }

                    _logger.LogWarning(
                        "ComboBox keyboard step search pressed Enter but verification failed. requested={Requested}, actual={Actual}, step={Step}",
                        requestedValue,
                        GetComboBoxCurrentValue(session, comboBox),
                        step);

                    return false;
                }

                if (!string.IsNullOrWhiteSpace(previousText) &&
                    ComboBoxTextMatchesExactly(currentText, previousText))
                {
                    repeatedSameTextCount++;
                }
                else
                {
                    repeatedSameTextCount = 0;
                }

                if (repeatedSameTextCount >= 5 && step > 5)
                {
                    _logger.LogInformation(
                        "ComboBox keyboard step search stopped because DOWN did not change current highlighted item. combo={Combo}, current={Current}, step={Step}",
                        SafeElementName(comboBox),
                        currentText,
                        step);

                    break;
                }

                previousText = currentText;

                Keyboard.Press(VirtualKeyShort.DOWN);
                Keyboard.Release(VirtualKeyShort.DOWN);

                Thread.Sleep(ComboBoxKeyboardStepDelayMs + ComboBoxKeyboardStepReadDelayMs);
            }

            _logger.LogWarning(
                "ComboBox keyboard step search did not find requested value. combo={Combo}, requested={Requested}, maxSteps={MaxSteps}, actual={Actual}",
                SafeElementName(comboBox),
                requestedValue,
                ComboBoxKeyboardStepSearchMaxSteps,
                GetComboBoxCurrentValue(session, comboBox));

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox keyboard step search failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                requestedValue);

            return false;
        }
    }

    private static bool IsComboBoxExpanded(AutomationElement comboBox)
    {
        try
        {
            if (comboBox.Patterns.ExpandCollapse.IsSupported)
            {
                return comboBox.Patterns.ExpandCollapse.Pattern.ExpandCollapseState
                    == ExpandCollapseState.Expanded;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static string GetComboBoxExpandState(AutomationElement comboBox)
    {
        try
        {
            if (comboBox.Patterns.ExpandCollapse.IsSupported)
                return comboBox.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.ToString();
        }
        catch
        {
            return "Unknown";
        }

        return "NotSupported";
    }

    private AutomationElement? FindDynamicComboBoxList(AutomationSession session, AutomationElement comboBox)
    {
        try
        {
            var comboRect = comboBox.BoundingRectangle;
            var root = GetSearchRootForDynamicPopup(session, comboBox) ?? session.Automation.GetDesktop();
            var cf = session.Automation.ConditionFactory;

            foreach (var list in root.FindAllChildren(cf.ByControlType(ControlType.List)))
            {
                if (IsComboListNearCombo(list, comboRect))
                    return list;
            }

            var checkedCount = 0;
            foreach (var list in root.FindAllDescendants(cf.ByControlType(ControlType.List)))
            {
                checkedCount++;

                if (checkedCount > MaxComboBoxDropdownListCandidates)
                    break;

                if (IsComboListNearCombo(list, comboRect))
                    return list;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindDynamicComboBoxList failed");
        }

        return null;
    }

    private bool IsComboListNearCombo(AutomationElement list, Rectangle comboRect)
    {
        try
        {
            var rect = list.BoundingRectangle;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return false;

            return rect.Top >= comboRect.Bottom - ComboBoxDropdownVerticalTolerancePx &&
                   rect.Left <= comboRect.Right + 150 &&
                   rect.Right >= comboRect.Left - 150;
        }
        catch
        {
            return false;
        }
    }

    private List<AutomationElement> GetListItemsBounded(
        AutomationSession session,
        AutomationElement list,
        int maxItems)
    {
        var results = new List<AutomationElement>();

        try
        {
            if (maxItems <= 0)
                return results;

            var walker = session.Automation.TreeWalkerFactory.GetControlViewWalker();
            var visited = 0;
            var maxVisitedNodes = GetComboBoxTraversalNodeLimit(maxItems);

            CollectListItemsBounded(
                walker,
                list,
                results,
                maxItems,
                ref visited,
                maxVisitedNodes,
                depth: 0,
                maxDepth: 8);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetListItemsBounded failed");
        }

        return results;
    }

    private static int GetComboBoxTraversalNodeLimit(int maxItems)
    {
        if (maxItems >= 1000)
            return maxItems;

        return Math.Max(maxItems * 20, 200);
    }

    private bool CollectListItemsBounded(
        ITreeWalker walker,
        AutomationElement parent,
        List<AutomationElement> results,
        int maxItems,
        ref int visited,
        int maxVisitedNodes,
        int depth,
        int maxDepth)
    {
        if (results.Count >= maxItems || visited >= maxVisitedNodes)
            return true;

        if (depth > maxDepth)
            return false;

        AutomationElement? child;

        try
        {
            child = walker.GetFirstChild(parent);
        }
        catch
        {
            return false;
        }

        while (child != null && results.Count < maxItems && visited < maxVisitedNodes)
        {
            visited++;

            var isListItem = false;

            try
            {
                isListItem = child.ControlType == ControlType.ListItem;
            }
            catch
            {
                // ignore
            }

            if (isListItem)
            {
                if (IsNamedListItem(child))
                    results.Add(child);
            }
            else if (CollectListItemsBounded(
                         walker,
                         child,
                         results,
                         maxItems,
                         ref visited,
                         maxVisitedNodes,
                         depth + 1,
                         maxDepth))
            {
                return true;
            }

            try
            {
                child = walker.GetNextSibling(child);
            }
            catch
            {
                break;
            }
        }

        return results.Count >= maxItems || visited >= maxVisitedNodes;
    }

    private bool IsNamedListItem(AutomationElement item)
    {
        try
        {
            return !string.IsNullOrWhiteSpace(SafeElementName(item)) ||
                   !string.IsNullOrWhiteSpace(SafeElementAutomationId(item));
        }
        catch
        {
            return false;
        }
    }

    private string GetComboBoxCurrentValue(AutomationSession session, AutomationElement comboBox)
    {
        try
        {
            if (comboBox.Patterns.Value.IsSupported)
            {
                string value = comboBox.Patterns.Value.Pattern.Value;
                if (!string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            var cf = session.Automation.ConditionFactory;

            var editChild = comboBox.FindFirstDescendant(cf.ByControlType(ControlType.Edit));
            if (editChild != null)
            {
                try
                {
                    if (editChild.Patterns.Value.IsSupported)
                    {
                        string editValue = editChild.Patterns.Value.Pattern.Value;
                        if (!string.IsNullOrWhiteSpace(editValue))
                            return editValue.Trim();
                    }
                }
                catch
                {
                    // ignore
                }

                var editName = SafeElementName(editChild);
                if (!string.IsNullOrWhiteSpace(editName))
                    return editName.Trim();
            }

            var textChild = comboBox.FindFirstDescendant(cf.ByControlType(ControlType.Text));
            if (textChild != null)
            {
                var text = SafeElementName(textChild);
                if (!string.IsNullOrWhiteSpace(text))
                    return text.Trim();
            }

            var selectedItem = FindSelectedListItemBounded(session, comboBox);

            if (selectedItem != null)
            {
                var selectedName = SafeElementName(selectedItem);
                if (!string.IsNullOrWhiteSpace(selectedName))
                    return selectedName.Trim();
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            var name = SafeElementName(comboBox);
            if (!string.IsNullOrWhiteSpace(name))
                return name.Trim();
        }
        catch
        {
            // ignore
        }

        return string.Empty;
    }

    private AutomationElement? FindSelectedListItemBounded(
        AutomationSession session,
        AutomationElement comboBox)
    {
        try
        {
            var candidates = GetListItemsBounded(session, comboBox, ComboBoxVisibleItemSearchLimit);

            foreach (var item in candidates)
            {
                try
                {
                    if (item.Patterns.SelectionItem.IsSupported &&
                        item.Patterns.SelectionItem.PatternOrDefault?.IsSelected == true)
                    {
                        return item;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            var list = FindDynamicComboBoxList(session, comboBox);
            if (list == null)
                return null;

            candidates = GetListItemsBounded(session, list, ComboBoxVisibleItemSearchLimit);

            foreach (var item in candidates)
            {
                try
                {
                    if (item.Patterns.SelectionItem.IsSupported &&
                        item.Patterns.SelectionItem.PatternOrDefault?.IsSelected == true)
                    {
                        return item;
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }
        catch
        {
            // ignore
        }

        return null;
    }

    private bool VerifyComboBoxSelectedValue(
        AutomationSession session,
        AutomationElement comboBox,
        string expectedValue)
    {
        var actual = GetComboBoxCurrentValue(session, comboBox);

        var expectedNorm = NormalizeMenuText(expectedValue);
        var actualNorm = NormalizeMenuText(actual);

        var matched = string.Equals(
            actualNorm,
            expectedNorm,
            StringComparison.OrdinalIgnoreCase);

        _logger.LogInformation(
            "ComboBox selection verification. expected={Expected}, actual={Actual}, matched={Matched}",
            expectedValue,
            actual,
            matched);

        return matched;
    }

    private bool ActivateComboBoxListItem(AutomationElement item, string itemName)
    {
        try
        {
            var rect = item.BoundingRectangle;
            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round(rect.Left + (double)Math.Max(
                        ComboBoxLeftEdgeMinOffsetPx,
                        Math.Min(ComboBoxLeftEdgeMaxOffsetPx, rect.Width / ComboBoxLeftEdgeOffsetDivisor))),
                    (int)Math.Round(rect.Top + (rect.Height / 2.0)));

                if (SendInstantLeftClick(point, $"Select ComboBox item {itemName}"))
                {
                    Thread.Sleep(MenuActionDelayMs);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox ListItem physical click failed for {Item}", itemName);
        }

        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox ListItem SelectionItem failed for {Item}", itemName);
        }

        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox ListItem Invoke failed for {Item}", itemName);
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(MenuActionDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ComboBox ListItem Focus+Enter failed for {Item}", itemName);
        }

        return false;
    }

    private bool OpenDynamicMenuParent(AutomationElement menuItem)
    {
        try
        {
            if (menuItem.Patterns.ExpandCollapse.IsSupported)
            {
                menuItem.Patterns.ExpandCollapse.Pattern.Expand();
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ExpandCollapse failed for dynamic menu parent {Menu}", SafeElementName(menuItem));
        }

        try
        {
            ActivateMenuItem(menuItem);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ActivateMenuItem failed for dynamic menu parent {Menu}", SafeElementName(menuItem));
            return false;
        }
    }

    private AutomationElement? FindActiveContextMenuPopup(AutomationSession session)
    {
        try
        {
            foreach (var candidate in GetContextMenuPopupCandidates(session))
            {
                var items = GetContextMenuItems(session, candidate, maxItems: 1);
                if (items.Count > 0)
                    return candidate;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to detect active context menu popup.");
        }

        return null;
    }

    private AutomationElement? FindContextSubMenuPopup(
        AutomationSession session,
        AutomationElement submenuItem)
    {
        try
        {
            var itemRect = submenuItem.BoundingRectangle;

            foreach (var candidate in GetContextMenuPopupCandidates(session))
            {
                var rect = candidate.BoundingRectangle;

                if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                    continue;

                var nearSubmenu =
                    rect.Left >= itemRect.Right - SubmenuHorizontalProximityPx &&
                    rect.Top <= itemRect.Bottom + SubmenuVerticalProximityPx &&
                    rect.Bottom >= itemRect.Top - SubmenuVerticalProximityPx;

                if (nearSubmenu && GetContextMenuItems(session, candidate, maxItems: 1).Count > 0)
                    return candidate;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to detect context submenu popup.");
        }

        return null;
    }

    private List<AutomationElement> GetContextMenuPopupCandidates(AutomationSession session)
    {
        var results = new List<AutomationElement>();

        try
        {
            var desktop = session.Automation.GetDesktop();
            var queue = new Queue<(AutomationElement Element, int Depth)>();

            foreach (var child in desktop.FindAllChildren())
                queue.Enqueue((child, 1));

            while (queue.Count > 0 && results.Count < MaxApplicationContextMenuCandidates)
            {
                var (element, depth) = queue.Dequeue();

                try
                {
                    var ct = element.ControlType;
                    var rect = element.BoundingRectangle;

                    if (IsContextMenuContainerType(ct) &&
                        !rect.IsEmpty &&
                        rect.Width > 0 &&
                        rect.Height > 0)
                    {
                        results.Add(element);
                    }

                    if (depth >= MaxApplicationContextMenuSearchDepth)
                        continue;

                    foreach (var child in element.FindAllChildren())
                    {
                        queue.Enqueue((child, depth + 1));
                        if (queue.Count + results.Count >= MaxApplicationContextMenuCandidates)
                            break;
                    }
                }
                catch
                {
                    // ignore stale candidate
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed collecting context menu popup candidates.");
        }

        return results;
    }

    private List<AutomationElement> GetContextMenuItems(
        AutomationSession session,
        AutomationElement menuRoot,
        int maxItems = MaxApplicationContextMenuPlaybackItems)
    {
        var results = new List<AutomationElement>();
        var seenKeys = new HashSet<string>(StringComparer.Ordinal);

        try
        {
            var queue = new Queue<AutomationElement>();
            queue.Enqueue(menuRoot);

            while (queue.Count > 0 && results.Count < maxItems)
            {
                var current = queue.Dequeue();
                IReadOnlyList<AutomationElement> children;

                try
                {
                    children = current.FindAllChildren();
                }
                catch
                {
                    continue;
                }

                foreach (var child in children)
                {
                    if (results.Count >= maxItems)
                        break;

                    try
                    {
                        var ct = child.ControlType;
                        var name = SafeElementName(child);

                        if (IsContextMenuItemType(ct) && !string.IsNullOrWhiteSpace(name))
                        {
                            var key = GetElementDedupeKey(child);
                            if (key == null || seenKeys.Add(key))
                                results.Add(child);
                        }

                        if (IsContextMenuContainerType(ct))
                            queue.Enqueue(child);
                    }
                    catch
                    {
                        // ignore stale child
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed reading context menu items.");
        }

        return results;
    }

    private AutomationElement? FindContextMenuItemByText(
        AutomationSession session,
        AutomationElement menuRoot,
        string itemText)
    {
        var requested = NormalizeMenuText(itemText);

        foreach (var item in GetContextMenuItems(session, menuRoot, MaxApplicationContextMenuPlaybackItems))
        {
            var name = NormalizeMenuText(SafeElementName(item));

            if (string.Equals(name, requested, StringComparison.OrdinalIgnoreCase))
                return item;
        }

        return null;
    }

    private bool ActivateContextMenuItem(
        AutomationElement item,
        string itemName)
    {
        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }

            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }

            var rect = item.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round(rect.Left + rect.Width / 2.0),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                return SendInstantLeftClick(point, $"Context Menu Item: {itemName}");
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed activating context menu item {Item}", itemName);
            return false;
        }
    }

    private bool OpenContextSubMenu(
        AutomationElement item,
        string itemName)
    {
        try
        {
            if (item.Patterns.ExpandCollapse.IsSupported)
            {
                item.Patterns.ExpandCollapse.Pattern.Expand();
                Thread.Sleep(MenuExpandDelayMs);
                return true;
            }

            var rect = item.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var point = new Point(
                    (int)Math.Round(rect.Right - CalculateSubmenuArrowOffset((double)rect.Width)),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                SendInstantLeftClick(point, $"Open Context Submenu: {itemName}");
                Thread.Sleep(MenuExpandDelayMs);
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed opening context submenu {Item}", itemName);
            return false;
        }
    }

    private static bool IsContextMenuContainerType(ControlType controlType)
    {
        return controlType == ControlType.Menu ||
               controlType == ControlType.Window ||
               controlType == ControlType.Pane ||
               controlType == ControlType.ToolBar ||
               controlType == ControlType.Custom;
    }

    private static bool IsContextMenuItemType(ControlType controlType)
    {
        return controlType == ControlType.MenuItem ||
               controlType == ControlType.ListItem ||
               controlType == ControlType.Button ||
               controlType == ControlType.Text;
    }

    private AutomationElement? FindDynamicMenuDropdown(AutomationSession session, AutomationElement parentMenuItem)
    {
        try
        {
            var parentName = SafeElementName(parentMenuItem);
            var parentRect = parentMenuItem.BoundingRectangle;

            var possibleContainers = FindDynamicMenuDropdownCandidates(session, parentMenuItem);

            foreach (var container in possibleContainers)
            {
                try
                {
                    var name = SafeElementName(container);
                    var rect = container.BoundingRectangle;

                    if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                        continue;

                    var nameLooksLikeDropdown =
                        !string.IsNullOrWhiteSpace(parentName) &&
                        !string.IsNullOrWhiteSpace(name) &&
                        name.Contains(parentName, StringComparison.OrdinalIgnoreCase) &&
                        name.Contains("DropDown", StringComparison.OrdinalIgnoreCase);

                    var nearParent =
                        rect.Top >= parentRect.Bottom - 10 &&
                        rect.Left <= parentRect.Right + 100 &&
                        rect.Right >= parentRect.Left - 100;

                    if (!nameLooksLikeDropdown && !nearParent)
                        continue;

                    var hasMenuItems = GetDynamicDropdownMenuItems(session, container, maxItems: 1).Count > 0;

                    if (hasMenuItems)
                    {
                        _logger.LogInformation(
                            "Found dynamic menu dropdown. parent={Parent}, dropdown={Dropdown}, controlType={ControlType}, bounds={Bounds}",
                            parentName,
                            name,
                            container.ControlType,
                            rect);

                        return container;
                    }
                }
                catch
                {
                    // ignore unstable UIA element
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindDynamicMenuDropdown failed in UiService");
        }

        return null;
    }

    private List<AutomationElement> FindDynamicMenuDropdownCandidates(
        AutomationSession session,
        AutomationElement parentMenuItem,
        int maxCandidates = 80)
    {
        var results = new List<AutomationElement>();

        try
        {
            var root = GetSearchRootForDynamicPopup(session, parentMenuItem) ?? session.Automation.GetDesktop();
            var cf = session.Automation.ConditionFactory;

            var controlTypes = new[]
            {
                ControlType.ToolBar,
                ControlType.Menu,
                ControlType.Pane,
                ControlType.Custom,
                ControlType.Window
            };

            foreach (var ct in controlTypes)
            {
                foreach (var child in root.FindAllChildren(cf.ByControlType(ct)))
                {
                    results.Add(child);
                    if (results.Count >= maxCandidates)
                        return results;
                }
            }

            foreach (var ct in controlTypes)
            {
                foreach (var item in root.FindAllDescendants(cf.ByControlType(ct)))
                {
                    results.Add(item);
                    if (results.Count >= maxCandidates)
                        return results;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindDynamicMenuDropdownCandidates failed");
        }

        return results;
    }

    private List<AutomationElement> GetDynamicDropdownMenuItems(
        AutomationSession session,
        AutomationElement dropdown,
        int maxItems = int.MaxValue)
    {
        var results = new List<AutomationElement>();
        var seenKeys = new HashSet<string>(StringComparer.Ordinal);

        try
        {
            var cf = session.Automation.ConditionFactory;

            foreach (var child in dropdown.FindAllChildren(cf.ByControlType(ControlType.MenuItem)))
            {
                AddMenuItemIfValid(child, results, seenKeys);

                if (results.Count >= maxItems)
                    return results;
            }

            if (results.Count == 0)
            {
                foreach (var item in dropdown.FindAllDescendants(cf.ByControlType(ControlType.MenuItem)))
                {
                    AddMenuItemIfValid(item, results, seenKeys);

                    if (results.Count >= maxItems)
                        return results;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetDynamicDropdownMenuItems bounded scan failed");
        }

        return results;
    }

    private void AddMenuItemIfValid(
        AutomationElement item,
        List<AutomationElement> results,
        HashSet<string> seenKeys)
    {
        if (string.IsNullOrWhiteSpace(SafeElementName(item)) &&
            string.IsNullOrWhiteSpace(SafeElementAutomationId(item)))
            return;

        var key = GetElementDedupeKey(item);

        if (key != null && !seenKeys.Add(key))
            return;

        results.Add(item);
    }

    private static string? GetElementDedupeKey(AutomationElement item)
    {
        try
        {
            var rect = item.BoundingRectangle;
            return string.Join(
                "|",
                item.ControlType,
                SafeElementAutomationId(item),
                SafeElementName(item),
                rect.Left,
                rect.Top,
                rect.Width,
                rect.Height);
        }
        catch
        {
            return null;
        }
    }

    private bool ActivateDynamicMenuItem(AutomationElement item, string itemName)
    {
        if (TryInstantPhysicalClick(item, $"ActivateDynamicMenuItem {itemName}"))
            return true;

        try
        {
            item.Click();
            Thread.Sleep(MenuActionDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Click failed for dynamic menu item {Item}", itemName);
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(MenuActionDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Focus+Enter failed for dynamic menu item {Item}", itemName);
        }

        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Invoke failed for dynamic menu item {Item}", itemName);
        }

        return false;
    }

    private static bool DynamicPathStartsWithParent(
        IReadOnlyList<string> pathParts,
        AutomationElement parentMenuItem)
    {
        if (pathParts.Count <= 1)
            return false;

        var firstPart = NormalizeMenuText(pathParts[0]);

        return string.Equals(firstPart, NormalizeMenuText(SafeElementName(parentMenuItem)), StringComparison.OrdinalIgnoreCase) ||
               string.Equals(firstPart, NormalizeMenuText(SafeElementAutomationId(parentMenuItem)), StringComparison.OrdinalIgnoreCase);
    }

    private static double CalculateSubmenuArrowOffset(double rectWidth)
    {
        return Math.Max(
            SubmenuArrowMinOffsetPx,
            Math.Min(SubmenuArrowMaxOffsetPx, rectWidth / SubmenuArrowWidthDivisor));
    }

    private AutomationElement? FindDynamicMenuItemByName(
        AutomationSession session,
        AutomationElement dropdown,
        string itemName)
    {
        return GetDynamicDropdownMenuItems(session, dropdown, MaxAssistiveDropdownItemsToDisplay * 4)
            .FirstOrDefault(x =>
                string.Equals(
                    NormalizeMenuText(SafeElementName(x)),
                    NormalizeMenuText(itemName),
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    NormalizeMenuText(SafeElementAutomationId(x)),
                    NormalizeMenuText(itemName),
                    StringComparison.OrdinalIgnoreCase));
    }

    private bool OpenDynamicSubMenuItem(AutomationElement item)
    {
        try
        {
            var rect = item.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var center = new Point(
                    (int)Math.Round(rect.Left + rect.Width / 2.0),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                Mouse.MoveTo(center);
                Thread.Sleep(MenuExpandDelayMs);

                var arrowOffset = CalculateSubmenuArrowOffset((double)rect.Width);
                var right = new Point(
                    (int)Math.Round((double)rect.Right - arrowOffset),
                    (int)Math.Round(rect.Top + rect.Height / 2.0));

                if (SendInstantLeftClick(right, $"Open submenu {SafeElementName(item)}"))
                {
                    Thread.Sleep(MenuExpandDelayMs);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenDynamicSubMenuItem physical failed");
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RIGHT);
            Thread.Sleep(MenuExpandDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenDynamicSubMenuItem RIGHT failed");
        }

        return false;
    }

    private AutomationElement? FindDynamicSubMenuDropdown(AutomationSession session, AutomationElement submenuItem)
    {
        try
        {
            var itemRect = submenuItem.BoundingRectangle;
            var root = GetSearchRootForDynamicPopup(session, submenuItem) ?? session.Automation.GetDesktop();
            foreach (var c in FindDynamicSubMenuDropdownCandidates(root))
            {
                try
                {
                    var rect = c.BoundingRectangle;
                    if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                        continue;

                    var nearSubmenu =
                        rect.Left >= itemRect.Right - SubmenuHorizontalProximityPx &&
                        rect.Top <= itemRect.Bottom + SubmenuVerticalProximityPx &&
                        rect.Bottom >= itemRect.Top - SubmenuVerticalProximityPx;

                    if (!nearSubmenu)
                        continue;

                    if (GetDynamicDropdownMenuItems(session, c, maxItems: 1).Count > 0)
                        return c;
                }
                catch
                {
                    // ignore unstable candidate
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindDynamicSubMenuDropdown failed");
        }

        return null;
    }

    private List<AutomationElement> FindDynamicSubMenuDropdownCandidates(AutomationElement root)
    {
        var results = new List<AutomationElement>();
        var queue = new Queue<(AutomationElement Element, int Depth)>();

        try
        {
            foreach (var child in root.FindAllChildren())
                queue.Enqueue((child, 1));

            while (queue.Count > 0 && results.Count < MaxDynamicDropdownCandidates)
            {
                var (element, depth) = queue.Dequeue();

                try
                {
                    if (DynamicMenuHelpers.IsDropdownContainerType(element.ControlType))
                        results.Add(element);

                    if (depth >= MaxDynamicPopupSearchDepth)
                        continue;

                    foreach (var child in element.FindAllChildren())
                    {
                        queue.Enqueue((child, depth + 1));
                        if (queue.Count + results.Count >= MaxDynamicDropdownCandidates)
                            break;
                    }
                }
                catch
                {
                    // ignore unstable UIA candidate
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindDynamicSubMenuDropdownCandidates failed");
        }

        return results;
    }

    private AutomationElement? OpenHeaderDropdownAndFindList(
        AutomationElement header,
        HeaderDropdownRegion region)
    {
        BringElementWindowToForeground(header);
        Thread.Sleep(WindowActivationDelayMs);

        var rect = header.BoundingRectangle;

        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            throw new InvalidOperationException("Header has invalid bounding rectangle.");

        AutomationElement? list = null;
        foreach (var (candidateRegion, point) in GridHeaderDropdownHelper.GetCandidatePoints(rect, region))
        {
            _logger.LogInformation(
                "Opening grid header dropdown. header={Header}, bounds={Bounds}, region={Region}, clickPoint={ClickPoint}",
                SafeElementName(header),
                rect,
                candidateRegion,
                point);

            SendInstantLeftClick(point, "OpenHeaderDropdown");
            Thread.Sleep(GridHeaderDropdownHelper.DropdownRetryDelayMs);

            list = FindRecentlyOpenedListNearHeader(header);
            if (list != null)
                return list;
        }

        Thread.Sleep(GridHeaderDropdownHelper.DropdownOpenDelayMs);
        return FindRecentlyOpenedListNearHeader(header);
    }

    private AutomationElement? FindRecentlyOpenedListNearHeader(AutomationElement header)
    {
        try
        {
            var session = RequireSession();
            var headerRect = header.BoundingRectangle;
            var desktop = session.Automation.GetDesktop();
            var cf = session.Automation.ConditionFactory;

            var lists = desktop.FindAllDescendants(cf.ByControlType(ControlType.List));

            foreach (var list in lists)
            {
                try
                {
                    var rect = list.BoundingRectangle;

                    if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                        continue;

                    if (GridHeaderDropdownHelper.IsListNearHeader(rect, headerRect))
                    {
                        _logger.LogInformation(
                            "Found header dropdown List near header. list={List}, bounds={Bounds}",
                            SafeElementName(list),
                            rect);

                        return list;
                    }
                }
                catch
                {
                    // Ignore unstable popup elements.
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindRecentlyOpenedListNearHeader failed");
        }

        return null;
    }

    private List<AutomationElement> GetListItems(AutomationElement list)
    {
        try
        {
            var cf = RequireSession().Automation.ConditionFactory;

            return list
                .FindAllDescendants(cf.ByControlType(ControlType.ListItem))
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private bool ActivateDropdownListItem(
        AutomationElement item,
        string itemName,
        DropdownItemClickRegion region = DropdownItemClickRegion.LeftCenter)
    {
        var rect = item.BoundingRectangle;
        var safeItemName = SanitizeValue(itemName);

        foreach (var candidateRegion in GetDropdownItemRegionOrder(region))
        {
            try
            {
                var point = GetDropdownItemClickPoint(rect, candidateRegion);

                _logger.LogInformation(
                    "Trying dropdown item click. item={Item}, region={Region}, point={Point}, bounds={Bounds}",
                    safeItemName,
                    candidateRegion,
                    point,
                    rect);

                if (SendInstantLeftClick(point, $"SelectHeaderDropdownItem {safeItemName} at {candidateRegion}"))
                {
                    Thread.Sleep(DropdownItemPhysicalClickSettleMs);

                    if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                        return true;

                    _logger.LogWarning(
                        "Dropdown item physical click was sent but selection was not verified. item={Item}, region={Region}",
                        safeItemName,
                        candidateRegion);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(
                    ex,
                    "Dropdown item physical click failed for {Item} at region {Region}",
                    safeItemName,
                    candidateRegion);
            }
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.SPACE);
            Thread.Sleep(DropdownItemFallbackDelayMs);

            if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Dropdown item Focus+Space failed for {Item}", safeItemName);
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(DropdownItemFallbackDelayMs);

            if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Dropdown item Focus+Enter failed for {Item}", safeItemName);
        }

        try
        {
            if (item.Patterns.Toggle.IsSupported)
            {
                item.Patterns.Toggle.Pattern.Toggle();
                Thread.Sleep(DropdownItemFallbackDelayMs);

                if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                    return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Dropdown item Toggle failed for {Item}", safeItemName);
        }

        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(DropdownItemFallbackDelayMs);

                if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                    return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Dropdown item SelectionItem.Select failed for {Item}", safeItemName);
        }

        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(DropdownItemFallbackDelayMs);

                if (VerifyHeaderDropdownItemSelection(item, safeItemName))
                    return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Dropdown item Invoke failed for {Item}", safeItemName);
        }

        return false;
    }

    private bool VerifyHeaderDropdownItemSelection(
        AutomationElement item,
        string itemName)
    {
        try
        {
            // Access BoundingRectangle; if the item is stale, the dropdown was likely dismissed after selection.
            _ = item.BoundingRectangle;
        }
        catch
        {
            return true;
        }

        try
        {
            if (item.Patterns.SelectionItem.IsSupported &&
                item.Patterns.SelectionItem.Pattern.IsSelected)
                return true;
        }
        catch
        {
            // ignore
        }

        try
        {
            if (item.Patterns.Toggle.IsSupported &&
                item.Patterns.Toggle.Pattern.ToggleState == ToggleState.On)
                return true;
        }
        catch
        {
            // ignore
        }

        _logger.LogWarning("Dropdown item selection could not be verified. item={Item}", itemName);
        return false;
    }

    private static Point GetDropdownItemClickPoint(
        RectangleF rect,
        DropdownItemClickRegion region)
    {
        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            throw new InvalidOperationException("Dropdown ListItem has invalid bounding rectangle.");

        var padX = Math.Max(DropdownItemMinPadX, Math.Min(DropdownItemMaxPadX, rect.Width / DropdownItemPadXDivisor));
        var padY = Math.Max(DropdownItemMinPadY, Math.Min(DropdownItemMaxPadY, rect.Height / DropdownItemPadYDivisor));

        return region switch
        {
            DropdownItemClickRegion.LeftCenter => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            DropdownItemClickRegion.Center => new Point(
                (int)Math.Round(rect.Left + rect.Width / 2),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            DropdownItemClickRegion.RightCenter => new Point(
                (int)Math.Round(rect.Right - padX),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            DropdownItemClickRegion.UpperLeft => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Top + padY)),

            DropdownItemClickRegion.LowerLeft => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Bottom - padY)),

            _ => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Top + rect.Height / 2))
        };
    }

    private static IReadOnlyList<DropdownItemClickRegion> GetDropdownItemRegionOrder(
        DropdownItemClickRegion region)
    {
        if (region != DropdownItemClickRegion.ProbeAll)
            return new[] { region };

        return new[]
        {
            DropdownItemClickRegion.LeftCenter,
            DropdownItemClickRegion.UpperLeft,
            DropdownItemClickRegion.LowerLeft,
            DropdownItemClickRegion.Center,
            DropdownItemClickRegion.RightCenter
        };
    }

    private static DropdownItemClickRegion ParseDropdownItemRegion(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return DropdownItemClickRegion.LeftCenter;

        var normalized = value
            .Trim()
            .Replace("-", string.Empty)
            .Replace("_", string.Empty)
            .Replace(" ", string.Empty)
            .ToLowerInvariant();

        return normalized switch
        {
            "leftcenter" or "left" => DropdownItemClickRegion.LeftCenter,
            "center" => DropdownItemClickRegion.Center,
            "rightcenter" or "right" => DropdownItemClickRegion.RightCenter,
            "upperleft" or "topleft" => DropdownItemClickRegion.UpperLeft,
            "lowerleft" or "bottomleft" => DropdownItemClickRegion.LowerLeft,
            "probeall" or "all" or "auto" => DropdownItemClickRegion.ProbeAll,
            _ => DropdownItemClickRegion.LeftCenter
        };
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

    private bool TryInstantPhysicalRightClick(AutomationElement element, string actionName)
    {
        try
        {
            BringElementWindowToForeground(element);
            Thread.Sleep(WindowActivationDelayMs);

            var rect = element.BoundingRectangle;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return false;

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            if (!SetCursorPos(point.X, point.Y))
                return false;

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
                            dwFlags = MOUSEEVENTF_RIGHTDOWN,
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
                            dwFlags = MOUSEEVENTF_RIGHTUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                }
            };

            var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());

            _logger.LogInformation(
                "{ActionName}: Physical right-click sent={Sent}, point={Point}, element={Element}",
                actionName,
                sent,
                point,
                SafeElementName(element));

            return sent == inputs.Length;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "{ActionName}: physical right-click failed", actionName);
            return false;
        }
    }

    private bool SendMouseWheel(int wheelClicks)
    {
        try
        {
            if (wheelClicks == 0)
                return true;

            var wheelDelta = wheelClicks * 120;

            var input = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx = 0,
                        dy = 0,
                        // INPUT.MOUSEINPUT.mouseData is uint; negative wheel deltas are
                        // represented as two's complement unsigned values.
                        mouseData = unchecked((uint)wheelDelta),
                        dwFlags = MOUSEEVENTF_WHEEL,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            var sent = SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
            return sent == 1;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SendMouseWheel failed for wheelClicks={WheelClicks}", wheelClicks);
            return false;
        }
    }

    private void BringActiveWindowToForeground()
    {
        var session = RequireSession();
        var root = GetWindowRoot(session);
        BringElementWindowToForeground(root);
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

    private static string SafeElementClassName(AutomationElement element)
    {
        try
        {
            return SanitizeValue(element.ClassName);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int? SafeProcessId(AutomationElement element)
    {
        try
        {
            return element.Properties.ProcessId.ValueOrDefault;
        }
        catch
        {
            return null;
        }
    }

    private static bool? SafeIsOffscreen(AutomationElement element)
    {
        try
        {
            return element.IsOffscreen;
        }
        catch
        {
            return null;
        }
    }

    private static bool? SafeIsEnabled(AutomationElement element)
    {
        try
        {
            return element.IsEnabled;
        }
        catch
        {
            return null;
        }
    }

    private static string SafeElementControlType(AutomationElement element)
    {
        try
        {
            return element.ControlType.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeBoundingRectangle(AutomationElement element)
    {
        try
        {
            return element.BoundingRectangle.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int SafeMenuItemChildCount(AutomationElement item, ConditionFactory cf)
    {
        try
        {
            return item.FindAllDescendants(cf.ByControlType(ControlType.MenuItem)).Length;
        }
        catch
        {
            return -1;
        }
    }

    private static List<object> GetMenuItemDescendantSummary(
        AutomationElement parent,
        ConditionFactory cf)
    {
        try
        {
            return parent
                .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                .Select(e => new
                {
                    name = SafeElementName(e),
                    automationId = SafeElementAutomationId(e),
                    controlType = SafeElementControlType(e)
                })
                .Cast<object>()
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static List<string> GetMenuItemDescendantNames(
        AutomationElement parent,
        ConditionFactory cf)
    {
        try
        {
            return parent
                .FindAllDescendants(cf.ByControlType(ControlType.MenuItem))
                .Select(SafeElementName)
                .ToList();
        }
        catch
        {
            return [];
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
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
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
