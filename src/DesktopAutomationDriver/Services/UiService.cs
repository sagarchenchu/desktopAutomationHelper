using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DesktopAutomationDriver.Models;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
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
public partial class UiService : IUiService
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
    private const int DragMoveToStartDelayMs = 120;
    private const int DragMouseDownHoldDelayMs = 200;
    private const int DragMouseUpSettleDelayMs = 120;
    private const int DragMinimumStepDelayMs = 15;
    private const string ClearSelectAllBackspaceKeys = "^a{BACKSPACE}";
    private const string ClearSelectAllDeleteKeys = "^a{DELETE}";
    private const int MenuExpandDelayMs = 250;
    private const int MenuActionDelayMs = 150;
    private const int MenuFocusDelayMs = 75;
    private const int KeyboardInputReadyDelayMs = 100;
    private const int ComboBoxSelectionCommitDelayMs = 150;
    // Kept separate from selection commit delay so optional TAB-blur fallback timing can diverge if re-enabled.
    private const int ComboBoxBlurCommitDelayMs = ComboBoxSelectionCommitDelayMs;
    private const int ComboBoxPostCommitCollapseTimeoutMs = 2500;
    private const int ComboBoxPostCommitPollDelayMs = 100;
    private const int ComboBoxPostCommitStableDelayMs = 500;
    private const int ComboBoxSingleSelectionTimeoutMs = 20000;
    private const int SmallComboBoxSelectionTimeoutMs = 4000;
    private const int SmallComboBoxMaxScrollAttempts = 1;
    private const bool ComboBoxAllowTabBlurCommitFallback = false;
    private const int ComboBoxRefetchRectangleTolerancePx = 3;
    // When the scroll target is found, we cap fallback wheel attempts to prevent over-scrolling.
    // 5 attempts is sufficient for the target to settle into view across typical virtual lists.
    private const int ScrollTargetFoundMaxFallbackAttempts = 5;
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
    private const int ComboBoxDirectUiaSearchLimit = 500;
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
    private const int ContextMenuNearClickMarginLeading = 30;
    private const int ContextMenuNearClickMarginTrailing = 250;
    private const int ContextMenuTopLeftProximityThreshold = 80;
    private const int ContextMenuFarFromClickThreshold = 300;
    private const int ContextMenuMinimumCandidateScore = -50;
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

    private static readonly TimeSpan ElementCacheDuration = TimeSpan.FromSeconds(10);
    private static readonly object ElementCacheLock = new();
    private static readonly Dictionary<string, CachedElementMatch> ElementCache = new();

    private sealed record CachedElementMatch(AutomationElement Element, DateTime ExpiresAt);

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

    private readonly DesktopAutomationDriver.Services.Resolution.ElementResolver _newResolver;
    private readonly DesktopAutomationDriver.Services.ElementResolution.ElementResolver _pywinautoResolver;

    private const bool UsePywinautoStyleResolver = true;
    private readonly INativeUiaComboBoxService _nativeUiaComboBoxService;

    public UiService(
        IUiSessionContext ctx,
        ILogger<UiService> logger,
        INativeUiaComboBoxService nativeUiaComboBoxService)
    {
        _ctx = ctx;
        _logger = logger;
        _nativeUiaComboBoxService = nativeUiaComboBoxService;
        _newResolver = new DesktopAutomationDriver.Services.Resolution.ElementResolver(ctx, logger, GetWindowRoot);
        _pywinautoResolver = new DesktopAutomationDriver.Services.ElementResolution.ElementResolver(ctx, logger, GetWindowRoot);
    }

    // =========================================================================
    // Public dispatch entry point
    // =========================================================================

    /// <inheritdoc/>
    public object? Execute(UiRequest request, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.Operation))
            throw new ArgumentException("'operation' is required.");

        _logger.LogDebug("UI operation: {Operation}", SanitizeValue(request.Operation));

        var sw = Stopwatch.StartNew();

        try
        {
            return request.Operation.ToLowerInvariant() switch
            {
                // ----- Session & Window Management -----
                "launch"       => Launch(request),
                "close"        => Close(),
                "quit"         => Quit(request),
                "closewindow"  => CloseActiveWindow(request),
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
                "listtrackedwindows" => ListTrackedWindows(),
                "getcurrentroot" => GetCurrentRoot(request),
                "findlocator"    => FindLocatorDebug(request),
                "inspectlocator" => FindLocatorDebug(request),
                "findall"        => FindElements(request),
                "findmany"       => FindElements(request),
                "resolvemany"    => FindElements(request),
                "dumptree"       => DumpTree(request),
                "dump_tree"      => DumpTree(request),
                "inspecttree"    => DumpTree(request),

                // ----- Open dropdown item operations -----
                "listopendropdownitems"   => ListOpenDropdownItems(request),
                "selectopendropdownitem"  => SelectOpenDropdownItem(request),
                "clickopendropdownitem"   => SelectOpenDropdownItem(request),
                "listheaderdropdownitems"  => ListHeaderDropdownItems(request),
                "clickheaderdropdownitem"  => SelectHeaderDropdownItem(request),

                // ----- Element Query -----
                "exists"         => Exists(request),
                "waitfor"        => WaitFor(request, cancellationToken),
                "wait"           => Wait(request, cancellationToken),
                "isenabled"      => IsEnabled(request),
                "isvisible"      => IsVisible(request),
                "isfocused"      => IsFocused(request),
                "hasfocus"       => IsFocused(request),
                "iswindowactive" => IsWindowActive(request),
                "isactive"       => IsWindowActive(request),
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
                "scroll"           => Scroll(request, cancellationToken),
                "mousescroll"      => ScrollNormal(request, cancellationToken),
                "wheelscroll"      => ScrollNormal(request, cancellationToken),
                "scrollintoview"   => ScrollTargetIntoView(request, cancellationToken),
                "check"            => Check(request),
                "uncheck"          => Uncheck(request),
                "select"           => Select(request, cancellationToken),
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
                // Backward-compatible alias – the canonical operation is "select".
                "selectcomboboxitem" => Select(request, cancellationToken),
                "selectcomboboxuia" => SelectComboBoxNativeUia(request, cancellationToken),
                "findcomboboxuia" => FindComboBoxNativeUia(request, cancellationToken),
                "inspectcombobox" => InspectComboBox(request, cancellationToken),
                "draganddrop"     => DragAndDrop(request),
                "dragbyoffset"    => DragByOffset(request),
                "dragcoordinates" => DragCoordinates(request),
                "mouse"           => MouseAction(request),

                // ----- Popup Pipeline -----
                "topwindow"    => TopWindow(request),
                "waitforpopup" => WaitForPopup(request),
                "popupexists"  => PopupExists(request),
                "popupaction"  => PopupAction(request),

                // ----- Popup Text Reading -----
                "popuptext"  => PopupText(request),
                "alerttext"  => PopupText(request),
                "readpopup"  => PopupText(request),

                // ----- Alert / Dialog Handling (backward-compatible aliases) -----
                "alertok"     => PopupAction(request, defaultAction: "button", defaultButton: "OK|Yes|&OK|&Yes|Save"),
                "alertcancel" => PopupAction(request, defaultAction: "button", defaultButton: "Cancel|No|&Cancel|&No"),
                "alertclose"  => PopupAction(request, defaultAction: "close"),
                "popupok"     => PopupAction(request, defaultAction: "enter", requireTarget: true),

                // ----- Pywinauto-style Debug / List Operations -----
                "findelement"             => FindElement(request),
                "findelements"            => FindElements(request),
                "inspectelement"          => InspectElement(request),
                "printcontrolidentifiers" => PrintControlIdentifiers(request),
                "dumpcontrols"            => PrintControlIdentifiers(request),
                "resolve"                 => ResolveDebug(request),

                _ => throw new ArgumentException(
                    $"Unknown operation '{request.Operation}'. " +
                    "See GET /ui/operations for the full list.")
            };
        }
        finally
        {
            sw.Stop();

            _logger.LogInformation(
                "UI API operation completed. operation={Operation}, elapsedMs={ElapsedMs}, policy={Policy}, locator={Locator}",
                SanitizeValue(request.Operation),
                sw.ElapsedMilliseconds,
                GetOperationPolicy(request).PolicyName,
                request.Locator == null ? "" : DescribeLocator(request.Locator));
        }
    }

    // =========================================================================
    // Session & Window Management
    // =========================================================================

    private object? Launch(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("'value' must be the path to the executable for 'launch'.");

        var session = _ctx.Launch(req.Value);
        return new
        {
            sessionId          = session.SessionId,
            app                = req.Value,
            launchedProcessId  = session.LaunchedProcessId,
            trackedWindows     = ToTrackedWindowDtos(session.GetTrackedWindows())
        };
    }

    private object? Close()
    {
        var session = RequireSession();
        _logger.LogInformation("UI operation: close (graceful WM_CLOSE, no process kill).");

        var trackedBefore = session.GetTrackedWindows();
        _ctx.Close();

        var closedCount = trackedBefore.Count;

        _logger.LogInformation(
            "Close completed gracefully. closedCount={ClosedCount}",
            closedCount);

        return new
        {
            operation    = "close",
            forceKilled  = false,
            closedCount,
        };
    }

    private object? Quit(UiRequest req)
    {
        var session = RequireSession();
        _logger.LogInformation(
            "UI operation: quit. wasLaunchedByDriver={WasLaunched}, forceKillAttached={ForceKill}",
            session.WasLaunchedByDriver,
            req.ForceKillAttachedProcess);

        var wasLaunchedByDriver = session.WasLaunchedByDriver;
        var processId           = session.LaunchedProcessId ?? (int?)session.Application?.ProcessId;

        _ctx.Quit(req.ForceKillAttachedProcess);

        return new
        {
            operation           = "quit",
            wasLaunchedByDriver,
            processId,
            forceKillAttached   = req.ForceKillAttachedProcess,
        };
    }

    private object? ListTrackedWindows()
    {
        var session = _ctx.ActiveSession;
        var windows = _ctx.ListTrackedWindows();
        return new
        {
            launchedProcessId  = session?.LaunchedProcessId,
            wasLaunchedByDriver = session?.WasLaunchedByDriver ?? false,
            windows            = ToTrackedWindowDtos(windows)
        };
    }

    private object ToTrackedWindowDtos(IEnumerable<TrackedWindowInfo> windows)
    {
        return windows.Select(w => new
        {
            hwnd         = w.Hwnd.ToInt64(),
            hwndHex      = $"0x{w.Hwnd.ToInt64():X}",
            processId    = w.ProcessId,
            title        = w.Title,
            className    = w.ClassName,
            isMainWindow = w.IsMainWindow,
            firstSeenUtc = w.FirstSeenUtc,
            lastSeenUtc  = w.LastSeenUtc,
            exists       = IsWindow(w.Hwnd),
            visible      = IsWindow(w.Hwnd) && IsWindowVisible(w.Hwnd)
        }).ToList();
    }

    /// <summary>
    /// Closes the currently active window for the session (i.e. <c>session.ActiveWindow</c>
    /// or the application's main window when no explicit active window is set).
    /// When <c>req.Value</c> is non-empty, falls back to title-based window search for
    /// backward compatibility.
    /// </summary>
    private object? CloseActiveWindow(UiRequest req)
    {
        // If a title is given, delegate to the existing title-based close for backward compat.
        if (!string.IsNullOrWhiteSpace(req.Value))
            return CloseWindowByTitle(req);

        var session = _ctx.ActiveSession;
        if (session == null)
            throw new InvalidOperationException("No active session. Launch or attach first.");

        // Use the session's current active window, or fall back to the application's main window.
        var window = session.ActiveWindow;
        if (window == null)
        {
            var windows = session.Application.GetAllTopLevelWindows(session.Automation);
            window = windows.FirstOrDefault();
        }

        if (window == null)
        {
            _logger.LogWarning("closewindow: no active or top-level window found in session.");
            return new { operation = "closewindow", closed = false, reason = "no active or top-level window found" };
        }

        var hwnd      = SafeWindowHandle(window);
        var title     = SafeElementName(window);
        var className = SafeElementClassName(window);
        var processId = SafeProcessId(window);

        _logger.LogInformation(
            "closewindow: closing active window. title={Title}, hwnd=0x{Hwnd:X}",
            title,
            hwnd.ToInt64());

        var cf = session.Automation.ConditionFactory;
        CloseWindowElement(window, cf);

        return new
        {
            operation = "closewindow",
            hwnd      = hwnd.ToInt64(),
            hwndHex   = $"0x{hwnd.ToInt64():X}",
            processId,
            title,
            className,
            closed    = true
        };
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
        // At least one of value, hwnd, or className must be provided.
        if (string.IsNullOrWhiteSpace(req.Value) &&
            !req.Hwnd.HasValue &&
            string.IsNullOrWhiteSpace(req.ClassName))
        {
            throw new ArgumentException(
                "'switchwindow' requires at least one of: value (title), hwnd, or className.");
        }

        var session = _ctx.ActiveSession;

        // --- Win32-first fast path ---
        // Enumerate top-level HWNDs directly (no UIA tree walk) and activate the
        // matching window via Win32, which is significantly faster than the UIA search.
        var win32Match = FindWin32WindowForSwitch(req, session);

        if (win32Match != null)
        {
            _logger.LogInformation(
                "SwitchWindow Win32-first match found. hwnd=0x{Hwnd:X}, pid={Pid}, title={Title}, className={ClassName}",
                win32Match.Hwnd.ToInt64(),
                win32Match.ProcessId,
                win32Match.Title,
                win32Match.ClassName);

            if (ActivateWin32Window(win32Match.Hwnd))
            {
                // Ensure a session exists (attach by process ID when none is active).
                if (session == null)
                    session = _ctx.Attach(win32Match.ProcessId);

                // Wrap the HWND in a FlaUI element so the session has a typed ActiveWindow.
                var flauiWindow = ResolveFlaUIWindowFromHwnd(session, win32Match.Hwnd);

                if (flauiWindow != null)
                    return SwitchToWindow(session, flauiWindow);

                // Activation succeeded but FlaUI cannot wrap the handle (rare; e.g. cross-
                // process/elevated window). Return a partial result without setting ActiveWindow.
                _logger.LogInformation(
                    "SwitchWindow Win32-first activated window but FlaUI resolution failed; returning foreground-only result. hwnd=0x{Hwnd:X}",
                    win32Match.Hwnd.ToInt64());

                return new
                {
                    title     = win32Match.Title,
                    className = win32Match.ClassName,
                    hwnd      = win32Match.Hwnd.ToInt64(),
                    processId = win32Match.ProcessId,
                    strategy  = "win32-first-foreground-only"
                };
            }
        }

        // --- Fallback: existing UIA/FlaUI search logic ---
        _logger.LogInformation(
            "SwitchWindow Win32-first failed for value={Value}, hwnd={Hwnd}, className={ClassName}, processId={ProcessId}, matchMode={MatchMode}. Falling back to UIA/FlaUI.",
            SanitizeValue(req.Value),
            req.Hwnd.HasValue ? req.Hwnd.Value.ToString() : "(none)",
            SanitizeValue(req.ClassName),
            req.ProcessId.HasValue ? req.ProcessId.Value.ToString() : "(none)",
            SanitizeValue(req.MatchMode));

        // UIA/FlaUI fallback requires a title fragment.
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new InvalidOperationException(
                "No matching window was found. The Win32 search by hwnd/className yielded no result and no title value was provided for the UIA/FlaUI fallback.");

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
        {
            var pid = SafeProcessId(asWindow) ?? session.Application.ProcessId;
            session.TrackWindow(
                switchedHandle,
                pid,
                SafeElementName(asWindow),
                SafeElementClassName(asWindow),
                isMainWindow: false);
        }

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

    // =========================================================================
    // Win32-first SwitchWindow helpers
    // =========================================================================

    /// <summary>
    /// Enumerates visible top-level windows via Win32 <c>EnumWindows</c>.
    /// When <paramref name="processId"/> is provided, only windows owned by that
    /// process are returned.
    /// </summary>
    private List<Win32WindowInfo> EnumerateTopLevelWin32Windows(int? processId = null)
    {
        var result = new List<Win32WindowInfo>();
        var shellWindow = GetShellWindow();

        EnumWindows((hWnd, _) =>
        {
            if (hWnd == IntPtr.Zero || hWnd == shellWindow)
                return true;

            if (!IsWindowVisible(hWnd))
                return true;

            var titleLen = GetWindowTextLength(hWnd);
            string title = string.Empty;
            if (titleLen > 0)
            {
                var sb = new StringBuilder(titleLen + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                title = sb.ToString();
            }

            var className = GetWin32ClassName(hWnd);

            // Skip windows that have neither a visible title nor a meaningful class name
            // to avoid returning noise windows in broad all-process searches.
            if (string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(className))
                return true;

            GetWindowThreadProcessId(hWnd, out var pid);

            if (processId.HasValue && pid != (uint)processId.Value)
                return true;

            result.Add(new Win32WindowInfo
            {
                Hwnd = hWnd,
                ProcessId = (int)pid,
                Title = title,
                ClassName = className
            });

            return true;
        }, IntPtr.Zero);

        return result;
    }

    /// <summary>
    /// Searches visible top-level windows via Win32 for the best match using the
    /// extended criteria from <paramref name="req"/> (HWND, className, processId,
    /// matchMode, and/or title/value).
    /// </summary>
    private Win32WindowInfo? FindWin32WindowForSwitch(
        UiRequest req,
        AutomationSession? session)
    {
        var criteria = BuildWin32SwitchCriteria(req, session);

        // 1. Exact HWND wins unconditionally.
        if (criteria.Hwnd.HasValue && criteria.Hwnd.Value != IntPtr.Zero)
        {
            var hwnd = criteria.Hwnd.Value;

            if (IsWindow(hwnd))
            {
                GetWindowThreadProcessId(hwnd, out var pid);

                var info = new Win32WindowInfo
                {
                    Hwnd      = hwnd,
                    ProcessId = (int)pid,
                    Title     = GetWindowTitle(hwnd),
                    ClassName = GetWin32ClassName(hwnd)
                };

                _logger.LogInformation(
                    "SwitchWindow Win32 exact HWND match. hwnd=0x{Hwnd:X}, pid={Pid}, title={Title}, className={ClassName}",
                    hwnd.ToInt64(), info.ProcessId, info.Title, info.ClassName);

                return info;
            }

            _logger.LogWarning(
                "SwitchWindow requested HWND was not a valid window. hwnd=0x{Hwnd:X}",
                hwnd.ToInt64());
        }

        // 2. Process-filtered search first.
        if (criteria.ProcessId.HasValue)
        {
            var processMatch = FindBestWin32WindowMatch(
                EnumerateTopLevelWin32Windows(criteria.ProcessId),
                criteria);

            if (processMatch != null)
                return processMatch;
        }

        // 3. Widen to all processes if process-filtered search failed.
        return FindBestWin32WindowMatch(
            EnumerateTopLevelWin32Windows(null),
            criteria);
    }

    /// <summary>
    /// Builds the switch criteria from a <see cref="UiRequest"/>, resolving the
    /// process ID from the active session when the request does not supply one.
    /// </summary>
    private Win32SwitchCriteria BuildWin32SwitchCriteria(
        UiRequest request,
        AutomationSession? session)
    {
        int? processId = request.ProcessId;
        if (!processId.HasValue)
        {
            try { processId = session?.Application?.ProcessId; }
            catch { processId = null; }
        }

        IntPtr? hwnd = null;
        if (request.Hwnd.HasValue && request.Hwnd.Value != 0)
            hwnd = new IntPtr(request.Hwnd.Value);

        var matchMode = string.IsNullOrWhiteSpace(request.MatchMode)
            ? "contains"
            : request.MatchMode.Trim().ToLowerInvariant();

        if (matchMode is not ("exact" or "contains" or "regex"))
            matchMode = "contains";

        return new Win32SwitchCriteria
        {
            TitleOrValue = !string.IsNullOrWhiteSpace(request.Value) ? request.Value : null,
            Hwnd         = hwnd,
            ClassName    = string.IsNullOrWhiteSpace(request.ClassName) ? null : request.ClassName,
            ProcessId    = processId,
            MatchMode    = matchMode
        };
    }

    /// <summary>
    /// Scores each window in <paramref name="windows"/> against <paramref name="criteria"/>
    /// and returns the highest-scoring candidate, or null when no window qualifies.
    /// </summary>
    private Win32WindowInfo? FindBestWin32WindowMatch(
        IReadOnlyCollection<Win32WindowInfo> windows,
        Win32SwitchCriteria criteria)
    {
        Win32WindowInfo? best = null;
        int bestScore = 0;
        string bestReason = string.Empty;

        foreach (var window in windows)
        {
            var score = ScoreWin32WindowCandidate(window, criteria, out var reason);
            if (score > bestScore || (score == bestScore && best != null &&
                window.Title.Length < best.Title.Length))
            {
                best = window;
                bestScore = score;
                bestReason = reason;
            }
        }

        if (best == null)
            return null;

        _logger.LogInformation(
            "SwitchWindow Win32 best match. score={Score}, reason={Reason}, hwnd=0x{Hwnd:X}, pid={Pid}, title={Title}, className={ClassName}",
            bestScore, bestReason, best.Hwnd.ToInt64(), best.ProcessId, best.Title, best.ClassName);

        return best;
    }

    /// <summary>
    /// Returns a positive match score for <paramref name="window"/> against the
    /// given criteria, or 0 when the window does not qualify.
    /// </summary>
    private static int ScoreWin32WindowCandidate(
        Win32WindowInfo window,
        Win32SwitchCriteria criteria,
        out string reason)
    {
        reason = string.Empty;
        var score = 0;

        if (criteria.ProcessId.HasValue && window.ProcessId == criteria.ProcessId.Value)
        {
            score += 50;
            reason += "pid;";
        }

        if (!string.IsNullOrWhiteSpace(criteria.ClassName))
        {
            if (!ClassNameMatches(window.ClassName, criteria.ClassName))
                return 0;

            score += string.Equals(window.ClassName, criteria.ClassName,
                StringComparison.OrdinalIgnoreCase) ? 50 : 25;
            reason += "class;";
        }

        if (!string.IsNullOrWhiteSpace(criteria.TitleOrValue))
        {
            if (!TitleMatches(window.Title, criteria.TitleOrValue, criteria.MatchMode))
                return 0;

            score += criteria.MatchMode switch
            {
                "exact" => 80,
                "regex" => 60,
                _ => string.Equals(window.Title, criteria.TitleOrValue,
                         StringComparison.OrdinalIgnoreCase) ? 80 : 40
            };
            reason += $"title-{criteria.MatchMode};";
        }

        // Require at least one of title or className when neither was provided; don't random-match.
        if (string.IsNullOrWhiteSpace(criteria.TitleOrValue) &&
            string.IsNullOrWhiteSpace(criteria.ClassName))
        {
            return 0;
        }

        return score;
    }

    /// <summary>Returns true when <paramref name="actualTitle"/> satisfies the match.</summary>
    private static bool TitleMatches(string actualTitle, string? expected, string matchMode)
    {
        if (string.IsNullOrWhiteSpace(expected))
            return true;

        actualTitle ??= string.Empty;

        return matchMode switch
        {
            "exact" => string.Equals(actualTitle, expected, StringComparison.OrdinalIgnoreCase),
            "regex" => TryRegexMatch(actualTitle, expected),
            _ => actualTitle.Contains(expected, StringComparison.OrdinalIgnoreCase)
        };
    }

    /// <summary>
    /// Attempts to match <paramref name="input"/> against the user-supplied
    /// <paramref name="pattern"/> using a case-insensitive regex with a short timeout
    /// to guard against catastrophic backtracking.  Returns false when the pattern is
    /// invalid or the match times out.
    /// </summary>
    private static bool TryRegexMatch(string input, string pattern)
    {
        try
        {
            return Regex.IsMatch(input, pattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                matchTimeout: TimeSpan.FromSeconds(1));
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            // Invalid regex pattern; treat as no match rather than throwing.
            return false;
        }
    }

    /// <summary>
    /// Returns true when <paramref name="actualClassName"/> equals or contains
    /// <paramref name="expectedClassName"/> (case-insensitive).
    /// </summary>
    private static bool ClassNameMatches(string actualClassName, string? expectedClassName)
    {
        if (string.IsNullOrWhiteSpace(expectedClassName))
            return true;

        actualClassName ??= string.Empty;

        return string.Equals(actualClassName, expectedClassName, StringComparison.OrdinalIgnoreCase)
            || actualClassName.Contains(expectedClassName, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Reads the Win32 class name of a window handle.</summary>
    private static string GetWin32ClassName(IntPtr hwnd)
    {
        try
        {
            var sb = new StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>Reads the window title for an HWND without going through UIA.</summary>
    private static string GetWindowTitle(IntPtr hwnd)
    {
        try
        {
            var len = GetWindowTextLength(hwnd);
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Restores a minimized window (if needed) and brings it to the foreground
    /// using Win32 APIs. Returns <c>true</c> when <c>SetForegroundWindow</c> reports
    /// success.
    /// </summary>
    private static bool ActivateWin32Window(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return false;

        try
        {
            if (IsIconic(hwnd))
                ShowWindow(hwnd, SW_RESTORE);

            return SetForegroundWindow(hwnd);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Converts a raw Win32 window handle back to a FlaUI <see cref="Window"/>
    /// using <c>AutomationBase.FromHandle</c>.
    /// Returns null when the conversion fails (e.g. the window is no longer valid).
    /// </summary>
    private static Window? ResolveFlaUIWindowFromHwnd(AutomationSession session, IntPtr hwnd)
    {
        try
        {
            var element = session.Automation.FromHandle(hwnd);
            return element?.AsWindow();
        }
        catch
        {
            return null;
        }
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
            visible = SafeIsOffscreen(e) is false,
            runtimeId = SafeRuntimeIdString(e),
            frameworkId = SafeFrameworkId(e),
            value = SafeElementValue(e),
            text = SafeElementText(e)
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
        var session = TryGetSessionOrNull();
        var seenHandles = new HashSet<long>();
        var windows = new List<AutomationElement>();

        // Collect app-level windows first (when a session is available).
        if (session != null)
        {
            try
            {
                foreach (var w in session.Application.GetAllTopLevelWindows(session.Automation))
                {
                    var h = SafeWindowHandle(w).ToInt64();
                    if (h != 0 && seenHandles.Add(h))
                        windows.Add(w);
                }
            }
            catch { /* ignore unstable enumerations */ }
        }

        // Desktop children (always; this covers other processes and session-less calls).
        var automation = (AutomationBase?)session?.Automation;
        UIA3Automation? tempAuto = null;

        if (automation == null)
        {
            tempAuto = new UIA3Automation();
            automation = tempAuto;
        }

        try
        {
            var cf = automation.ConditionFactory;
            var desktopChildren = automation.GetDesktop().FindAllChildren(cf.ByControlType(ControlType.Window));
            foreach (var w in desktopChildren)
            {
                var h = SafeWindowHandle(w).ToInt64();
                if (h != 0 && seenHandles.Add(h))
                    windows.Add(w);
            }

            if (req.IncludeDesktopDescendants == true)
            {
                var descendants = automation.GetDesktop().FindAllDescendants(cf.ByControlType(ControlType.Window));
                foreach (var w in descendants)
                {
                    var h = SafeWindowHandle(w).ToInt64();
                    if (h != 0 && seenHandles.Add(h))
                        windows.Add(w);
                }
            }
        }
        catch { /* ignore */ }
        finally
        {
            tempAuto?.Dispose();
        }

        IEnumerable<AutomationElement> filtered = windows;
        if (!string.IsNullOrWhiteSpace(req.Value))
            filtered = windows.Where(w => TitleContains(w, req.Value));

        var fg = GetForegroundWindow();

        return filtered.Select(w =>
        {
            var hwndPtr = SafeWindowHandle(w);
            var hwndVal = hwndPtr.ToInt64();
            return new
            {
                title        = SafeElementName(w),
                automationId = SafeElementAutomationId(w),
                className    = SafeElementClassName(w),
                processId    = SafeProcessId(w),
                hwnd         = hwndVal,
                hwndHex      = $"0x{hwndVal:X}",
                visible      = hwndPtr != IntPtr.Zero && IsWindowVisible(hwndPtr),
                foreground   = hwndPtr == fg,
                isOffscreen  = SafeIsOffscreen(w)
            };
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

    private object Exists(UiRequest req)
    {
        var result = PywinautoResolve(req, QueryResolveOptions);
        if (result.Success && result.Element != null)
        {
            var element = result.Element;
            return new
            {
                operation = "exists",
                exists = true,
                strategy = result.Strategy,
                automationId = SafeElementAutomationId(element),
                name = SafeElementName(element),
                controlType = SafeElementControlType(element),
                className = SafeElementClassName(element),
                isOffscreen = element.Properties.IsOffscreen.ValueOrDefault,
                boundingRectangle = SafeBoundingRectangleObject(element),
                snapshot = result.Snapshot,
                candidateCount = result.CandidateCount,
                searchRoot = result.SearchRoot,
                criteria = result.Criteria,
                elapsedMs = result.ElapsedMs,
                fallbackUsed = result.FallbackUsed
            };
        }

        return BuildPywinautoQueryResponse("exists", result);
    }

    private object WaitFor(UiRequest req, CancellationToken cancellationToken)
    {
        // Backward compatibility: old waitfor commonly checked IsEnabled + !IsOffscreen,
        // so default to "enabled" when no explicit state is requested.
        var state = string.IsNullOrWhiteSpace(req.State) ? "enabled" : req.State;

        // Backward compatibility: older callers pass the timeout in milliseconds via
        // request.Value (e.g. "10000").
        int? timeoutMs = req.TimeoutMs;
        if (!timeoutMs.HasValue &&
            !string.IsNullOrWhiteSpace(req.Value) &&
            int.TryParse(req.Value, out var parsedTimeoutMs))
        {
            timeoutMs = parsedTimeoutMs;
        }

        var normalized = new UiRequest
        {
            Operation          = "wait",
            Locator            = req.Locator,
            ParentLocator      = req.ParentLocator,
            State              = state,
            TimeoutMs          = timeoutMs,
            PollIntervalMs     = req.PollIntervalMs ?? 200,
            Fast               = req.Fast,
            DisableAutoFollow  = req.DisableAutoFollow,
            UseCache           = req.UseCache,
            PreferXPath        = req.PreferXPath,
            XPathOnly          = req.XPathOnly,
            PreferAttributes   = req.PreferAttributes,
            FallbackToWindowRootIfParentChildNotFound = req.FallbackToWindowRootIfParentChildNotFound
        };

        return Wait(normalized, cancellationToken);
    }

    private object IsEnabled(UiRequest req)
    {
        var resolved   = ResolveForStateQuery(req);
        var element    = resolved.Element;
        var uiaEnabled = element.IsEnabled;

        bool? win32Enabled = null;
        var hwnd = SafeWindowHandle(element);
        if (hwnd != IntPtr.Zero)
            win32Enabled = IsWindowEnabled(hwnd);

        return new
        {
            operation    = "isenabled",
            exists       = true,
            enabled      = uiaEnabled,
            uiaEnabled,
            win32Enabled,
            strategy     = resolved.Strategy,
            automationId = SafeElementAutomationId(element),
            name         = SafeElementName(element),
            controlType  = SafeElementControlType(element),
            className    = SafeElementClassName(element)
        };
    }

    private object IsVisible(UiRequest req)
    {
        var resolved  = ResolveForStateQuery(req);
        var element   = resolved.Element;
        var rect      = element.BoundingRectangle;
        var isOffscreen = element.Properties.IsOffscreen.ValueOrDefault;
        var hasRect   = !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
        var visible   = hasRect && !isOffscreen;

        return new
        {
            operation   = "isvisible",
            exists      = true,
            visible,
            isOffscreen,
            hasRect,
            rect = new
            {
                left   = rect.Left,
                top    = rect.Top,
                right  = rect.Right,
                bottom = rect.Bottom,
                width  = rect.Width,
                height = rect.Height
            },
            strategy = resolved.Strategy
        };
    }

    private object IsFocused(UiRequest req)
    {
        var resolved = ResolveForStateQuery(req);
        var element  = resolved.Element;
        var focused  = element.Properties.HasKeyboardFocus.ValueOrDefault;

        return new
        {
            operation = "isfocused",
            exists    = true,
            focused,
            strategy  = resolved.Strategy
        };
    }

    private object IsWindowActive(UiRequest req)
    {
        var resolved   = ResolveForStateQuery(req);
        var element    = resolved.Element;
        var hwnd       = SafeWindowHandle(element);
        var rootHwnd   = hwnd != IntPtr.Zero ? GetAncestor(hwnd, GA_ROOT) : IntPtr.Zero;
        var foreground = GetForegroundWindow();
        var active     = rootHwnd != IntPtr.Zero && rootHwnd == foreground;

        return new
        {
            operation      = "iswindowactive",
            exists         = true,
            active,
            hwnd           = rootHwnd.ToInt64(),
            foregroundHwnd = foreground.ToInt64(),
            strategy       = resolved.Strategy
        };
    }

    private object IsClickable(UiRequest req)
    {
        var resolved   = ResolveForStateQuery(req);
        var element    = resolved.Element;
        var enabled    = element.IsEnabled;
        var visible    = IsTargetPracticallyVisible(element, null, out var visibilityStrategy);
        var hasClickablePoint = TryGetElementClickablePoint(element, out var pointX, out var pointY);
        var clickable  = enabled && visible && hasClickablePoint;

        return new
        {
            operation         = "isclickable",
            exists            = true,
            clickable,
            enabled,
            visible,
            visibilityStrategy,
            hasClickablePoint,
            point = hasClickablePoint
                ? (object)new { x = (int)pointX, y = (int)pointY }
                : null,
            strategy = resolved.Strategy
        };
    }

    private object IsEditable(UiRequest req)
    {
        var resolved     = ResolveForStateQuery(req);
        var element      = resolved.Element;
        var enabled      = element.IsEnabled;
        var valuePattern = element.Patterns.Value.PatternOrDefault;
        var hasValuePattern = valuePattern != null;
        var isReadOnly   = valuePattern?.IsReadOnly ?? true;
        var editable     = enabled && hasValuePattern && !isReadOnly;

        return new
        {
            operation       = "iseditable",
            exists          = true,
            editable,
            enabled,
            hasValuePattern,
            isReadOnly,
            strategy        = resolved.Strategy
        };
    }

    private object? IsChecked(UiRequest req)
    {
        var element = ResolveElementForOperation(req, "ischecked");

        if (element.Patterns.Toggle.IsSupported)
            return new { @checked = element.Patterns.Toggle.Pattern.ToggleState == ToggleState.On };

        if (element.Patterns.SelectionItem.IsSupported)
            return new { @checked = element.Patterns.SelectionItem.Pattern.IsSelected };

        return new { @checked = false };
    }

    private object? GetValue(UiRequest req)
    {
        var resolved = PywinautoResolve(req, ReadResolveOptions);
        if (!resolved.Success || resolved.Element == null)
            throw new InvalidOperationException(
                $"Element not found for getvalue. criteria={System.Text.Json.JsonSerializer.Serialize(resolved.Criteria)}, candidates={resolved.CandidateCount}");

        var element = resolved.Element;

        if (element.Patterns.Value.IsSupported)
            return new { value = element.Patterns.Value.Pattern.Value ?? string.Empty, resolverStrategy = resolved.Strategy };

        if (element.Patterns.RangeValue.IsSupported)
            return new { value = element.Patterns.RangeValue.Pattern.Value.ToString(), resolverStrategy = resolved.Strategy };

        return new { value = element.Name ?? string.Empty, resolverStrategy = resolved.Strategy };
    }

    private object? GetText(UiRequest req)
    {
        var resolved = PywinautoResolve(req, ReadResolveOptions);
        if (!resolved.Success || resolved.Element == null)
            throw new InvalidOperationException(
                $"Element not found for gettext. criteria={System.Text.Json.JsonSerializer.Serialize(resolved.Criteria)}, candidates={resolved.CandidateCount}");

        var element = resolved.Element;
        const int MaxText = 1_048_576;

        try
        {
            if (element.Patterns.Text.IsSupported)
                return new { text = element.Patterns.Text.Pattern.DocumentRange.GetText(MaxText), resolverStrategy = resolved.Strategy };
        }
        catch { /* fall through */ }

        try
        {
            if (element.Patterns.Value.IsSupported)
                return new { text = element.Patterns.Value.Pattern.Value ?? string.Empty, resolverStrategy = resolved.Strategy };
        }
        catch { /* fall through */ }

        return new { text = element.Name ?? string.Empty, resolverStrategy = resolved.Strategy };
    }

    private object? GetName(UiRequest req)
    {
        var resolved = PywinautoResolve(req, ReadResolveOptions);
        if (!resolved.Success || resolved.Element == null)
            throw new InvalidOperationException(
                $"Element not found for getname. criteria={System.Text.Json.JsonSerializer.Serialize(resolved.Criteria)}, candidates={resolved.CandidateCount}");

        return new { name = resolved.Element.Name ?? string.Empty, resolverStrategy = resolved.Strategy };
    }

    private object? GetControlType(UiRequest req)
    {
        var resolved = PywinautoResolve(req, ReadResolveOptions);
        if (!resolved.Success || resolved.Element == null)
            throw new InvalidOperationException(
                $"Element not found for getcontroltype. criteria={System.Text.Json.JsonSerializer.Serialize(resolved.Criteria)}, candidates={resolved.CandidateCount}");

        return new { controlType = resolved.Element.ControlType.ToString(), resolverStrategy = resolved.Strategy };
    }

    private object? GetSelected(UiRequest req)
    {
        var element = ResolveElementForOperation(req, "getselected");

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
        var element = ResolveElementForOperation(
            req,
            purpose: "gettable",
            action: false,
            allowOffscreen: true,
            requireClickable: false);
        return ReadTableData(element, headersOnly: false, session);
    }

    private object? GetTableHeaders(UiRequest req)
    {
        var session = RequireSession();
        var element = ResolveElementForOperation(
            req,
            purpose: "gettableheaders",
            action: false,
            allowOffscreen: true,
            requireClickable: false);
        return ReadTableData(element, headersOnly: true, session);
    }

    // =========================================================================
    // Position Comparison
    // =========================================================================

    private object? IsRightOf(UiRequest req)
    {
        var (r1, r2, e1, e2) = GetTwoRectsLive(req);

        var strict      = r1.Left >= r2.Right;
        var centerBased = CenterX(r1) > CenterX(r2);
        var sameRow     = RectanglesOverlapVertically(r1, r2);

        return new
        {
            operation        = "isrightof",
            isRightOf        = strict,
            strict,
            centerBased,
            sameRow,
            element1         = RectObj(r1),
            element2         = RectObj(r2),
            element1Strategy = e1.Strategy,
            element2Strategy = e2.Strategy,
            element1Info     = ElementInfoObj(e1.Element),
            element2Info     = ElementInfoObj(e2.Element)
        };
    }

    private object? IsLeftOf(UiRequest req)
    {
        var (r1, r2, e1, e2) = GetTwoRectsLive(req);

        var strict      = r1.Right <= r2.Left;
        var centerBased = CenterX(r1) < CenterX(r2);
        var sameRow     = RectanglesOverlapVertically(r1, r2);

        return new
        {
            operation        = "isleftof",
            isLeftOf         = strict,
            strict,
            centerBased,
            sameRow,
            element1         = RectObj(r1),
            element2         = RectObj(r2),
            element1Strategy = e1.Strategy,
            element2Strategy = e2.Strategy,
            element1Info     = ElementInfoObj(e1.Element),
            element2Info     = ElementInfoObj(e2.Element)
        };
    }

    private object? IsAbove(UiRequest req)
    {
        var (r1, r2, e1, e2) = GetTwoRectsLive(req);

        var strict      = r1.Bottom <= r2.Top;
        var centerBased = CenterY(r1) < CenterY(r2);
        var sameColumn  = RectanglesOverlapHorizontally(r1, r2);

        return new
        {
            operation        = "isabove",
            isAbove          = strict,
            strict,
            centerBased,
            sameColumn,
            element1         = RectObj(r1),
            element2         = RectObj(r2),
            element1Strategy = e1.Strategy,
            element2Strategy = e2.Strategy,
            element1Info     = ElementInfoObj(e1.Element),
            element2Info     = ElementInfoObj(e2.Element)
        };
    }

    private object? IsBelow(UiRequest req)
    {
        var (r1, r2, e1, e2) = GetTwoRectsLive(req);

        var strict      = r1.Top >= r2.Bottom;
        var centerBased = CenterY(r1) > CenterY(r2);
        var sameColumn  = RectanglesOverlapHorizontally(r1, r2);

        return new
        {
            operation        = "isbelow",
            isBelow          = strict,
            strict,
            centerBased,
            sameColumn,
            element1         = RectObj(r1),
            element2         = RectObj(r2),
            element1Strategy = e1.Strategy,
            element2Strategy = e2.Strategy,
            element1Info     = ElementInfoObj(e1.Element),
            element2Info     = ElementInfoObj(e2.Element)
        };
    }

    private object? GetPosition(UiRequest req)
    {
        var (r1, r2, e1, e2) = GetTwoRectsLive(req);

        var cx1 = CenterX(r1);
        var cy1 = CenterY(r1);
        var cx2 = CenterX(r2);
        var cy2 = CenterY(r2);

        return new
        {
            operation             = "getposition",
            element1              = RectObj(r1),
            element2              = RectObj(r2),
            isRightOf             = r1.Left   >= r2.Right,
            isLeftOf              = r1.Right  <= r2.Left,
            isAbove               = r1.Bottom <= r2.Top,
            isBelow               = r1.Top    >= r2.Bottom,
            center1               = new { x = cx1, y = cy1 },
            center2               = new { x = cx2, y = cy2 },
            center1RightOfCenter2 = cx1 > cx2,
            center1LeftOfCenter2  = cx1 < cx2,
            center1AboveCenter2   = cy1 < cy2,
            center1BelowCenter2   = cy1 > cy2,
            sameRow               = RectanglesOverlapVertically(r1, r2),
            sameColumn            = RectanglesOverlapHorizontally(r1, r2),
            element1Strategy      = e1.Strategy,
            element2Strategy      = e2.Strategy,
            element1Info          = ElementInfoObj(e1.Element),
            element2Info          = ElementInfoObj(e2.Element)
        };
    }

    // =========================================================================
    // Element Actions
    // =========================================================================

    private object? Click(UiRequest req)
    {
        var resolved = PywinautoResolve(req, ClickResolveOptions);
        if (!resolved.Success || resolved.Element == null)
        {
            throw new InvalidOperationException(
                $"Element not found for click. criteria={System.Text.Json.JsonSerializer.Serialize(resolved.Criteria)}, candidates={System.Text.Json.JsonSerializer.Serialize(resolved.Candidates)}");
        }

        var element = resolved.Element;

        BringElementWindowToForeground(element);

        string clickStrategy = "physical-click";
        bool clicked = false;

        if (string.Equals(req.Mode, "invoke", StringComparison.OrdinalIgnoreCase) && 
            (element.ControlType == ControlType.Button || element.ControlType == ControlType.MenuItem))
        {
            if (TryInvokePattern(element, "Click"))
            {
                clicked = true;
                clickStrategy = "uia-invoke";
            }
        }

        if (!clicked)
        {
            var point = element.GetClickablePoint();
            if (point.IsEmpty)
            {
                clickStrategy = "physical-click-center";
            }
            else
            {
                clickStrategy = "physical-click-clickable-point";
            }

            clicked = TryPhysicalClick(element, "Click") ||
                      TryElementClick(element, "Click") ||
                      TryInvokePattern(element, "Click");
        }

        if (!clicked)
        {
            throw new InvalidOperationException(
                $"Click failed after trying physical click, FlaUI click, and InvokePattern for " +
                $"name='{SafeElementName(element)}' controlType={element.ControlType}");
        }

        return new
        {
            operation = "click",
            clicked = true,
            success = true,
            strategy = clickStrategy,
            resolverStrategy = resolved.Strategy,
            element = resolved.Snapshot ?? CreateElementSnapshot(element),
            fallbackUsed = resolved.FallbackUsed
        };
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
        var locator = req.Locator;
        if (locator == null && req.LocatorPath == null && req.Criteria == null)
            throw new ArgumentException("'locator', 'locatorPath' or 'criteria' is required for findlocator/inspectlocator.");

        RequireSession();

        var result = PywinautoResolve(req, QueryResolveOptions);
        return BuildPywinautoQueryResponse(req.Operation ?? "findlocator", result);
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
        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "doubleclick",
            action: true,
            allowOffscreen: false,
            requireClickable: true);

        var element = resolved.Element;
        element.DoubleClick();

        return new
        {
            doubleclicked = true,
            strategy = resolved.Strategy,
            target = CreateElementSnapshot(element)
        };
    }

    private object? RightClick(UiRequest req)
    {
        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "rightclick",
            action: true,
            allowOffscreen: false,
            requireClickable: true);

        var element = resolved.Element;

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        try
        {
            element.RightClick();

            return new
            {
                rightclicked = true,
                strategy = resolved.Strategy,
                target = CreateElementSnapshot(element)
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
                rightclicked = true,
                strategy = "physical-right-click",
                target = CreateElementSnapshot(element)
            };
        }

        throw new InvalidOperationException($"Failed to right-click element '{SafeElementName(element)}'.");
    }

    private object? ContextMenuPath(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("value is required for contextmenupath.");

        var element = ResolveElementForOperation(
            req,
            purpose: "contextmenupath",
            action: true,
            allowOffscreen: false,
            requireClickable: false);
        var rawValue = System.Net.WebUtility.HtmlDecode(req.Value).Trim();
        var pathParts = rawValue
            .Split('>', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (pathParts.Count == 0)
            throw new ArgumentException("contextmenupath requires at least one menu item in value.");

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        // Capture the right-click point and target HWND before the right-click so we can
        // scope context-menu popup detection to only the relevant popup.
        var rect = element.BoundingRectangle;
        var rightClickPoint = new Point(
            (int)Math.Round(rect.Left + rect.Width / 2.0),
            (int)Math.Round(rect.Top + rect.Height / 2.0));

        IntPtr targetHwnd = IntPtr.Zero;
        try
        {
            var hwnd = SafeWindowHandle(element);
            if (hwnd != IntPtr.Zero)
            {
                var root = GetAncestor(hwnd, GA_ROOT);
                targetHwnd = root != IntPtr.Zero ? root : hwnd;
            }
        }
        catch { /* best effort */ }

        if (!TryInstantPhysicalRightClick(element, "ContextMenuPath RightClick"))
            throw new InvalidOperationException($"Failed to right-click element '{SafeElementName(element)}'.");

        Thread.Sleep(MenuExpandDelayMs);

        var session = RequireSession();
        var currentRoot = FindActiveContextMenuPopup(session, rightClickPoint, targetHwnd)
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

                _logger.LogInformation(
                    "ContextMenuPath selected item. value={Value}, activation=physical-click-first. Returning immediately; caller should wait/switch to new window if needed.",
                    rawValue);

                return new
                {
                    selected = rawValue,
                    strategy = "context-menu-path",
                    activation = "physical-click-first",
                    target = CreateElementSnapshot(element)
                };
            }

            if (!OpenContextSubMenu(item, part))
                throw new InvalidOperationException($"Failed to open context submenu '{part}' for path '{rawValue}'.");

            Thread.Sleep(MenuExpandDelayMs);

            currentRoot = FindContextSubMenuPopup(session, item, rightClickPoint, targetHwnd)
                ?? FindActiveContextMenuPopup(session, rightClickPoint, targetHwnd);

            if (currentRoot == null)
                throw new InvalidOperationException($"Context submenu popup was not found after opening '{part}'.");
        }

        throw new InvalidOperationException($"Context menu path '{rawValue}' was not completed.");
    }

    private object? Hover(UiRequest req)
    {
        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "hover",
            action: true,
            allowOffscreen: false,
            requireClickable: true);

        var element = resolved.Element;
        var pt = element.GetClickablePoint();
        Mouse.MoveTo(pt);

        return new
        {
            hovered = true,
            strategy = resolved.Strategy,
            target = CreateElementSnapshot(element)
        };
    }

    private object? Focus(UiRequest req)
    {
        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "focus",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        var element = resolved.Element;
        element.Focus();

        return new
        {
            focused = true,
            strategy = resolved.Strategy,
            target = CreateElementSnapshot(element)
        };
    }

    private object? TypeText(UiRequest req)
    {
        if (req.Value == null)
            throw new ArgumentException("'value' is required for 'type'.");

        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "type",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        var element = resolved.Element;

        if (WinFormsDateTimePickerHelper.IsDateTimePicker(element))
            return TypeDate(element, req.Value);

        if (!FocusElementForKeyboardInput(element, "TypeText"))
        {
            throw new InvalidOperationException(
                $"Keyboard focus could not be confirmed on target before typing. target='{SafeElementName(element)}'");
        }

        Thread.Sleep(KeyboardInputReadyDelayMs);

        Keyboard.Type(req.Value);

        return new
        {
            typed = true,
            strategy = resolved.Strategy,
            target = CreateElementSnapshot(element)
        };
    }

    private object? TypeDate(UiRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Value))
            throw new ArgumentException("Date value is required.");

        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "typedate",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        var element = resolved.Element;

        if (!WinFormsDateTimePickerHelper.IsDateTimePicker(element))
        {
            _logger.LogWarning(
                "typedate called on non-DateTimePicker. name={Name}, className={ClassName}, controlType={ControlType}",
                SafeElementName(element),
                SafeElementClassName(element),
                element.ControlType);
        }

        var result = TypeDate(element, req.Value);

        return new
        {
            typed = true,
            strategy = resolved.Strategy,
            target = CreateElementSnapshot(element)
        };
    }

    private object? TypeDate(AutomationElement element, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Date value is required.");

        var format = WinFormsDateTimePickerHelper.DetectDateFormat(element);

        if (!WinFormsDateTimePickerHelper.TryParseDateParts(value, format, out var first, out var second, out var third, out _))
            throw new ArgumentException($"Invalid date value. Use {format.DisplayFormat}.");

        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        ClickDatePickerMonthSection(element);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerClickDelayMs);

        SendDatePickerKey(VirtualKeyShort.HOME);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(first);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        SendDatePickerKey(VirtualKeyShort.RIGHT);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(second);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        SendDatePickerKey(VirtualKeyShort.RIGHT);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        Keyboard.Type(third);
        Thread.Sleep(WinFormsDateTimePickerHelper.DatePickerSegmentDelayMs);

        SendDatePickerKey(VirtualKeyShort.RETURN);

        return new
        {
            typed = true,
            strategy = "date-segments",
            format = format.DisplayFormat,
            formatSource = format.Source,
            value = $"{first}{format.Separator}{second}{format.Separator}{third}",
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
        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "clear",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        var element = resolved.Element;

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
                        strategy = resolved.Strategy,
                        target = CreateElementSnapshot(element)
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
                target = CreateElementSnapshot(element)
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
                target = CreateElementSnapshot(element)
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
                target = CreateElementSnapshot(element)
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
        string strategy = "active-window";

        if (req.Locator != null && !IsEmptyLocator(req.Locator))
        {
            var resolved = ResolveElementResultForOperation(
                req,
                purpose: "sendkeys",
                action: true,
                allowOffscreen: false,
                requireClickable: false);

            element = resolved.Element;
            strategy = resolved.Strategy;

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
            strategy = strategy,
            target = element == null ? null : CreateElementSnapshot(element)
        };
    }

    private object? ExpandTreeItem(UiRequest req)
    {
        var item = ResolveElementForOperation(
            req,
            purpose: "expandtreeitem",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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
        var item = ResolveElementForOperation(
            req,
            purpose: "collapsetreeitem",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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
        var item = ResolveElementForOperation(
            req,
            purpose: "selecttreeitem",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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

        var root = ResolveElementForOperation(
            req,
            purpose: "selecttreepath",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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

        var root = ResolveElementForOperation(
            req,
            purpose: "expandtreepath",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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

    /// <summary>
    /// Smart scroll dispatcher. Infers the right scroll behavior from request shape:
    /// Case A: locator + direction/amount → normal container scroll.
    /// Case B: locator only               → scroll target into view.
    /// Case C: containerLocator + locator → scroll target into view inside container.
    /// Case D: x/y only                  → coordinate-based wheel scroll.
    /// </summary>
    private object Scroll(UiRequest request, CancellationToken cancellationToken)
    {
        var hasContainer        = request.ContainerLocator != null && !IsEmptyLocator(request.ContainerLocator);
        var hasLocator          = request.Locator != null && !IsEmptyLocator(request.Locator);
        var hasDirectionOrAmount =
            !string.IsNullOrWhiteSpace(request.Direction) ||
            request.Amount.HasValue;
        var hasCoordinates = request.X.HasValue && request.Y.HasValue;

        // Case D: coordinate-based wheel scroll
        if (!hasLocator && !hasContainer && hasCoordinates)
            return ScrollCoordinates(request, cancellationToken);

        // Case C: containerLocator + locator → target inside container
        if (hasContainer && hasLocator)
            return ScrollTargetIntoView(request, cancellationToken);

        // Case A: locator + direction/amount → normal container scroll
        if (hasLocator && hasDirectionOrAmount)
            return ScrollNormal(request, cancellationToken);

        // Case B: locator only → scroll target into view
        if (hasLocator)
            return ScrollTargetIntoView(request, cancellationToken);

        throw new ArgumentException(
            "Invalid scroll request. Provide: locator (target), locator+direction/amount (container scroll), containerLocator+locator (target inside container), or x/y (coordinate scroll).");
    }

    /// <summary>
    /// Normal container scroll. Supports mode=auto|wheel|pattern, direction, and amount.
    /// Used directly by mousescroll/wheelscroll aliases (forces mode=wheel) and by
    /// Case A of the smart Scroll dispatcher.
    /// </summary>
    private object ScrollNormal(UiRequest request, CancellationToken cancellationToken)
    {
        var mode      = NormalizeScrollMode(request);
        var direction = NormalizeScrollDirection(request.Direction);
        var amount    = request.Amount ?? 1;

        if (amount == 0)
            throw new ArgumentException("'scroll' requires non-zero amount.");

        if (amount < 0)
        {
            amount    = Math.Abs(amount);
            direction = ReverseScrollDirection(direction);
        }

        AutomationElement? element = null;

        if (request.Locator != null && !IsEmptyLocator(request.Locator))
        {
            var resolved = ResolveElementResultForOperation(
                request,
                purpose: "scroll",
                action: true,
                allowOffscreen: false,
                requireClickable: false);
            element = resolved.Element;
            BringElementWindowToForeground(element);
            Thread.Sleep(WindowActivationDelayMs);
        }

        var beforeState = request.VerifyScroll == true && element != null
            ? CaptureScrollState(element)
            : null;

        var strategy = string.Empty;
        var success  = false;

        if (mode == "pattern" || mode == "auto")
        {
            if (element != null && TryScrollByPattern(element, direction, amount, out strategy))
                success = true;

            if (mode == "pattern" && !success)
                throw new InvalidOperationException(
                    $"ScrollPattern failed or is not supported. direction={direction}, amount={amount}, element={SafeElementName(element!)}");
        }

        if (!success && (mode == "wheel" || mode == "auto"))
        {
            if (element == null)
            {
                if (!request.X.HasValue || !request.Y.HasValue)
                    throw new ArgumentException(
                        "Wheel scroll requires either a locator or x/y coordinates.");

                cancellationToken.ThrowIfCancellationRequested();
                success = TryMouseWheelScrollAtPoint(
                    new Point(request.X.Value, request.Y.Value),
                    direction,
                    amount,
                    cancellationToken,
                    out strategy);
            }
            else
            {
                cancellationToken.ThrowIfCancellationRequested();
                success = TryMouseWheelScrollOnElement(element, direction, amount, cancellationToken, out strategy);
            }
        }

        if (!success)
            throw new InvalidOperationException(
                $"Unable to scroll. mode={mode}, direction={direction}, amount={amount}");

        Thread.Sleep(request.ScrollDelayMs ?? 100);

        var verified = false;

        if (request.VerifyScroll == true && element != null)
        {
            var afterState = CaptureScrollState(element);
            verified = !string.Equals(beforeState, afterState, StringComparison.Ordinal);
        }

        return new
        {
            operation   = "scroll",
            success     = true,
            strategy,
            mode,
            direction,
            amount,
            verified,
            element     = element == null ? null : SafeElementName(element),
            controlType = element == null ? null : SafeElementControlType(element),
            automationId = element == null ? null : SafeElementAutomationId(element)
        };
    }

    /// <summary>
    /// Coordinate-based wheel scroll (Case D).
    /// </summary>
    private object ScrollCoordinates(UiRequest request, CancellationToken cancellationToken)
    {
        var direction = NormalizeScrollDirection(request.Direction);
        var amount    = Math.Abs(request.Amount ?? 1);

        if (amount == 0)
            amount = 1;

        cancellationToken.ThrowIfCancellationRequested();

        var success = TryMouseWheelScrollAtPoint(
            new Point(request.X!.Value, request.Y!.Value),
            direction,
            amount,
            cancellationToken,
            out var strategy);

        if (!success)
            throw new InvalidOperationException(
                $"Coordinate scroll failed. x={request.X}, y={request.Y}, direction={direction}, amount={amount}");

        Thread.Sleep(request.ScrollDelayMs ?? 100);

        return new
        {
            operation = "scroll",
            success   = true,
            strategy,
            direction,
            amount,
            x         = request.X.Value,
            y         = request.Y.Value
        };
    }

    /// <summary>
    /// Smart scroll-into-view. Handles Cases B and C (and is also used by the
    /// scrollintoview alias).  Algorithm:
    /// 1. Optionally resolve explicit container via ContainerLocator.
    /// 2. Try to find target (including offscreen by default).
    /// 3. If target visible → success (strategy=already-visible).
    /// 4. If target found but offscreen → try ScrollItemPattern.
    /// 5. If still not visible → rectangle-align loop against nearest/provided container.
    /// 6. If target not found → scroll likely containers and retry.
    /// </summary>
    private object ScrollTargetIntoView(UiRequest request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var session      = RequireSession();
        var root         = GetWindowRoot(session, allowDesktopPopupScan: false);
        var maxAttempts  = request.MaxAttempts ?? 10;
        var delayMs      = request.DelayMs ?? 150;
        var scrollAmount = request.Amount ?? 5;
        var scrollMode   = string.IsNullOrWhiteSpace(request.Mode) ? "auto" : request.Mode.Trim().ToLowerInvariant();
        var targetLocator = request.Locator != null && !IsEmptyLocator(request.Locator)
            ? request.Locator
            : throw new ArgumentException("'scroll' target-into-view requires a Locator.");

        _logger.LogInformation(
            "ScrollTargetIntoView started. containerLocator={ContainerLocator}, targetLocator={TargetLocator}, mode={Mode}, maxAttempts={MaxAttempts}",
            SafeDescribeLocator(request.ContainerLocator),
            SafeDescribeLocator(request.Locator),
            scrollMode,
            maxAttempts);

        // Resolve explicit container if provided.
        AutomationElement? container = null;
        if (request.ContainerLocator != null && !IsEmptyLocator(request.ContainerLocator))
        {
            var containerSearchReq = new UiRequest
            {
                Operation = request.Operation,
                Locator   = request.ContainerLocator,
                TimeoutMs = request.TimeoutMs,
                Fast      = request.Fast
            };
            container = ResolveElementForOperation(
                containerSearchReq,
                purpose: "scroll",
                action: true,
                allowOffscreen: false,
                requireClickable: false);
            BringElementWindowToForeground(container);
            Thread.Sleep(WindowActivationDelayMs);
        }
        else if (request.Locator != null && !IsEmptyLocator(request.Locator))
        {
            // Window must be foreground before searching; bring it via the session root.
            BringElementWindowToForeground(root);
            Thread.Sleep(WindowActivationDelayMs);
        }

        var searchRoot = container ?? root;

        // Attempt 1: find target including offscreen.
        var target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator);

        _logger.LogInformation(
            "Scroll target lookup result. targetFound={TargetFound}, targetName={TargetName}, automationId={AutomationId}, controlType={ControlType}, isOffscreen={IsOffscreen}, rect={Rect}, hasScrollItemPattern={HasScrollItemPattern}, hasExpandCollapsePattern={HasExpandCollapsePattern}",
            target != null,
            target == null ? null : SafeElementName(target),
            target == null ? null : SafeElementAutomationId(target),
            target == null ? null : SafeElementControlType(target),
            target == null ? null : (bool?)target.Properties.IsOffscreen.ValueOrDefault,
            target == null ? null : target.BoundingRectangle.ToString(),
            target != null && target.Patterns.ScrollItem.PatternOrDefault != null,
            target != null && target.Patterns.ExpandCollapse.PatternOrDefault != null);

        if (target != null)
        {
            var scrollItemAvailable = target.Patterns.ScrollItem.PatternOrDefault != null;

            // Already visible?
            if (IsTargetPracticallyVisible(target, container, out var alreadyVisibleStrategy))
            {
                _logger.LogInformation(
                    "Scroll target already visible. strategy=already-visible, visibilityStrategy={VisibilityStrategy}",
                    alreadyVisibleStrategy);

                return new
                {
                    operation                    = "scroll",
                    success                      = true,
                    strategy                     = "already-visible",
                    visibilityStrategy           = alreadyVisibleStrategy,
                    targetFound                  = true,
                    targetVisible                = true,
                    scrollItemPatternAvailable   = scrollItemAvailable,
                    scrollItemPatternAttempted   = false,
                    scrollItemPatternSucceeded   = false,
                    attempts                     = 0,
                    fallbackUsed                 = false,
                    stoppedReason                = "target-already-visible"
                };
            }

            cancellationToken.ThrowIfCancellationRequested();

            // Try ScrollItemPattern.
            var attemptedScrollItem  = false;
            var scrollItemSuccess    = false;
            var scrollItemStrategy   = "scrollitempattern-not-attempted";

            var scrollItem = target.Patterns.ScrollItem.PatternOrDefault;

            if (scrollItem != null)
            {
                attemptedScrollItem = true;

                _logger.LogInformation(
                    "Attempting ScrollItemPattern.ScrollIntoView for target={Target}, automationId={AutomationId}",
                    SafeElementName(target),
                    SafeElementAutomationId(target));

                try
                {
                    scrollItem.ScrollIntoView();
                    scrollItemSuccess  = true;
                    scrollItemStrategy = "scrollitempattern";

                    Thread.Sleep(delayMs);

                    // Re-find after ScrollIntoView so UIA state is refreshed.
                    target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator) ?? target;

                    _logger.LogInformation(
                        "ScrollItemPattern result. attempted={Attempted}, success={Success}, strategy={Strategy}",
                        attemptedScrollItem,
                        scrollItemSuccess,
                        scrollItemStrategy);

                    if (IsTargetPracticallyVisible(target, container, out var sipVisibleStrategy))
                    {
                        return new
                        {
                            operation                    = "scroll",
                            success                      = true,
                            strategy                     = "scrollitempattern",
                            visibilityStrategy           = sipVisibleStrategy,
                            targetFound                  = true,
                            targetVisible                = true,
                            scrollItemPatternAvailable   = scrollItemAvailable,
                            scrollItemPatternAttempted   = attemptedScrollItem,
                            scrollItemPatternSucceeded   = scrollItemSuccess,
                            attempts                     = 1,
                            fallbackUsed                 = false,
                            stoppedReason                = "scrollitempattern-success"
                        };
                    }
                }
                catch (Exception ex)
                {
                    scrollItemStrategy = "scrollitempattern-failed";

                    _logger.LogWarning(
                        ex,
                        "ScrollItemPattern.ScrollIntoView failed. target={Target}, automationId={AutomationId}",
                        SafeElementName(target),
                        SafeElementAutomationId(target));

                    _logger.LogInformation(
                        "ScrollItemPattern result. attempted={Attempted}, success={Success}, strategy={Strategy}",
                        attemptedScrollItem,
                        scrollItemSuccess,
                        scrollItemStrategy);
                }
            }
            else
            {
                scrollItemStrategy = "scrollitempattern-not-available";

                _logger.LogInformation(
                    "ScrollItemPattern not available on found target. target={Target}, automationId={AutomationId}, controlType={ControlType}",
                    SafeElementName(target),
                    SafeElementAutomationId(target),
                    SafeElementControlType(target));

                _logger.LogInformation(
                    "ScrollItemPattern result. attempted={Attempted}, success={Success}, strategy={Strategy}",
                    attemptedScrollItem,
                    scrollItemSuccess,
                    scrollItemStrategy);
            }

            // ScrollItemPattern failed or did not make it visible.
            // Use container rectangle-align fallback with limited attempts.
            var scrollContainer = container
                ?? FindNearestScrollableContainer(target)
                ?? root;

            // When target was found, limit fallback wheel attempts to avoid over-scrolling.
            var targetFoundFallbackAttempts = Math.Min(maxAttempts, ScrollTargetFoundMaxFallbackAttempts);

            Rectangle? previousRect = null;
            var unchangedCount = 0;

            // Rectangle-align loop.
            for (var attempt = 1; attempt <= targetFoundFallbackAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Re-find target to get current UIA state before checking visibility.
                target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator) ?? target;

                if (IsTargetPracticallyVisible(target, scrollContainer, out var loopVisibleStrategy))
                {
                    return new
                    {
                        operation                    = "scroll",
                        success                      = true,
                        strategy                     = "container-wheel-rect-align",
                        visibilityStrategy           = loopVisibleStrategy,
                        targetFound                  = true,
                        targetVisible                = true,
                        scrollItemPatternAvailable   = scrollItemAvailable,
                        scrollItemPatternAttempted   = attemptedScrollItem,
                        scrollItemPatternSucceeded   = scrollItemSuccess,
                        attempts                     = attempt,
                        fallbackUsed                 = true,
                        stoppedReason                = "target-visible"
                    };
                }

                var targetRect    = TrySafeGetBoundingRect(target);
                var containerRect = TrySafeGetBoundingRect(scrollContainer);

                // No-progress guard: if target rectangle has not moved, stop.
                if (targetRect.HasValue)
                {
                    if (previousRect.HasValue && RectanglesEqual(previousRect.Value, targetRect.Value))
                    {
                        unchangedCount++;
                    }
                    else
                    {
                        unchangedCount = 0;
                    }

                    if (unchangedCount >= 2)
                    {
                        return new
                        {
                            operation                    = "scroll",
                            success                      = false,
                            strategy                     = "container-wheel-rect-align",
                            targetFound                  = true,
                            targetVisible                = false,
                            scrollItemPatternAvailable   = scrollItemAvailable,
                            scrollItemPatternAttempted   = attemptedScrollItem,
                            scrollItemPatternSucceeded   = scrollItemSuccess,
                            attempts                     = attempt,
                            fallbackUsed                 = true,
                            stoppedReason                = "no-scroll-progress",
                            message                      = "Target found but scroll is not moving it. Container may not be scrollable."
                        };
                    }

                    previousRect = targetRect;
                }

                string direction;
                if (targetRect.HasValue && containerRect.HasValue)
                {
                    direction = DecideScrollDirectionToTarget(targetRect.Value, containerRect.Value);
                    if (direction == "none")
                        break;
                }
                else
                {
                    direction = "down";
                }

                // Re-check visibility immediately before wheel scroll to avoid unnecessary scrolling.
                target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator) ?? target;
                var targetVisibleBeforeWheel = IsTargetPracticallyVisible(target, scrollContainer, out var visibleBeforeWheelStrategy);

                _logger.LogInformation(
                    "Before wheel fallback scroll. attempt={Attempt}, targetFound={TargetFound}, targetVisible={TargetVisible}, visibilityStrategy={VisibilityStrategy}, direction={Direction}, strategy={Strategy}",
                    attempt,
                    target != null,
                    targetVisibleBeforeWheel,
                    visibleBeforeWheelStrategy,
                    direction,
                    scrollItemStrategy);

                if (targetVisibleBeforeWheel)
                {
                    return new
                    {
                        operation                    = "scroll",
                        success                      = true,
                        strategy                     = "visible-before-wheel-scroll",
                        visibilityStrategy           = visibleBeforeWheelStrategy,
                        targetFound                  = true,
                        targetVisible                = true,
                        scrollItemPatternAvailable   = scrollItemAvailable,
                        scrollItemPatternAttempted   = attemptedScrollItem,
                        scrollItemPatternSucceeded   = scrollItemSuccess,
                        attempts                     = attempt,
                        fallbackUsed                 = true,
                        stoppedReason                = "target-visible-before-wheel"
                    };
                }

                cancellationToken.ThrowIfCancellationRequested();
                TryScrollContainerOneStep(scrollContainer, direction, scrollMode, scrollAmount, cancellationToken, out _);

                Thread.Sleep(delayMs);
            }

            // Final visibility check.
            target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator) ?? target;
            if (IsTargetPracticallyVisible(target, container, out var finalVisibleStrategy))
            {
                return new
                {
                    operation                    = "scroll",
                    success                      = true,
                    strategy                     = "container-wheel-rect-align",
                    visibilityStrategy           = finalVisibleStrategy,
                    targetFound                  = true,
                    targetVisible                = true,
                    scrollItemPatternAvailable   = scrollItemAvailable,
                    scrollItemPatternAttempted   = attemptedScrollItem,
                    scrollItemPatternSucceeded   = scrollItemSuccess,
                    attempts                     = targetFoundFallbackAttempts,
                    fallbackUsed                 = true,
                    stoppedReason                = "target-visible"
                };
            }

            return new
            {
                operation                    = "scroll",
                success                      = false,
                strategy                     = "target-found-but-not-visible",
                targetFound                  = true,
                targetVisible                = false,
                scrollItemPatternAvailable   = scrollItemAvailable,
                scrollItemPatternAttempted   = attemptedScrollItem,
                scrollItemPatternSucceeded   = scrollItemSuccess,
                attempts                     = targetFoundFallbackAttempts,
                fallbackUsed                 = true,
                stoppedReason                = "target-found-scroll-item-pattern-and-fallback-failed",
                message                      = "Target was found, but driver could not verify visibility after ScrollItemPattern and fallback scrolling."
            };
        }

        // Target not found: scroll likely containers and retry.
        var scrollContainers = container != null
            ? new[] { container }
            : FindLikelyScrollableContainers(session, root);

        var searchAttempts = 0;

        foreach (var sc in scrollContainers)
        {
            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                searchAttempts++;
                TryScrollContainerOneStep(sc, "down", scrollMode, scrollAmount, cancellationToken, out _);
                Thread.Sleep(delayMs);

                target = TryFindElementIncludingOffscreen(session, searchRoot, targetLocator);
                if (target != null && IsTargetPracticallyVisible(target, container, out var foundVisibleStrategy))
                {
                    return new
                    {
                        operation                    = "scroll",
                        success                      = true,
                        strategy                     = "scroll-search-loop",
                        visibilityStrategy           = foundVisibleStrategy,
                        targetFound                  = true,
                        targetVisible                = true,
                        scrollItemPatternAvailable   = target.Patterns.ScrollItem.PatternOrDefault != null,
                        scrollItemPatternAttempted   = false,
                        scrollItemPatternSucceeded   = false,
                        attempts                     = searchAttempts,
                        fallbackUsed                 = true,
                        stoppedReason                = "target-found-and-visible"
                    };
                }
            }
        }

        return new
        {
            operation                    = "scroll",
            success                      = false,
            strategy                     = "scroll-search-loop",
            targetFound                  = false,
            targetVisible                = false,
            scrollItemPatternAvailable   = false,
            scrollItemPatternAttempted   = false,
            scrollItemPatternSucceeded   = false,
            attempts                     = searchAttempts,
            fallbackUsed                 = true,
            stoppedReason                = "target-not-found-after-scroll",
            message                      = request.ContainerLocator != null
                ? "Target not found after scrolling the specified container."
                : "Target not found after scrolling. Provide containerLocator for better accuracy."
        };
    }

    /// <summary>
    /// Finds an element including off-screen descendants (does not filter by visibility).
    /// </summary>
    private AutomationElement? TryFindElementIncludingOffscreen(
        AutomationSession session,
        AutomationElement searchRoot,
        UiLocator locator)
    {
        try
        {
            // Use existing fast-attribute strategy which searches all descendants
            // via FindFirstDescendant (no visibility filter in FlaUI).
            var result = TryFindElementBySmartStrategy(
                searchRoot,
                session,
                locator,
                preferAttributes: true,
                xpathOnly: false);

            return result.Element;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Checks whether two rectangles have the same position and size.
    /// Used to detect no-scroll-progress in the container fallback loop.
    /// </summary>
    private static bool RectanglesEqual(Rectangle a, Rectangle b) =>
        a.Left == b.Left && a.Top == b.Top && a.Right == b.Right && a.Bottom == b.Bottom;

    /// <summary>
    /// Practical visibility check that handles DataItem rows whose UIA IsOffscreen
    /// property is unreliable. Returns true when the element is practically visible:
    /// its bounding rectangle intersects the container (or screen), and either
    /// UIA confirms it is on-screen, or DataItem-specific heuristics apply.
    /// </summary>
    private bool IsTargetPracticallyVisible(
        AutomationElement target,
        AutomationElement? container,
        out string visibilityStrategy)
    {
        visibilityStrategy = "none";

        try
        {
            var targetRect = target.BoundingRectangle;

            if (targetRect.IsEmpty || targetRect.Width <= 0 || targetRect.Height <= 0)
                return false;

            var intersectsContainer = true;

            if (container != null)
            {
                var containerRect = container.BoundingRectangle;

                if (containerRect.IsEmpty || containerRect.Width <= 0 || containerRect.Height <= 0)
                    return false;

                intersectsContainer = RectanglesIntersect(targetRect, containerRect);
            }

            if (!intersectsContainer)
                return false;

            // Strong signal: UIA says visible and rect intersects container/window.
            if (SafeIsOffscreen(target) == false)
            {
                visibilityStrategy = "isoffscreen-false-rect-intersects";
                return true;
            }

            // DataItem-specific fallback:
            // Some row controls report IsOffscreen=true even though the row is visible.
            if (target.ControlType == ControlType.DataItem)
            {
                // If the element has a non-empty name, treat it as visible.
                var name = SafeElementName(target);

                if (!string.IsNullOrWhiteSpace(name))
                {
                    visibilityStrategy = "dataitem-rect-intersects-name-present";
                    return true;
                }

                // Check if any visible child is within the container bounds.
                var visibleChild = target.FindAllDescendants()
                    .FirstOrDefault(child =>
                    {
                        try
                        {
                            var childRect = child.BoundingRectangle;

                            if (childRect.IsEmpty || childRect.Width <= 0 || childRect.Height <= 0)
                                return false;

                            if (container != null &&
                                !RectanglesIntersect(childRect, container.BoundingRectangle))
                                return false;

                            var childName = SafeElementName(child);

                            return SafeIsOffscreen(child) == false &&
                                   !string.IsNullOrWhiteSpace(childName);
                        }
                        catch
                        {
                            return false;
                        }
                    });

                if (visibleChild != null)
                {
                    visibilityStrategy = "dataitem-visible-child";
                    return true;
                }

                // Fallback for owner-drawn/odd grids: the rectangle already intersects
                // the container (verified by the intersectsContainer check above), so treat
                // the DataItem as visible even if UIA IsOffscreen is unreliable for
                // virtualised table rows.
                visibilityStrategy = "dataitem-rect-intersects";
                return true;
            }

            // Last fallback: bounding rect intersects container but IsOffscreen is unreliable
            // (e.g. virtualised rows that UIA has not yet refreshed). We treat intersection as
            // sufficient evidence of visibility; callers should verify with a re-check loop.
            visibilityStrategy = "rect-intersects";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "Practical visibility check failed. target={Target}",
                SafeElementName(target));

            return false;
        }
    }

    /// <summary>
    /// Returns true when the target element is visible within the container's bounds.
    /// When container is null, returns true if the element is simply not off-screen.
    /// </summary>
    private static bool IsTargetVisibleInContainer(
        AutomationElement target,
        AutomationElement? container)
    {
        try
        {
            if (target.IsOffscreen)
                return false;

            var targetRect = target.BoundingRectangle;

            if (targetRect.IsEmpty || targetRect.Width <= 0 || targetRect.Height <= 0)
                return false;

            if (container == null)
                return true;

            var containerRect = container.BoundingRectangle;

            if (containerRect.IsEmpty || containerRect.Width <= 0 || containerRect.Height <= 0)
                return false;

            return RectanglesIntersect(targetRect, containerRect);
        }
        catch
        {
            return false;
        }
    }

    private static bool RectanglesIntersect(Rectangle a, Rectangle b)
    {
        return a.Left < b.Right &&
               a.Right > b.Left &&
               a.Top < b.Bottom &&
               a.Bottom > b.Top;
    }

    private static string DecideScrollDirectionToTarget(Rectangle targetRect, Rectangle containerRect)
    {
        if (targetRect.Bottom > containerRect.Bottom)
            return "down";

        if (targetRect.Top < containerRect.Top)
            return "up";

        if (targetRect.Right > containerRect.Right)
            return "right";

        if (targetRect.Left < containerRect.Left)
            return "left";

        return "none";
    }

    private static Rectangle? TrySafeGetBoundingRect(AutomationElement element)
    {
        try
        {
            var r = element.BoundingRectangle;
            return r.IsEmpty ? (Rectangle?)null : r;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Walks the UIA parent chain to find the nearest element that has a ScrollPattern
    /// or is a likely scrollable container (DataGrid, Table, List, Pane, Document).
    /// </summary>
    private AutomationElement? FindNearestScrollableContainer(AutomationElement target)
    {
        var current = target.Parent;

        while (current != null)
        {
            try
            {
                var scrollPattern = current.Patterns.Scroll.PatternOrDefault;

                if (scrollPattern != null &&
                    (scrollPattern.VerticallyScrollable || scrollPattern.HorizontallyScrollable))
                {
                    return current;
                }

                var ct = current.ControlType;

                if (ct == ControlType.DataGrid ||
                    ct == ControlType.Table     ||
                    ct == ControlType.List      ||
                    ct == ControlType.Pane      ||
                    ct == ControlType.Document)
                {
                    return current;
                }
            }
            catch
            {
                // Ignore and continue parent chain.
            }

            current = current.Parent;
        }

        return null;
    }

    /// <summary>
    /// Finds likely scrollable containers in the current window for target-not-found search.
    /// </summary>
    private AutomationElement[] FindLikelyScrollableContainers(
        AutomationSession session,
        AutomationElement root)
    {
        try
        {
            var cf = session.Automation.ConditionFactory;
            var candidates = new List<AutomationElement>();

            foreach (var ct in new[]
            {
                ControlType.DataGrid,
                ControlType.Table,
                ControlType.List,
                ControlType.Tree,
                ControlType.Document,
                ControlType.Pane
            })
            {
                try
                {
                    var found = root.FindAllDescendants(cf.ByControlType(ct));
                    candidates.AddRange(found);
                }
                catch
                {
                    // Ignore per-type failures.
                }
            }

            return candidates.ToArray();
        }
        catch
        {
            return Array.Empty<AutomationElement>();
        }
    }

    /// <summary>
    /// Attempts to use UIA ScrollItemPattern.ScrollIntoView on the target element.
    /// </summary>
    private bool TryScrollItemPattern(AutomationElement target, out string strategy)
    {
        strategy = "scrollitempattern";

        try
        {
            var scrollItem = target.Patterns.ScrollItem.PatternOrDefault;

            if (scrollItem == null)
            {
                strategy = "scrollitempattern-not-supported";
                return false;
            }

            scrollItem.ScrollIntoView();
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "ScrollItemPattern.ScrollIntoView failed. target={Target}",
                SafeElementName(target));

            strategy = "scrollitempattern-failed";
            return false;
        }
    }

    /// <summary>
    /// Scrolls a container one step using the appropriate mode.
    /// </summary>
    private bool TryScrollContainerOneStep(
        AutomationElement container,
        string direction,
        string mode,
        int amount,
        CancellationToken cancellationToken,
        out string strategy)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (mode == "pattern" || mode == "auto")
        {
            if (TryScrollByPattern(container, direction, amount, out strategy))
                return true;

            if (mode == "pattern")
                return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return TryMouseWheelScrollOnElement(container, direction, amount, cancellationToken, out strategy);
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

    private object? Select(UiRequest req, CancellationToken cancellationToken = default)
    {
        if (req.Value == null && req.Index == null)
            throw new ArgumentException("Either 'value' (item name) or 'index' is required for 'select'.");

        // Log when the request arrives via the deprecated alias so it is visible in traces.
        if (string.Equals(req.Operation, "selectcomboboxitem", StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogInformation(
                "selectcomboboxitem is deprecated alias. Routing through canonical select pipeline.");
        }

        if (IsComboBoxSelectRequest(req))
        {
            var nativeSession = RequireSession();
            return _nativeUiaComboBoxService.SelectComboBox(
                req,
                GetActiveWindowHwndOrNull(nativeSession),
                GetActiveProcessIdOrNull(nativeSession),
                cancellationToken);
        }

        var session = RequireSession();
        var element = ResolveElementForOperation(
            req,
            purpose: "select",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        if (IsComboBoxElement(element))
        {
            _logger.LogInformation(
                "Select resolved ComboBox element. Routing to native UIA-only pipeline. operation={Operation}, value={Value}, index={Index}, comboBox={ComboBox}",
                SanitizeValue(req.Operation),
                SanitizeValue(req.Value),
                req.Index,
                SafeElementName(element));

            return _nativeUiaComboBoxService.SelectComboBox(
                req,
                GetActiveWindowHwndOrNull(session),
                GetActiveProcessIdOrNull(session),
                cancellationToken);
        }

        var cf = session.Automation.ConditionFactory;

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

        var element = ResolveElementForOperation(
            req,
            purpose: "selectaid",
            action: true,
            allowOffscreen: false,
            requireClickable: false);
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

        var element = ResolveElementForOperation(
            req,
            purpose: "typeandselect",
            action: true,
            allowOffscreen: false,
            requireClickable: false);
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
                var selectionItemPattern = item.Patterns.SelectionItem.Pattern;
                selectionItemPattern.Select();

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
                var invokePattern = item.Patterns.Invoke.Pattern;
                invokePattern.Invoke();
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

    private object DragByOffset(UiRequest request)
    {
        var offsetX = request.OffsetX ?? 0;
        var offsetY = request.OffsetY ?? 0;

        var hasFromTo =
            request.FromX.HasValue ||
            request.FromY.HasValue ||
            request.ToX.HasValue ||
            request.ToY.HasValue;

        var hasStartPoint =
            request.X.HasValue ||
            request.Y.HasValue;

        var hasLocator = !IsEmptyLocator(request.Locator);

        // Mode 1: from/to coordinate drag — delegate to DragCoordinates.
        if (hasFromTo)
        {
            // hasFromTo is true when *any* of the four fields is present; require all four.
            if (!request.FromX.HasValue ||
                !request.FromY.HasValue ||
                !request.ToX.HasValue ||
                !request.ToY.HasValue)
            {
                throw new ArgumentException(
                    "'dragbyoffset' coordinate start/end mode requires fromX, fromY, toX, and toY.");
            }

            return DragCoordinates(request);
        }

        // Mode 2: x/y + offset coordinate drag.
        if (hasStartPoint)
        {
            // hasStartPoint is true when *either* x or y is present; require both.
            if (!request.X.HasValue || !request.Y.HasValue)
            {
                throw new ArgumentException(
                    "'dragbyoffset' coordinate offset mode requires both x and y.");
            }

            if (offsetX == 0 && offsetY == 0)
            {
                throw new ArgumentException(
                    "'dragbyoffset' coordinate offset mode requires non-zero offsetX or offsetY.");
            }

            return DragByCoordinateOffset(request);
        }

        // Mode 3: locator-based drag.
        if (!hasLocator)
        {
            throw new ArgumentException(
                "'dragbyoffset' requires either locator, x/y + offsetX/offsetY, or fromX/fromY/toX/toY.");
        }

        if (offsetX == 0 && offsetY == 0)
        {
            throw new ArgumentException(
                "'dragbyoffset' locator mode requires non-zero offsetX or offsetY.");
        }

        return DragElementByOffset(request);
    }

    private object DragElementByOffset(UiRequest request)
    {
        var element = ResolveElementForOperation(
            request,
            purpose: "dragbyoffset",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

        // Physical input must target foreground window.
        BringElementWindowToForeground(element);
        Thread.Sleep(WindowActivationDelayMs);

        var offsetX = request.OffsetX ?? 0;
        var offsetY = request.OffsetY ?? 0;

        var dragStart = string.IsNullOrWhiteSpace(request.DragStart)
            ? "center"
            : request.DragStart.Trim();

        var durationMs = request.DragDurationMs ?? 250;
        var steps = request.DragSteps ?? 10;
        var button = NormalizeMouseButton(request.Button);

        if (durationMs < 0)
            durationMs = 0;

        if (steps < 1)
            steps = 1;

        var rect = element.BoundingRectangle;

        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
        {
            throw new InvalidOperationException(
                $"Cannot drag element because bounding rectangle is empty. element={SafeElementName(element)}");
        }

        var start = GetDragStartPoint(rect, dragStart);
        var end = new Point(start.X + offsetX, start.Y + offsetY);

        _logger.LogInformation(
            "DragByOffset element mode. element={Element}, dragStart={DragStart}, start=({StartX},{StartY}), end=({EndX},{EndY}), offset=({OffsetX},{OffsetY}), button={Button}, durationMs={DurationMs}, steps={Steps}",
            SafeElementName(element),
            dragStart,
            start.X,
            start.Y,
            end.X,
            end.Y,
            offsetX,
            offsetY,
            button,
            durationMs,
            steps);

        if (!PerformPhysicalDrag(start, end, durationMs, steps, button))
        {
            throw new InvalidOperationException(
                $"dragbyoffset element mode failed. element={SafeElementName(element)}, start=({start.X},{start.Y}), end=({end.X},{end.Y})");
        }

        Thread.Sleep(150);

        return new
        {
            operation = "dragbyoffset",
            success = true,
            strategy = "element-offset",
            element = SafeElementName(element),
            controlType = SafeElementControlType(element),
            className = SafeElementClassName(element),
            automationId = SafeElementAutomationId(element),
            dragStart,
            offsetX,
            offsetY,
            start = new { x = start.X, y = start.Y },
            end = new { x = end.X, y = end.Y },
            button,
            durationMs,
            steps
        };
    }

    private object DragByCoordinateOffset(UiRequest request)
    {
        var offsetX = request.OffsetX ?? 0;
        var offsetY = request.OffsetY ?? 0;

        var durationMs = request.DragDurationMs ?? 250;
        var steps = request.DragSteps ?? 10;
        var button = NormalizeMouseButton(request.Button);

        if (durationMs < 0)
            durationMs = 0;

        if (steps < 1)
            steps = 1;

        var start = new Point(request.X!.Value, request.Y!.Value);
        var end = new Point(start.X + offsetX, start.Y + offsetY);

        _logger.LogInformation(
            "DragByOffset coordinate mode. start=({StartX},{StartY}), end=({EndX},{EndY}), offset=({OffsetX},{OffsetY}), button={Button}, durationMs={DurationMs}, steps={Steps}",
            start.X,
            start.Y,
            end.X,
            end.Y,
            offsetX,
            offsetY,
            button,
            durationMs,
            steps);

        if (!PerformPhysicalDrag(start, end, durationMs, steps, button))
        {
            throw new InvalidOperationException(
                $"dragbyoffset coordinate mode failed. start=({start.X},{start.Y}), end=({end.X},{end.Y})");
        }

        Thread.Sleep(150);

        return new
        {
            operation = "dragbyoffset",
            success = true,
            strategy = "coordinate-offset",
            offsetX,
            offsetY,
            start = new { x = start.X, y = start.Y },
            end = new { x = end.X, y = end.Y },
            button,
            durationMs,
            steps
        };
    }

    private object DragCoordinates(UiRequest request)
    {
        if (!request.FromX.HasValue || !request.FromY.HasValue ||
            !request.ToX.HasValue   || !request.ToY.HasValue)
        {
            throw new ArgumentException(
                "'dragcoordinates' requires fromX, fromY, toX, and toY.");
        }

        var start = new Point(request.FromX.Value, request.FromY.Value);
        var end   = new Point(request.ToX.Value,   request.ToY.Value);

        var durationMs = request.DragDurationMs ?? 250;
        var steps      = request.DragSteps      ?? 10;

        if (durationMs < 0)
            durationMs = 0;

        if (steps < 1)
            steps = 1;

        var button = NormalizeMouseButton(request.Button);

        _logger.LogInformation(
            "DragCoordinates starting. start=({StartX},{StartY}), end=({EndX},{EndY}), durationMs={DurationMs}, steps={Steps}",
            start.X,
            start.Y,
            end.X,
            end.Y,
            durationMs,
            steps);

        if (!PerformPhysicalDrag(start, end, durationMs, steps, button))
        {
            throw new InvalidOperationException(
                $"dragcoordinates failed. start=({start.X},{start.Y}), end=({end.X},{end.Y})");
        }

        return new
        {
            operation = "dragcoordinates",
            success   = true,
            strategy  = "coordinate-drag",
            button,
            start     = new { x = start.X, y = start.Y },
            end       = new { x = end.X,   y = end.Y   },
            durationMs,
            steps
        };
    }

    private object MouseAction(UiRequest request)
    {
        var action = string.IsNullOrWhiteSpace(request.Action)
            ? throw new ArgumentException("'mouse' operation requires action.")
            : request.Action.Trim().ToLowerInvariant();

        var button = NormalizeMouseButton(request.Button);

        var x = request.X;
        var y = request.Y;

        if (action is "move" or "down" or "up" or "click" or "doubleclick" or "rightclick"
            && (!x.HasValue || !y.HasValue))
        {
            throw new ArgumentException(
                $"'mouse' action '{action}' requires x and y.");
        }

        switch (action)
        {
            case "move":
                if (!SendMouseMoveTo(x!.Value, y!.Value))
                {
                    throw new InvalidOperationException(
                        $"'mouse' action 'move' failed. x={x.Value}, y={y.Value}");
                }
                break;

            case "down":
                SetCursorPos(x!.Value, y!.Value);
                MouseDown(button);
                break;

            case "up":
                SetCursorPos(x!.Value, y!.Value);
                MouseUp(button);
                break;

            case "click":
                SetCursorPos(x!.Value, y!.Value);
                MouseDown(button);
                Thread.Sleep(50);
                MouseUp(button);
                break;

            case "doubleclick":
                SetCursorPos(x!.Value, y!.Value);
                MouseDown(button);
                Thread.Sleep(50);
                MouseUp(button);
                Thread.Sleep(75);
                MouseDown(button);
                Thread.Sleep(50);
                MouseUp(button);
                break;

            case "rightclick":
                SetCursorPos(x!.Value, y!.Value);
                MouseDown("right");
                Thread.Sleep(50);
                MouseUp("right");
                break;

            case "drag":
                return DragCoordinates(request);

            case "scroll":
            {
                if (!x.HasValue || !y.HasValue)
                    throw new ArgumentException("'mouse' action 'scroll' requires x and y.");

                var wheelDelta = request.WheelDelta;

                if (!wheelDelta.HasValue)
                {
                    var direction = NormalizeScrollDirection(request.Direction);
                    var amount    = request.Amount ?? 1;
                    wheelDelta    = ToWheelDelta(direction, amount);
                }

                SetCursorPos(x.Value, y.Value);
                Thread.Sleep(50);

                var horizontal = IsHorizontalScrollDelta(request.Direction, wheelDelta.Value);

                if (!SendMouseWheelRawDelta(wheelDelta.Value, horizontal))
                    throw new InvalidOperationException(
                        $"'mouse' action 'scroll' SendInput failed. x={x.Value}, y={y.Value}, wheelDelta={wheelDelta.Value}");

                return new
                {
                    operation  = "mouse",
                    success    = true,
                    action     = "scroll",
                    x          = x.Value,
                    y          = y.Value,
                    wheelDelta = wheelDelta.Value
                };
            }

            default:
                throw new ArgumentException(
                    $"Unsupported mouse action '{action}'. Supported: move, down, up, click, doubleclick, rightclick, drag, scroll.");
        }

        return new
        {
            operation = "mouse",
            success   = true,
            action,
            button,
            x,
            y
        };
    }

    private static Point GetDragStartPoint(
        Rectangle rect,
        string dragStart)
    {
        var normalized = dragStart.ToLowerInvariant();

        var left    = rect.Left;
        var right   = rect.Right;
        var top     = rect.Top;
        var bottom  = rect.Bottom;
        var centerX = rect.Left + rect.Width / 2;
        var centerY = rect.Top  + rect.Height / 2;

        // Keep point slightly inside the element border.
        const int inset = 3;

        return normalized switch
        {
            "center"      => new Point(centerX, centerY),

            "topedge"     => new Point(centerX, top    + inset),
            "bottomedge"  => new Point(centerX, bottom - inset),
            "leftedge"    => new Point(left  + inset, centerY),
            "rightedge"   => new Point(right - inset, centerY),

            "topleft"     => new Point(left  + inset, top    + inset),
            "topright"    => new Point(right - inset, top    + inset),
            "bottomleft"  => new Point(left  + inset, bottom - inset),
            "bottomright" => new Point(right - inset, bottom - inset),

            _ => throw new ArgumentException(
                $"Unsupported dragStart '{dragStart}'. Supported values (case-insensitive): center, topEdge, bottomEdge, leftEdge, rightEdge, topLeft, topRight, bottomLeft, bottomRight.")
        };
    }

    private bool PerformPhysicalDrag(
        Point start,
        Point end,
        int durationMs,
        int steps,
        string button = "left")
    {
        button = NormalizeMouseButton(button);

        var btnDown = button switch
        {
            "right"  => MOUSEEVENTF_RIGHTDOWN,
            "middle" => MOUSEEVENTF_MIDDLEDOWN,
            _        => MOUSEEVENTF_LEFTDOWN
        };
        var btnUp = button switch
        {
            "right"  => MOUSEEVENTF_RIGHTUP,
            "middle" => MOUSEEVENTF_MIDDLEUP,
            _        => MOUSEEVENTF_LEFTUP
        };

        var buttonIsDown = false;

        try
        {
            steps = Math.Max(steps, 1);
            durationMs = Math.Max(durationMs, 0);

            _logger.LogInformation(
                "PerformPhysicalDrag starting. start=({StartX},{StartY}), end=({EndX},{EndY}), button={Button}, durationMs={DurationMs}, steps={Steps}",
                start.X,
                start.Y,
                end.X,
                end.Y,
                button,
                durationMs,
                steps);

            if (!SetCursorPos(start.X, start.Y))
            {
                _logger.LogWarning(
                    "PerformPhysicalDrag: SetCursorPos start failed. start=({StartX},{StartY}), lastError={LastError}",
                    start.X,
                    start.Y,
                    Marshal.GetLastWin32Error());

                return false;
            }

            Thread.Sleep(DragMoveToStartDelayMs);

            if (!SendMouseInputChecked(btnDown, $"drag-{button}-down"))
            {
                return false;
            }

            buttonIsDown = true;

            // Important for splitter resize: app needs time to enter capture/resize mode.
            Thread.Sleep(DragMouseDownHoldDelayMs);

            var stepDelay = Math.Max(DragMinimumStepDelayMs, durationMs / steps);

            for (var i = 1; i <= steps; i++)
            {
                var t = i / (double)steps;

                var x = (int)Math.Round(start.X + ((end.X - start.X) * t));
                var y = (int)Math.Round(start.Y + ((end.Y - start.Y) * t));

                // Use SendInput mouse move instead of only SetCursorPos.
                if (!SendMouseMoveTo(x, y))
                {
                    _logger.LogWarning(
                        "PerformPhysicalDrag: SendMouseMoveTo failed at step {Step}/{Steps}. point=({X},{Y}), lastError={LastError}",
                        i,
                        steps,
                        x,
                        y,
                        Marshal.GetLastWin32Error());

                    // fallback to SetCursorPos, but still keep button down
                    SetCursorPos(x, y);
                }

                Thread.Sleep(stepDelay);
            }

            Thread.Sleep(DragMouseUpSettleDelayMs);

            if (!SendMouseInputChecked(btnUp, $"drag-{button}-up"))
            {
                return false;
            }

            buttonIsDown = false;

            Thread.Sleep(DragMouseUpSettleDelayMs);

            _logger.LogInformation(
                "PerformPhysicalDrag completed. start=({StartX},{StartY}), end=({EndX},{EndY}), button={Button}",
                start.X,
                start.Y,
                end.X,
                end.Y,
                button);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "PerformPhysicalDrag failed. start=({StartX},{StartY}), end=({EndX},{EndY}), button={Button}",
                start.X,
                start.Y,
                end.X,
                end.Y,
                button);

            return false;
        }
        finally
        {
            if (buttonIsDown)
            {
                try
                {
                    _logger.LogWarning("PerformPhysicalDrag cleanup: mouse button still down, sending button up. button={Button}", button);
                    SendMouseInputChecked(btnUp, $"drag-{button}-up-cleanup");
                }
                catch
                {
                    // ignored
                }
            }
        }
    }

    /// <summary>
    /// Normalizes a raw button string to one of "left", "right", or "middle".
    /// Returns "left" when <paramref name="button"/> is null or whitespace.
    /// Throws <see cref="ArgumentException"/> for any other value.
    /// </summary>
    private static string NormalizeMouseButton(string? button)
    {
        if (string.IsNullOrWhiteSpace(button))
            return "left";

        var normalized = button.Trim().ToLowerInvariant();

        return normalized switch
        {
            "left"   => "left",
            "right"  => "right",
            "middle" => "middle",
            _ => throw new ArgumentException(
                $"Unsupported mouse button '{button}'. Supported buttons: left, right, middle.")
        };
    }

    /// <summary>Sends the mouse-button-down event for <paramref name="button"/>.</summary>
    private void MouseDown(string button)
    {
        button = NormalizeMouseButton(button);

        switch (button)
        {
            case "left":
                SendMouseInput(MOUSEEVENTF_LEFTDOWN);
                break;

            case "right":
                SendMouseInput(MOUSEEVENTF_RIGHTDOWN);
                break;

            case "middle":
                SendMouseInput(MOUSEEVENTF_MIDDLEDOWN);
                break;

            default:
                throw new InvalidOperationException($"Unexpected button value after normalization: '{button}'.");
        }
    }

    /// <summary>Sends the mouse-button-up event for <paramref name="button"/>.</summary>
    private void MouseUp(string button)
    {
        button = NormalizeMouseButton(button);

        switch (button)
        {
            case "left":
                SendMouseInput(MOUSEEVENTF_LEFTUP);
                break;

            case "right":
                SendMouseInput(MOUSEEVENTF_RIGHTUP);
                break;

            case "middle":
                SendMouseInput(MOUSEEVENTF_MIDDLEUP);
                break;

            default:
                throw new InvalidOperationException($"Unexpected button value after normalization: '{button}'.");
        }
    }

    /// <summary>Sends a single mouse input event via SendInput.</summary>
    private bool SendMouseInput(uint dwFlags)
    {
        var input = new[]
        {
            new INPUT
            {
                type = INPUT_MOUSE,
                U    = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx          = 0,
                        dy          = 0,
                        mouseData   = 0,
                        dwFlags     = dwFlags,
                        time        = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            }
        };

        return SendInput((uint)input.Length, input, Marshal.SizeOf<INPUT>()) == input.Length;
    }

    /// <summary>
    /// Sends a single mouse input event via SendInput and logs success or failure.
    /// </summary>
    private bool SendMouseInputChecked(uint dwFlags, string actionName)
    {
        var input = new[]
        {
            new INPUT
            {
                type = INPUT_MOUSE,
                U    = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx          = 0,
                        dy          = 0,
                        mouseData   = 0,
                        dwFlags     = dwFlags,
                        time        = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            }
        };

        var sent = SendInput((uint)input.Length, input, Marshal.SizeOf<INPUT>());

        if (sent != input.Length)
        {
            var error = Marshal.GetLastWin32Error();

            _logger.LogWarning(
                "SendMouseInput failed. action={ActionName}, flags={Flags}, sent={Sent}, expected={Expected}, lastError={LastError}",
                actionName,
                dwFlags,
                sent,
                input.Length,
                error);

            return false;
        }

        _logger.LogDebug(
            "SendMouseInput succeeded. action={ActionName}, flags={Flags}, sent={Sent}",
            actionName,
            dwFlags,
            sent);

        return true;
    }

    /// <summary>
    /// Moves the cursor to the given screen coordinates via <see cref="SetCursorPos"/> and then
    /// delivers a <c>MOUSEEVENTF_MOVE</c> event via <see cref="SendInput"/> so that the target
    /// application receives a genuine WM_MOUSEMOVE while any button is held down.
    /// </summary>
    private bool SendMouseMoveTo(int x, int y)
    {
        if (!SetCursorPos(x, y))
        {
            _logger.LogWarning(
                "SendMouseMoveTo: SetCursorPos failed. x={X}, y={Y}, lastError={LastError}",
                x,
                y,
                Marshal.GetLastWin32Error());

            return false;
        }

        var input = new[]
        {
            new INPUT
            {
                type = INPUT_MOUSE,
                U    = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx          = 0,
                        dy          = 0,
                        mouseData   = 0,
                        dwFlags     = MOUSEEVENTF_MOVE,
                        time        = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            }
        };

        var sent = SendInput((uint)input.Length, input, Marshal.SizeOf<INPUT>());

        if (sent != input.Length)
        {
            _logger.LogWarning(
                "SendMouseMoveTo: SendInput MOVE failed. x={X}, y={Y}, sent={Sent}, expected={Expected}, lastError={LastError}",
                x,
                y,
                sent,
                input.Length,
                Marshal.GetLastWin32Error());

            return false;
        }

        return true;
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

    // =========================================================================
    // Unified Popup Pipeline
    // =========================================================================

    private sealed class PopupWindowInfo
    {
        public AutomationElement Element { get; init; } = null!;
        public IntPtr Hwnd { get; init; }
        public int ProcessId { get; init; }
        public string Title { get; init; } = string.Empty;
        public string ClassName { get; init; } = string.Empty;
        public bool IsModal { get; init; }
        public bool IsForeground { get; init; }
        public bool SameProcess { get; init; }
        public int Score { get; init; }
        public string Reason { get; init; } = string.Empty;
    }

    private static object ToPopupDto(PopupWindowInfo popup) => new
    {
        hwnd       = popup.Hwnd.ToInt64(),
        hwndHex    = $"0x{popup.Hwnd.ToInt64():X}",
        popup.ProcessId,
        title      = popup.Title,
        className  = popup.ClassName,
        isModal    = popup.IsModal,
        isForeground = popup.IsForeground,
        sameProcess  = popup.SameProcess,
        score      = popup.Score,
        reason     = popup.Reason
    };

    private sealed class PopupSearchCriteria
    {
        public string? TitleOrValue { get; init; }
        public IntPtr? Hwnd { get; init; }
        public string? ClassName { get; init; }
        public int? ProcessId { get; init; }
        public string MatchMode { get; init; } = "contains";
        public bool DesktopSearch { get; init; } = true;
        public bool SameProcessOnly { get; init; }
    }

    private PopupSearchCriteria BuildPopupSearchCriteria(UiRequest request, AutomationSession? session)
    {
        IntPtr? hwnd = null;
        if (request.Hwnd.HasValue && request.Hwnd.Value != 0)
            hwnd = new IntPtr(request.Hwnd.Value);

        var matchMode = string.IsNullOrWhiteSpace(request.MatchMode)
            ? "contains"
            : request.MatchMode.Trim().ToLowerInvariant();

        if (matchMode is not ("exact" or "contains" or "regex"))
            matchMode = "contains";

        int? processId = request.ProcessId;

        if (!processId.HasValue && request.SameProcessOnly == true)
        {
            try { processId = session?.Application?.ProcessId; }
            catch { processId = null; }
        }

        return new PopupSearchCriteria
        {
            TitleOrValue   = request.Value,
            Hwnd           = hwnd,
            ClassName      = request.ClassName,
            ProcessId      = processId,
            MatchMode      = matchMode,
            DesktopSearch  = request.DesktopSearch != false,
            SameProcessOnly = request.SameProcessOnly == true
        };
    }

    private PopupWindowInfo? FindBestPopupWindow(UiRequest request, AutomationSession? session)
    {
        var criteria = BuildPopupSearchCriteria(request, session);

        // Direct HWND lookup bypasses the scoring path.
        if (criteria.Hwnd.HasValue && criteria.Hwnd.Value != IntPtr.Zero)
        {
            var byHwnd = TryBuildPopupInfoFromHwnd(criteria.Hwnd.Value, session, criteria, "hwnd");
            if (byHwnd != null)
                return byHwnd;
        }

        var candidates = new List<PopupWindowInfo>();

        // 1. App top-level windows.
        if (session?.Application != null)
        {
            try
            {
                foreach (var win in session.Application.GetAllTopLevelWindows(session.Automation))
                {
                    var info = TryBuildPopupInfo(win, session, criteria, "app-top-level");
                    if (info != null) candidates.Add(info);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Popup discovery: failed reading app top-level windows.");
            }
        }

        // 2. Active/main window descendants.
        try
        {
            var root = session?.ActiveWindow;
            if (root != null)
            {
                foreach (var win in root.FindAllDescendants(cf => cf.ByControlType(ControlType.Window)))
                {
                    var info = TryBuildPopupInfo(win, session, criteria, "active-descendant");
                    if (info != null) candidates.Add(info);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Popup discovery: failed reading active window descendants.");
        }

        // 3. Desktop-level search.
        if (criteria.DesktopSearch && session?.Automation != null)
        {
            try
            {
                var desktop = session.Automation.GetDesktop();
                foreach (var win in desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window)))
                {
                    var info = TryBuildPopupInfo(win, session, criteria, "desktop-child");
                    if (info != null) candidates.Add(info);
                }

                // Deeper scan only when no direct desktop child matched.
                if (candidates.Count == 0)
                {
                    foreach (var win in desktop.FindAllDescendants(cf => cf.ByControlType(ControlType.Window)))
                    {
                        var info = TryBuildPopupInfo(win, session, criteria, "desktop-descendant");
                        if (info != null) candidates.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Popup discovery: desktop scan failed.");
            }
        }
        else if (criteria.DesktopSearch && session?.Automation == null)
        {
            // No session – use a temporary automation for desktop scan.
            try
            {
                using var tempAuto = new UIA3Automation();
                foreach (var win in tempAuto.GetDesktop().FindAllChildren(cf => cf.ByControlType(ControlType.Window)))
                {
                    var info = TryBuildPopupInfo(win, null, criteria, "desktop-child-nosession");
                    if (info != null) candidates.Add(info);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Popup discovery: session-less desktop scan failed.");
            }
        }

        // 4. Foreground HWND fallback.
        try
        {
            var fg = GetForegroundWindow();
            if (fg != IntPtr.Zero)
            {
                var fgInfo = TryBuildPopupInfoFromHwnd(fg, session, criteria, "foreground");
                if (fgInfo != null) candidates.Add(fgInfo);
            }
        }
        catch { /* ignore */ }

        var best = candidates
            .GroupBy(x => x.Hwnd)
            .Select(g => g.OrderByDescending(x => x.Score).First())
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();

        if (best != null)
            _logger.LogInformation(
                "Popup selected. hwnd=0x{Hwnd:X}, title={Title}, class={Class}, score={Score}, reason={Reason}",
                best.Hwnd.ToInt64(), best.Title, best.ClassName, best.Score, best.Reason);

        return best;
    }

    private PopupWindowInfo? TryBuildPopupInfo(
        AutomationElement element,
        AutomationSession? session,
        PopupSearchCriteria criteria,
        string reasonPrefix)
    {
        try
        {
            var hwnd = SafeWindowHandle(element);
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return null;
            return TryBuildPopupInfoFromHwndAndElement(hwnd, element, session, criteria, reasonPrefix);
        }
        catch { return null; }
    }

    private PopupWindowInfo? TryBuildPopupInfoFromHwnd(
        IntPtr hwnd,
        AutomationSession? session,
        PopupSearchCriteria criteria,
        string reasonPrefix)
    {
        try
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return null;

            AutomationElement? element = null;
            try { element = session?.Automation?.FromHandle(hwnd); }
            catch { element = null; }

            if (element == null) return null;

            return TryBuildPopupInfoFromHwndAndElement(hwnd, element, session, criteria, reasonPrefix);
        }
        catch { return null; }
    }

    private PopupWindowInfo? TryBuildPopupInfoFromHwndAndElement(
        IntPtr hwnd,
        AutomationElement element,
        AutomationSession? session,
        PopupSearchCriteria criteria,
        string reasonPrefix)
    {
        if (!IsWindowVisible(hwnd)) return null;

        GetWindowThreadProcessId(hwnd, out var pid);
        var processId = (int)pid;
        var title     = GetWindowTitle(hwnd);
        var className = GetWin32ClassName(hwnd);

        var appPid = 0;
        try { appPid = session?.Application?.ProcessId ?? 0; }
        catch { appPid = 0; }

        var sameProcess = appPid != 0 && processId == appPid;

        if (criteria.SameProcessOnly && !sameProcess) return null;
        if (criteria.ProcessId.HasValue && processId != criteria.ProcessId.Value) return null;
        if (!TitleMatches(title, criteria.TitleOrValue, criteria.MatchMode)) return null;
        if (!ClassNameMatches(className, criteria.ClassName)) return null;

        var foreground   = GetForegroundWindow();
        var isForeground = foreground == hwnd;

        var isModal = false;
        try
        {
            var wp = element.Patterns.Window.PatternOrDefault;
            if (wp != null) isModal = wp.IsModal;
        }
        catch { isModal = false; }

        var score  = 0;
        var reason = reasonPrefix + ";";

        if (isForeground)  { score += 100; reason += "foreground;"; }
        if (sameProcess)   { score +=  80; reason += "same-process;"; }
        if (isModal)       { score +=  70; reason += "modal;"; }

        if (string.Equals(className, "#32770", StringComparison.OrdinalIgnoreCase))
        { score += 50; reason += "dialog-class;"; }

        if (!string.IsNullOrWhiteSpace(criteria.TitleOrValue))
        {
            var exact = string.Equals(title, criteria.TitleOrValue, StringComparison.OrdinalIgnoreCase);
            score += exact ? 50 : 25;
            reason += "title;";
        }

        if (!string.IsNullOrWhiteSpace(criteria.ClassName))
        { score += 40; reason += "class;"; }

        // Penalize the main application window when no title filter is set.
        try
        {
            var activeHwnd = session?.ActiveWindow != null
                ? SafeWindowHandle(session.ActiveWindow)
                : IntPtr.Zero;

            if (activeHwnd == hwnd && string.IsNullOrWhiteSpace(criteria.TitleOrValue))
            { score -= 40; reason += "active-main-penalty;"; }
        }
        catch { /* ignore */ }

        if (score <= 0) score = 1;

        return new PopupWindowInfo
        {
            Element     = element,
            Hwnd        = hwnd,
            ProcessId   = processId,
            Title       = title,
            ClassName   = className,
            IsModal     = isModal,
            IsForeground = isForeground,
            SameProcess  = sameProcess,
            Score        = score,
            Reason       = reason
        };
    }

    private AutomationSession? TryGetSessionOrNull()
    {
        try { return RequireSession(); }
        catch { return null; }
    }

    private object TopWindow(UiRequest request)
    {
        var session = TryGetSessionOrNull();
        var popup   = FindBestPopupWindow(request, session);

        if (popup == null)
            return new { found = false, message = "No top/popup window found." };

        if (request.MakeCurrent != false)
            MakePopupCurrent(session, popup);

        return new { found = true, window = ToPopupDto(popup) };
    }

    private object WaitForPopup(UiRequest request)
    {
        var session   = TryGetSessionOrNull();
        var timeoutMs = request.TimeoutMs ?? 5000;
        var deadline  = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        PopupWindowInfo? popup = null;
        while (DateTime.UtcNow <= deadline)
        {
            popup = FindBestPopupWindow(request, session);
            if (popup != null) break;
            Thread.Sleep(200);
        }

        if (popup == null)
            return new
            {
                found     = false,
                timeoutMs,
                value     = request.Value,
                className = request.ClassName,
                hwnd      = request.Hwnd
            };

        if (request.MakeCurrent != false)
            MakePopupCurrent(session, popup);

        return new { found = true, timeoutMs, window = ToPopupDto(popup) };
    }

    private object PopupExists(UiRequest request)
    {
        var session   = TryGetSessionOrNull();
        var timeoutMs = request.TimeoutMs ?? 500;
        var deadline  = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        PopupWindowInfo? popup = null;
        while (DateTime.UtcNow <= deadline)
        {
            popup = FindBestPopupWindow(request, session);
            if (popup != null) break;
            Thread.Sleep(100);
        }

        return new
        {
            exists = popup != null,
            window = popup == null ? (object?)null : ToPopupDto(popup)
        };
    }

    private object PopupAction(
        UiRequest request,
        string? defaultAction = null,
        string? defaultButton = null,
        bool requireTarget = false)
    {
        var session = TryGetSessionOrNull();

        var action = string.IsNullOrWhiteSpace(request.Action)
            ? defaultAction ?? "button"
            : request.Action.Trim().ToLowerInvariant();

        if (action is not ("button" or "close" or "enter" or "escape" or "makecurrent"))
            throw new ArgumentException(
                "Unsupported popup action. " +
                "Supported actions: button, close, enter, escape, makecurrent.");

        if (requireTarget &&
            string.IsNullOrWhiteSpace(request.Value) &&
            !request.Hwnd.HasValue &&
            string.IsNullOrWhiteSpace(request.ClassName))
        {
            throw new ArgumentException(
                "'popupok' requires value, hwnd, or className. " +
                "Use alertok for best-detected alert OK.");
        }

        var button = string.IsNullOrWhiteSpace(request.Button)
            ? defaultButton
            : request.Button;

        var popup = WaitForPopupInternal(request, session);

        if (popup == null)
            return new { success = false, action, button, message = "Popup not found." };

        MakePopupCurrent(session, popup);

        var success = action switch
        {
            "close"       => TryClosePopup(popup),
            "enter"       => TrySendEnterToPopup(popup),
            "escape"      => TrySendEscapeToPopup(popup),
            "makecurrent" => true,
            "button"      => TryClickPopupButton(popup.Element, button),
            _             => throw new InvalidOperationException($"Unhandled action '{action}'.")
        };

        return new { success, action, button, window = ToPopupDto(popup) };
    }

    private PopupWindowInfo? WaitForPopupInternal(UiRequest request, AutomationSession? session)
    {
        var timeoutMs = request.TimeoutMs ?? 5000;
        var deadline  = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        while (DateTime.UtcNow <= deadline)
        {
            var popup = FindBestPopupWindow(request, session);
            if (popup != null) return popup;
            Thread.Sleep(200);
        }

        return null;
    }

    private object PopupText(UiRequest request)
    {
        var session = TryGetSessionOrNull();

        var popup = WaitForPopupInternal(request, session);

        if (popup == null)
        {
            return new
            {
                found        = false,
                operation    = "popuptext",
                timeoutMs    = request.TimeoutMs ?? 5000,
                value        = request.Value,
                className    = request.ClassName,
                hwnd         = request.Hwnd,
                message      = string.Empty,
                texts        = Array.Empty<string>(),
                messageTexts = Array.Empty<string>(),
                buttons      = Array.Empty<string>()
            };
        }

        if (request.MakeCurrent != false)
        {
            MakePopupCurrent(session, popup);
            Thread.Sleep(WindowActivationDelayMs);
        }

        var snapshot = ReadPopupTextSnapshot(popup.Element);

        return new
        {
            found        = true,
            operation    = "popuptext",
            message      = snapshot.Message,
            texts        = snapshot.Texts,
            messageTexts = snapshot.MessageTexts,
            buttons      = snapshot.Buttons,
            edits        = snapshot.Edits,
            documents    = snapshot.Documents,
            children     = snapshot.Children,
            window       = ToPopupDto(popup)
        };
    }

    private void MakePopupCurrent(AutomationSession? session, PopupWindowInfo popup)
    {
        try
        {
            if (IsIconic(popup.Hwnd))
                ShowWindow(popup.Hwnd, SW_RESTORE);

            SetForegroundWindow(popup.Hwnd);

            if (session != null)
            {
                try
                {
                    session.ActiveWindow = popup.Element;
                    session.TrackWindow(
                        popup.Hwnd,
                        popup.ProcessId,
                        popup.Title,
                        popup.ClassName,
                        isMainWindow: false);
                }
                catch { /* ignore session update failure */ }
            }
        }
        catch { /* ignore */ }
    }

    private static bool TryClosePopup(PopupWindowInfo popup)
    {
        try
        {
            CloseElement(popup.Element);
            return true;
        }
        catch
        {
            try { return ActivateWin32Window(popup.Hwnd) && PostCloseMessage(popup.Hwnd); }
            catch { return false; }
        }
    }

    private static bool PostCloseMessage(IntPtr hwnd)
    {
        const uint WM_CLOSE = 0x0010;
        return PostWinMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostWinMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private static bool TrySendEnterToPopup(PopupWindowInfo popup)
    {
        try
        {
            popup.Element.Focus();
            Keyboard.Press(VirtualKeyShort.RETURN);
            Keyboard.Release(VirtualKeyShort.RETURN);
            return true;
        }
        catch { return false; }
    }

    private static bool TrySendEscapeToPopup(PopupWindowInfo popup)
    {
        try
        {
            popup.Element.Focus();
            Keyboard.Press(VirtualKeyShort.ESCAPE);
            Keyboard.Release(VirtualKeyShort.ESCAPE);
            return true;
        }
        catch { return false; }
    }

    private static bool TryClickPopupButton(AutomationElement popupRoot, string? buttonSpec)
    {
        var buttonNames = ParseButtonNames(buttonSpec);

        // Collect all visible buttons once to avoid multiple UIA tree traversals.
        AutomationElement[] allButtons;
        try
        {
            allButtons = popupRoot.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
        }
        catch
        {
            allButtons = [];
        }

        // Pass 1: exact UIA name match (as before — fastest and most precise).
        foreach (var name in buttonNames)
        {
            foreach (var btn in allButtons)
            {
                try
                {
                    if (!string.Equals(btn.Name, name, StringComparison.OrdinalIgnoreCase))
                        continue;

                    return InvokeOrClickButton(btn);
                }
                catch { /* try next */ }
            }
        }

        // Pass 2: normalized match — strip access-key prefix '&' and trim whitespace.
        foreach (var name in buttonNames)
        {
            var normalizedTarget = NormalizePopupButtonText(name);

            foreach (var btn in allButtons)
            {
                try
                {
                    if (!string.Equals(NormalizePopupButtonText(btn.Name), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                        continue;

                    return InvokeOrClickButton(btn);
                }
                catch { /* try next */ }
            }
        }

        // Pass 3: contains match, but only when exactly one button matches (to avoid ambiguity).
        foreach (var name in buttonNames)
        {
            var normalizedTarget = NormalizePopupButtonText(name);

            var containsMatches = allButtons
                .Where(btn =>
                {
                    try { return NormalizePopupButtonText(btn.Name).Contains(normalizedTarget, StringComparison.OrdinalIgnoreCase); }
                    catch { return false; }
                })
                .ToList();

            if (containsMatches.Count == 1)
            {
                try { return InvokeOrClickButton(containsMatches[0]); }
                catch { /* try next name */ }
            }
        }

        return false;
    }

    private static bool InvokeOrClickButton(AutomationElement btn)
    {
        if (btn.Patterns.Invoke.PatternOrDefault != null)
        {
            btn.Patterns.Invoke.Pattern.Invoke();
            return true;
        }

        btn.Click();
        return true;
    }

    /// <summary>
    /// Strips the Win32 access-key prefix (<c>&amp;</c>) from a button label and trims whitespace,
    /// so that "&amp;OK" and "OK" compare equal.
    /// </summary>
    private static string NormalizePopupButtonText(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return string.Empty;

        return name.Replace("&", string.Empty).Trim();
    }

    private static IReadOnlyList<string> ParseButtonNames(string? buttonSpec)
    {
        if (string.IsNullOrWhiteSpace(buttonSpec))
            return ["OK", "Yes", "Save"];

        return buttonSpec
            .Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
    }

    // -------------------------------------------------------------------------
    // Popup text snapshot
    // -------------------------------------------------------------------------

    private sealed class PopupTextSnapshot
    {
        public string Message { get; init; } = string.Empty;
        public List<string> Texts { get; init; } = new();
        public List<string> MessageTexts { get; init; } = new();
        public List<string> Buttons { get; init; } = new();
        public List<string> Edits { get; init; } = new();
        public List<string> Documents { get; init; } = new();
        public List<object> Children { get; init; } = new();
    }

    private static PopupTextSnapshot ReadPopupTextSnapshot(AutomationElement popupRoot)
    {
        var allTextFragments = new List<string>();
        var messageTexts     = new List<string>();
        var buttons          = new List<string>();
        var edits            = new List<string>();
        var documents        = new List<string>();
        var children         = new List<object>();

        AutomationElement[] descendants;
        try
        {
            descendants = popupRoot.FindAllDescendants();
        }
        catch
        {
            descendants = [];
        }

        foreach (var child in descendants)
        {
            var name        = SafeElementName(child);
            var automationId = SafeElementAutomationId(child);
            var className   = SafeElementClassName(child);
            var controlType = SafeElementControlType(child);
            var value       = SafeElementValue(child);

            var displayText = FirstNonEmpty(value, name);

            if (!string.IsNullOrWhiteSpace(displayText))
                allTextFragments.Add(displayText);

            if (IsPopupMessageCandidate(controlType, name, value))
            {
                var messageText = FirstNonEmpty(value, name);
                if (!string.IsNullOrWhiteSpace(messageText))
                    messageTexts.Add(messageText);
            }

            if (string.Equals(controlType, "Button", StringComparison.OrdinalIgnoreCase))
            {
                var buttonText = FirstNonEmpty(name, value);
                if (!string.IsNullOrWhiteSpace(buttonText))
                    buttons.Add(buttonText);
            }

            if (string.Equals(controlType, "Edit", StringComparison.OrdinalIgnoreCase))
            {
                var editText = FirstNonEmpty(value, name);
                if (!string.IsNullOrWhiteSpace(editText))
                    edits.Add(editText);
            }

            if (string.Equals(controlType, "Document", StringComparison.OrdinalIgnoreCase))
            {
                var docText = FirstNonEmpty(value, name);
                if (!string.IsNullOrWhiteSpace(docText))
                    documents.Add(docText);
            }

            children.Add(new
            {
                name,
                value,
                automationId,
                className,
                controlType,
                rectangle  = SafeBoundingRectangleObject(child),
                isOffscreen = SafeIsOffscreen(child),
                isEnabled  = SafeIsEnabled(child)
            });
        }

        var cleanTexts      = DistinctCleanText(allTextFragments);
        var cleanButtons    = DistinctCleanText(buttons);

        bool IsNotButton(string x) => !cleanButtons.Contains(x, StringComparer.OrdinalIgnoreCase);

        var cleanMessageTexts = DistinctCleanText(messageTexts).Where(IsNotButton).ToList();

        if (cleanMessageTexts.Count == 0)
            cleanMessageTexts = cleanTexts.Where(IsNotButton).ToList();

        var message = string.Join(" ", cleanMessageTexts);

        return new PopupTextSnapshot
        {
            Message      = message,
            Texts        = cleanTexts,
            MessageTexts = cleanMessageTexts,
            Buttons      = cleanButtons,
            Edits        = DistinctCleanText(edits),
            Documents    = DistinctCleanText(documents),
            Children     = children
        };
    }

    internal static string SafeElementValue(AutomationElement element)
    {
        try
        {
            var valuePattern = element.Patterns.Value.PatternOrDefault;
            if (valuePattern != null)
            {
                var val = valuePattern.Value.ValueOrDefault;
                if (!string.IsNullOrWhiteSpace(val))
                    return val.Trim();
            }
        }
        catch { /* ignore */ }

        try
        {
            var textPattern = element.Patterns.Text.PatternOrDefault;
            if (textPattern != null)
            {
                var text = textPattern.DocumentRange.GetText(-1);
                if (!string.IsNullOrWhiteSpace(text))
                    return text.Trim();
            }
        }
        catch { /* ignore */ }

        return string.Empty;
    }

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var v in values)
        {
            if (!string.IsNullOrWhiteSpace(v))
                return v.Trim();
        }
        return string.Empty;
    }

    private static List<string> DistinctCleanText(IEnumerable<string> values)
    {
        return values
            .Select(NormalizePopupText)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizePopupText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return Regex.Replace(
            value.Replace("\r", " ").Replace("\n", " "),
            @"\s+",
            " ").Trim();
    }

    private static bool IsPopupMessageCandidate(string controlType, string name, string value)
    {
        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(value))
            return false;

        if (string.Equals(controlType, "Button", StringComparison.OrdinalIgnoreCase))
            return false;

        return string.Equals(controlType, "Text",     StringComparison.OrdinalIgnoreCase) ||
               string.Equals(controlType, "Edit",     StringComparison.OrdinalIgnoreCase) ||
               string.Equals(controlType, "Document", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(controlType, "Custom",   StringComparison.OrdinalIgnoreCase);
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
        var header = ResolveElementForOperation(
            req,
            purpose: "openheaderdropdown",
            action: true,
            allowOffscreen: false,
            requireClickable: true);
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

        var items = GetDropdownSelectableItems(list)
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
        if (string.IsNullOrWhiteSpace(req.Value) && !req.Index.HasValue)
            throw new ArgumentException("value or index is required for selectheaderdropdownitem.");

        var resolved = ResolveElementResultForOperation(
            req,
            purpose: "selectheaderdropdownitem",
            action: true,
            allowOffscreen: false,
            requireClickable: true);
        var header = resolved.Element;

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
        var matchMode = req.MatchMode ?? req.Locator?.MatchMode ?? "contains";

        var container = OpenHeaderDropdownAndFindContainer(header, region);
        if (container == null)
            throw new InvalidOperationException("Header dropdown container was not found after opening.");

        var selectableItems = GetDropdownSelectableItems(container);
        var matched = FindDropdownSelectableItem(selectableItems, req.Value, req.Index, matchMode);
        if (matched == null)
        {
            var available = BuildDropdownItemSummaries(selectableItems);
            throw new InvalidOperationException(
                $"Dropdown item '{req.Value ?? req.Index?.ToString()}' was not found. Available: {string.Join(", ", available.Select(i => i.name))}");
        }

        var activation = ActivateOpenDropdownItemSoft(matched, req.Value ?? SafeElementName(matched), itemRegion);
        var availableItems = BuildDropdownItemSummaries(selectableItems);

        return new
        {
            operation = "selectheaderdropdownitem",
            success = activation.Success,
            requested = req.Value ?? req.Index?.ToString(),
            matchedText = activation.MatchedText,
            activationStrategy = activation.ActivationStrategy,
            verified = activation.Verified,
            verificationReason = activation.VerificationReason,
            header = SafeElementName(header),
            headerRegion = region.ToString(),
            itemRegion = itemRegion.ToString(),
            container = activation.Container,
            matchedItem = activation.MatchedItem,
            availableItems
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

        var parentMenuItem = ResolveElementForOperation(req, "selectdynamicmenupath");
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

    /// <summary>
    /// Backward-compatible alias. The canonical operation is <c>select</c>.
    /// Delegates directly to <see cref="Select"/> so both code paths share
    /// the identical ComboBox selection pipeline.
    /// </summary>
    private object? SelectComboBoxItem(UiRequest req, CancellationToken cancellationToken = default)
    {
        return Select(req, cancellationToken);
    }

    /// <summary>
    /// Core ComboBox value-selection pipeline, shared by <c>select</c> and
    /// the legacy <c>selectcomboboxitem</c> alias.
    /// </summary>
    private object? SelectComboBoxByValue(
        AutomationSession session,
        AutomationElement comboBox,
        string itemName,
        UiRequest req)
    {
        var guard = CaptureComboBoxTargetGuard(comboBox);
        var operationDeadline = DateTime.UtcNow.AddMilliseconds(ComboBoxSingleSelectionTimeoutMs);

        _logger.LogInformation(
            "ComboBox selection starting. Trying Win32 native strategy first. combo={Combo}, value={Value}",
            SafeElementName(comboBox),
            itemName);

        // --- Strategy 1: Win32 native (pywinauto-style, no dropdown) ---
        if (TrySelectNativeWin32ComboBox(session, comboBox, itemName, out var win32NativeStrategy))
        {
            _logger.LogInformation(
                "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
                win32NativeStrategy,
                itemName,
                GetComboBoxCurrentValue(session, comboBox));

            return new
            {
                selected = itemName,
                actual   = GetComboBoxCurrentValue(session, comboBox),
                comboBox = SafeElementName(comboBox),
                verified = true,
                strategy = win32NativeStrategy,
                operation = "select"
            };
        }

        _logger.LogInformation(
            "ComboBox Win32 native strategy did not apply. Falling back to UIA direct strategy. combo={Combo}, value={Value}",
            SafeElementName(comboBox),
            itemName);

        var probedCapabilities = ProbeComboBoxDropdownItemCapabilities(
            session,
            comboBox,
            itemName,
            guard,
            operationDeadline);
        if (probedCapabilities.HasAnyUsefulPattern &&
            TrySelectComboBoxByDirectUia(
                session,
                comboBox,
                itemName,
                guard,
                operationDeadline,
                out var directStrategy))
        {
            _logger.LogInformation(
                "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
                directStrategy,
                itemName,
                GetComboBoxCurrentValue(session, comboBox));

            return new
            {
                selected = itemName,
                actual = GetComboBoxCurrentValue(session, comboBox),
                comboBox = SafeElementName(comboBox),
                verified = true,
                strategy = directStrategy,
                operation = "select"
            };
        }

        if (!probedCapabilities.HasAnyUsefulPattern)
        {
            _logger.LogInformation(
                "ComboBox pattern capability decision. combo={Combo}, capabilities={Capabilities}, decision={Decision}",
                SafeElementName(comboBox),
                probedCapabilities.ToString(),
                "Fallback to visual search because no useful ListItem patterns were detected");
        }

        if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
            !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "reopen"))
        {
            throw new InvalidOperationException($"ComboBox selection aborted for '{itemName}'.");
        }

        if (!OpenComboBoxDropdown(session, comboBox))
            throw new InvalidOperationException($"Failed to open ComboBox '{SafeElementName(comboBox)}'.");

        Thread.Sleep(MenuExpandDelayMs);

        if (IsHugeComboBoxDropdown(session, comboBox))
        {
            _logger.LogInformation(
                "ComboBox detected as huge list. Using paged visible-list search. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                itemName);

            if (TrySelectComboBoxByPagedVisibleSearch(session, comboBox, itemName, guard, operationDeadline))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                _logger.LogInformation(
                    "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
                    "huge-list-paged-visible-search",
                    itemName,
                    actual);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "huge-list-paged-visible-search",
                    operation = "select"
                };
            }

            _logger.LogInformation(
                "Huge ComboBox paged visible-list search failed. Trying visible anchor-window search. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                itemName);

            if (TrySelectComboBoxByVisibleAnchorWindowSearch(session, comboBox, itemName, guard, operationDeadline))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                _logger.LogInformation(
                    "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
                    "huge-list-visible-anchor-window-search",
                    itemName,
                    actual);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "huge-list-visible-anchor-window-search",
                    operation = "select"
                };
            }

            if (req.AllowKeyboardFallback == true)
            {
                _logger.LogInformation(
                    "Huge ComboBox visible anchor-window search failed. allowKeyboardFallback=true, trying type-ahead. combo={Combo}, value={Value}",
                    SafeElementName(comboBox),
                    itemName);

                if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) &&
                    IsComboBoxTargetGuardValid(comboBox, guard, itemName, "keyboard fallback") &&
                    TrySelectComboBoxByKeyboardSafe(session, comboBox, itemName, guard, operationDeadline))
                {
                    var actual = GetComboBoxCurrentValue(session, comboBox);

                    _logger.LogInformation(
                        "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
                        "huge-list-explicit-typeahead-fallback",
                        itemName,
                        actual);

                    return new
                    {
                        selected = itemName,
                        actual,
                        comboBox = SafeElementName(comboBox),
                        verified = true,
                        strategy = "huge-list-explicit-typeahead-fallback",
                        operation = "select"
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

        var smallOperationDeadline = DateTime.UtcNow.AddMilliseconds(SmallComboBoxSelectionTimeoutMs);
        if (!CommitSmallComboBoxItemUserLike(
                session,
                comboBox,
                itemName,
                guard,
                smallOperationDeadline))
        {
            _logger.LogWarning(
                "Small ComboBox user-like commit did not succeed. Falling back to visual search strategies. requested={Requested}, combo={Combo}",
                itemName,
                SafeElementName(comboBox));

            if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) &&
                IsComboBoxTargetGuardValid(comboBox, guard, itemName, "small-combobox paged visible search fallback") &&
                TrySelectComboBoxByPagedVisibleSearch(session, comboBox, itemName, guard, operationDeadline))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "small-combobox-paged-visible-search-fallback",
                    operation = "select"
                };
            }

            if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) &&
                IsComboBoxTargetGuardValid(comboBox, guard, itemName, "small-combobox anchor-window search fallback") &&
                TrySelectComboBoxByVisibleAnchorWindowSearch(session, comboBox, itemName, guard, operationDeadline))
            {
                var actual = GetComboBoxCurrentValue(session, comboBox);

                return new
                {
                    selected = itemName,
                    actual,
                    comboBox = SafeElementName(comboBox),
                    verified = true,
                    strategy = "small-combobox-anchor-window-search-fallback",
                    operation = "select"
                };
            }

            var actualAfterCommit = GetComboBoxCurrentValue(session, comboBox);

            _logger.LogWarning(
                "Small ComboBox item did not commit or stabilize after all fallback strategies. requested={Requested}, actual={Actual}, combo={Combo}",
                itemName,
                actualAfterCommit,
                SafeElementName(comboBox));

            throw new InvalidOperationException(
                $"ComboBox item '{itemName}' was visible/found but did not commit. Actual='{actualAfterCommit}'.");
        }

        var actualAfterVerifiedCommit = GetComboBoxCurrentValue(session, comboBox);

        _logger.LogInformation(
            "ComboBox selected by value. operation=select, strategy={Strategy}, requested={Requested}, actual={Actual}",
            "small-combobox-exact-visible-commit",
            itemName,
            actualAfterVerifiedCommit);

        return new
        {
            selected = itemName,
            actual = actualAfterVerifiedCommit,
            comboBox = SafeElementName(comboBox),
            verified = true,
            strategy = "small-combobox-exact-visible-commit",
            operation = "select"
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

        var element = ResolveElementForOperation(
            req,
            purpose: doubleClick ? "doubleclickgridcell" : "clickgridcell",
            action: true,
            allowOffscreen: false,
            requireClickable: false);

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

    // =========================================================================
    // Operation performance policy
    // =========================================================================

    private sealed class UiOperationPolicy
    {
        public TimeSpan Timeout { get; init; }
        public TimeSpan RetryInterval { get; init; }
        public bool AllowDesktopPopupScan { get; init; }
        public bool RefreshRootEveryRetry { get; init; }
        public bool UseElementCache { get; init; }
        public string PolicyName { get; init; } = "default";
    }

    private UiOperationPolicy GetOperationPolicy(UiRequest req)
    {
        var op = req.Operation.ToLowerInvariant();

        if (req.TimeoutMs.HasValue)
        {
            return new UiOperationPolicy
            {
                Timeout = TimeSpan.FromMilliseconds(req.TimeoutMs.Value),
                RetryInterval = TimeSpan.FromMilliseconds(req.Fast == true ? 100 : 500),
                AllowDesktopPopupScan = req.DisableAutoFollow != true && IsSlowWindowOperation(op),
                RefreshRootEveryRetry = IsSlowWindowOperation(op),
                UseElementCache = req.UseCache == true || req.Fast == true,
                PolicyName = req.Fast == true ? "request-fast-override" : "request-timeout-override"
            };
        }

        if (IsFastElementOperation(op))
        {
            return new UiOperationPolicy
            {
                Timeout = TimeSpan.FromMilliseconds(1000),
                RetryInterval = TimeSpan.FromMilliseconds(100),
                AllowDesktopPopupScan = false,
                RefreshRootEveryRetry = false,
                UseElementCache = true,
                PolicyName = "fast-element-operation"
            };
        }

        if (IsSlowWindowOperation(op))
        {
            return new UiOperationPolicy
            {
                Timeout = TimeSpan.FromMilliseconds(8000),
                RetryInterval = TimeSpan.FromMilliseconds(300),
                AllowDesktopPopupScan = true,
                RefreshRootEveryRetry = true,
                UseElementCache = false,
                PolicyName = "slow-window-operation"
            };
        }

        return new UiOperationPolicy
        {
            Timeout = DefaultRetry,
            RetryInterval = RetryInterval,
            AllowDesktopPopupScan = true,
            RefreshRootEveryRetry = true,
            UseElementCache = false,
            PolicyName = "default-stable"
        };
    }

    private static bool IsFastElementOperation(string op)
    {
        return op is
            "click" or
            "type" or
            "sendkeys" or
            "clear" or
            "gettext" or
            "getvalue" or
            "setvalue" or
            "check" or
            "uncheck" or
            "select" or
            "typedate" or
            "rightclick" or
            "doubleclick" or
            "findall" or
            "findlocator" or
            "inspectlocator";
    }

    private static bool IsSlowWindowOperation(string op)
    {
        return op is
            "switchwindow" or
            "waitforwindow" or
            "waitforpopup" or
            "topwindow" or
            "popupaction" or
            "alertok" or
            "alertcancel" or
            "alertclose" or
            "popupok" or
            "popuptext" or
            "alerttext" or
            "readpopup" or
            "contextmenupath" or
            "openheaderdropdown" or
            "selectheaderdropdownitem" or
            "selectopendropdownitem" or
            "clickopendropdownitem" or
            "gettable" or
            "gettableheaders" or
            "selectdynamicmenuitem" or
            "selectdynamicmenupath";
    }

    /// <summary>
    /// Returns <c>true</c> when attribute-based search (AutomationId, Name, ClassName, ControlType)
    /// should be tried before XPath for the given request.
    /// Fast element operations use this by default; callers can opt in or out explicitly.
    /// </summary>
    private bool ShouldPreferAttributeSearch(UiRequest req)
    {
        if (req.XPathOnly == true)
            return false;

        if (req.PreferAttributes == true)
            return true;

        if (req.PreferXPath == true)
            return false;

        var policy = GetOperationPolicy(req);
        return policy.PolicyName == "fast-element-operation" || req.Fast == true;
    }

    /// <summary>
    /// Returns <c>true</c> when XPath should be the primary (or sole) search strategy.
    /// </summary>
    private bool ShouldUseXPathFirst(UiRequest req)
    {
        if (req.XPathOnly == true)
            return true;

        if (req.PreferXPath == true)
            return true;

        return !ShouldPreferAttributeSearch(req);
    }

    /// <summary>
    /// Returns <c>true</c> when the locator has at least one non-XPath attribute that
    /// can be used for a focused UIA tree query (AutomationId, Name, ClassName, or ControlType).
    /// </summary>
    private static bool HasUsableAttributeLocator(UiLocator locator)
    {
        return !string.IsNullOrWhiteSpace(locator.AutomationId) ||
               !string.IsNullOrWhiteSpace(locator.Name) ||
               !string.IsNullOrWhiteSpace(locator.ClassName) ||
               !string.IsNullOrWhiteSpace(locator.ControlType);
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
    /// <param name="session">The active automation session.</param>
    /// <param name="allowDesktopPopupScan">
    /// When <c>true</c> (default), performs the full desktop descendant scan and popup
    /// auto-follow logic.  Set to <c>false</c> for fast element operations (click, type,
    /// clear, etc.) to skip the expensive desktop traversal.
    /// </param>
    private AutomationElement GetWindowRoot(AutomationSession session, bool allowDesktopPopupScan = true)
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

                    // Register the newly-appeared window in the ownership tracker so it
                    // will be closed when the session ends.
                    var newHwnd = SafeWindowHandle(newWindow);
                    if (newHwnd != IntPtr.Zero)
                        session.TrackWindow(
                            newHwnd,
                            SafeProcessId(newWindow) ?? session.Application.ProcessId,
                            SafeElementName(newWindow),
                            SafeElementClassName(newWindow),
                            isMainWindow: false);
                }
                else if (allowDesktopPopupScan)
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

                                // Track the popup so it will be closed during cleanup.
                                var popupHwnd = SafeWindowHandle(newPopup);
                                if (popupHwnd != IntPtr.Zero)
                                    session.TrackWindow(
                                        popupHwnd,
                                        SafeProcessId(newPopup) ?? session.Application.ProcessId,
                                        SafeElementName(newPopup),
                                        SafeElementClassName(newPopup),
                                        isMainWindow: false);
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

    internal static bool IsEmptyLocator(UiLocator? locator)
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
    internal static bool IsElementAlive(AutomationElement element)
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

    private static string BuildElementCacheKey(
        AutomationSession session,
        AutomationElement root,
        UiLocator locator,
        UiLocator? parentLocator = null,
        bool preferAttributes = false,
        bool xpathOnly = false,
        bool preferXPath = false)
    {
        var rootHandle = SafeWindowHandle(root).ToInt64();
        var parentKey = parentLocator == null ? "" : DescribeLocator(parentLocator);
        var strategyKey = xpathOnly ? "xo" : preferXPath ? "px" : preferAttributes ? "pa" : "def";
        var locatorKey = string.Join(
            ";",
            locator.XPath ?? "",
            locator.Name ?? "",
            locator.AutomationId ?? "",
            locator.ClassName ?? "",
            locator.ControlType ?? "");
        return string.Join("|", "element", session.SessionId, rootHandle, parentKey, strategyKey, locatorKey);
    }

    private static bool TryGetCachedElement(
        string cacheKey,
        UiLocator locator,
        out AutomationElement? cached)
    {
        lock (ElementCacheLock)
        {
            if (ElementCache.TryGetValue(cacheKey, out var entry))
            {
                if (entry.ExpiresAt > DateTime.UtcNow && IsElementAlive(entry.Element))
                {
                    cached = entry.Element;
                    return true;
                }

                ElementCache.Remove(cacheKey);
            }
        }

        cached = null;
        return false;
    }

    private static void StoreCachedElement(string cacheKey, AutomationElement element)
    {
        lock (ElementCacheLock)
            ElementCache[cacheKey] = new CachedElementMatch(element, DateTime.UtcNow + ElementCacheDuration);
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
    /// Finds an element using a locator, applying the operation-based performance policy
    /// (<see cref="GetOperationPolicy"/>) to determine timeout, retry interval, and whether
    /// to perform desktop popup scanning on each retry.
    /// <para>
    /// For fast element operations (click, type, clear, etc.) the root is fetched once and
    /// reused across retries with a short timeout, avoiding the expensive desktop scan.
    /// For slow window operations the root is re-evaluated on every iteration so that newly
    /// opened dialogs are picked up by the auto-follow logic inside <see cref="GetWindowRoot"/>.
    /// </para>
    /// <para>
    /// When <see cref="UiRequest.ParentLocator"/> is set, the parent element is resolved
    /// first and the child <see cref="UiRequest.Locator"/> is searched within that narrower scope.
    /// Set <see cref="UiRequest.FallbackToWindowRootIfParentChildNotFound"/> to <c>true</c>
    /// to allow a full-window retry when the child is not found inside the parent.
    /// </para>
    /// </summary>
    private AutomationElement ResolveElementForOperation(
        UiRequest request,
        string purpose,
        bool action = false,
        bool allowOffscreen = true,
        bool requireClickable = false)
    {
        var resolved = ResolveElementResultForOperation(
            request,
            purpose,
            action,
            allowOffscreen,
            requireClickable);

        return resolved.Element;
    }

    private DesktopAutomationDriver.Models.Resolver.ResolvedElement ResolveElementResultForOperation(
        UiRequest request,
        string purpose,
        bool action = false,
        bool allowOffscreen = true,
        bool requireClickable = false)
    {
        var locator = request.Locator ?? new UiLocator();
        if (!allowOffscreen)
        {
            locator.IncludeOffscreen = false;
        }

        var matchResult = _newResolver.ResolveOne(locator, request, purpose);
        var resolved = new DesktopAutomationDriver.Models.Resolver.ResolvedElement
        {
            Element = matchResult.Element,
            Strategy = matchResult.Strategy,
            Score = matchResult.Candidates.FirstOrDefault()?.Score ?? 100,
            Index = 0
        };

        if (action && requireClickable)
        {
            ValidateActionTarget(resolved.Element, purpose);
        }

        return resolved;
    }

    private void ValidateActionTarget(AutomationElement? element, string purpose)
    {
        if (element == null)
        {
            throw new InvalidOperationException($"Action target element is null for operation: {purpose}");
        }
        if (SafeIsOffscreen(element) == true)
        {
            throw new InvalidOperationException($"Target element is offscreen and cannot receive action: {purpose}");
        }
    }

    // Helper required by specification for central resolver options.
    private ResolveOptions BuildResolveOptionsForOperation(
        UiRequest request,
        string purpose,
        bool action,
        bool allowOffscreen,
        bool requireClickable)
    {
        return new ResolveOptions
        {
            Purpose = purpose,
            Action = action,
            AllowOffscreen = allowOffscreen,
            RequireClickable = requireClickable,
            TimeoutMs = request.TimeoutMs ?? 5000,
            ReturnCandidates = request.ReturnCandidates == true,
            IncludeDiagnostics = request.IncludeDiagnostics == true || request.Debug == true,
            Ambiguity = request.Ambiguity ?? "error",
            SearchRoot = request.SearchRoot ?? "currentWindow",
            TreeView = request.TreeView ?? "control",
            Backend = request.Backend ?? "uia",
            MaxCandidates = request.MaxMatches ?? 100
        };
    }

    [Obsolete("Use ResolveElementForOperation or _uiElementResolver directly. This method will be removed in a future version.")]
    private AutomationElement FindWithRetry(UiRequest req)
    {
        return ResolveElementForOperation(
            req,
            purpose: req.Operation,
            action: false,
            allowOffscreen: true,
            requireClickable: false);
    }

    /// <summary>
    /// Finds an element by <paramref name="locator"/> with up to 5 s retry.
    /// Used directly when a locator is already in hand (e.g. locator2 lookups).
    /// </summary>
    private AutomationElement FindLocatorWithRetry(
        AutomationSession session, AutomationElement root, UiLocator locator)
    {
        var request = new UiRequest { Locator = locator };
        var resolved = _newResolver.ResolveOne(locator, request, "findlocatorwithretry");
        return resolved.Element;
    }

    /// <summary>
    /// Attempts a single element search using the legacy XPath-first strategy.
    /// Used by <see cref="FindLocatorWithRetry"/> (which has no <see cref="UiRequest"/>
    /// context) and as a fallback when neither strategy flag is set.
    /// When XPath is present it is tried first; attribute-priority search follows.
    /// </summary>
    private AutomationElement? TryFindElement(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(locator.XPath))
            {
                var byXPath = FindByXPath(root, session, locator.XPath);
                if (byXPath != null)
                    return byXPath;
            }

            return TryFindElementByFastAttributes(root, session, locator).Element;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Attribute-condition-based element search (no XPath).
    /// Builds a combined AND condition from all non-empty locator properties and
    /// also handles WinForms <c>DateTimePicker</c> partial class-name matching.
    /// </summary>
    private static AutomationElement? TryFindElementByCondition(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        var condition = BuildCondition(session, locator);
        var element = root.FindFirstDescendant(condition);
        if (element != null)
            return element;

        return TryFindWinFormsDateTimePickerByPartialClassName(root, session, locator);
    }

    /// <summary>
    /// Dispatches element search based on the caller's strategy flags.
    /// <list type="bullet">
    ///   <item><term><paramref name="xpathOnly"/>=true</term><description>Only XPath is tried.</description></item>
    ///   <item><term><paramref name="preferAttributes"/>=true</term><description>Attribute search runs first; XPath is the fallback.</description></item>
    ///   <item><term>default</term><description>XPath first (legacy behavior), then attribute conditions.</description></item>
    /// </list>
    /// </summary>
    internal static LocatorSearchResult TryFindElementBySmartStrategy(
        AutomationElement root,
        AutomationSession session,
        UiLocator locator,
        bool preferAttributes,
        bool xpathOnly)
    {
        if (xpathOnly)
        {
            if (string.IsNullOrWhiteSpace(locator.XPath))
                return LocatorSearchResult.NotFound("xpath-only-no-xpath");

            var byXPath = FindByXPath(root, session, locator.XPath);
            return byXPath != null
                ? LocatorSearchResult.Found(byXPath, "xpath-only")
                : LocatorSearchResult.NotFound("xpath-only");
        }

        if (preferAttributes)
        {
            var byAttributes = TryFindElementByFastAttributes(root, session, locator);
            if (byAttributes.Element != null)
                return byAttributes;

            if (!string.IsNullOrWhiteSpace(locator.XPath))
            {
                var byXPath = FindByXPath(root, session, locator.XPath);
                if (byXPath != null)
                    return LocatorSearchResult.Found(byXPath, "xpath-fallback");
            }

            return LocatorSearchResult.NotFound("attribute-first-not-found");
        }

        // XPath-first (legacy / stable)
        if (!string.IsNullOrWhiteSpace(locator.XPath))
        {
            var byXPath = FindByXPath(root, session, locator.XPath);
            if (byXPath != null)
                return LocatorSearchResult.Found(byXPath, "xpath-first");
        }

        var fallbackAttributes = TryFindElementByFastAttributes(root, session, locator);
        if (fallbackAttributes.Element != null)
            return fallbackAttributes;

        return LocatorSearchResult.NotFound("not-found");
    }

    /// <summary>
    /// Searches for an element using the prioritised attribute strategy:
    /// <list type="number">
    ///   <item>AutomationId + ControlType</item>
    ///   <item>AutomationId only</item>
    ///   <item>Name + ControlType</item>
    ///   <item>Name + ClassName</item>
    ///   <item>ControlType + ClassName</item>
    ///   <item>Full condition-based fallback (all attributes AND-ed)</item>
    /// </list>
    /// </summary>
    internal static LocatorSearchResult TryFindElementByFastAttributes(
        AutomationElement root,
        AutomationSession session,
        UiLocator locator)
    {
        try
        {
            if (!HasUsableAttributeLocator(locator))
                return LocatorSearchResult.NotFound("no-attribute-locator");

            // 1. AutomationId + ControlType
            if (!string.IsNullOrWhiteSpace(locator.AutomationId) &&
                !string.IsNullOrWhiteSpace(locator.ControlType))
            {
                var match = FindFirstByAutomationIdAndControlType(root, session, locator);
                if (match != null && LocatorMatchesExceptXPath(match, locator))
                    return LocatorSearchResult.Found(match, "automationid-controltype");
            }

            // 2. AutomationId only
            if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            {
                var cf = session.Automation.ConditionFactory;
                var match = root.FindFirstDescendant(cf.ByAutomationId(locator.AutomationId));
                if (match != null && LocatorMatchesExceptXPath(match, locator))
                    return LocatorSearchResult.Found(match, "automationid");
            }

            // 3. Name + ControlType
            if (!string.IsNullOrWhiteSpace(locator.Name) &&
                !string.IsNullOrWhiteSpace(locator.ControlType))
            {
                var match = FindFirstByNameAndControlType(root, session, locator);
                if (match != null && LocatorMatchesExceptXPath(match, locator))
                    return LocatorSearchResult.Found(match, "name-controltype");
            }

            // 4. Name + ClassName
            if (!string.IsNullOrWhiteSpace(locator.Name) &&
                !string.IsNullOrWhiteSpace(locator.ClassName))
            {
                var cf = session.Automation.ConditionFactory;
                var match = root.FindFirstDescendant(
                    new AndCondition(cf.ByName(locator.Name), cf.ByClassName(locator.ClassName)));
                if (match != null && LocatorMatchesExceptXPath(match, locator))
                    return LocatorSearchResult.Found(match, "name-classname");
            }

            // 5. ControlType + ClassName
            if (!string.IsNullOrWhiteSpace(locator.ControlType) &&
                !string.IsNullOrWhiteSpace(locator.ClassName))
            {
                var cf = session.Automation.ConditionFactory;
                var controlType = ParseControlType(locator.ControlType);
                var match = root.FindFirstDescendant(
                    new AndCondition(cf.ByControlType(controlType), cf.ByClassName(locator.ClassName)));
                if (match != null && LocatorMatchesExceptXPath(match, locator))
                    return LocatorSearchResult.Found(match, "controltype-classname");
            }

            // 6. Full condition-based fallback
            var fallback = TryFindElementByCondition(root, session, locator);
            if (fallback != null)
                return LocatorSearchResult.Found(fallback, "condition-fallback");

            return LocatorSearchResult.NotFound("attributes-not-found");
        }
        catch
        {
            return LocatorSearchResult.NotFound("attributes-exception");
        }
    }

    /// <summary>
    /// Finds the first descendant matching both <c>AutomationId</c> and <c>ControlType</c>
    /// from <paramref name="locator"/>.
    /// </summary>
    private static AutomationElement? FindFirstByAutomationIdAndControlType(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        var cf = session.Automation.ConditionFactory;
        var controlType = ParseControlType(locator.ControlType!);
        return root.FindFirstDescendant(
            new AndCondition(
                cf.ByAutomationId(locator.AutomationId!),
                cf.ByControlType(controlType)));
    }

    /// <summary>
    /// Finds the first descendant matching both <c>Name</c> and <c>ControlType</c>
    /// from <paramref name="locator"/>.
    /// </summary>
    private static AutomationElement? FindFirstByNameAndControlType(
        AutomationElement root, AutomationSession session, UiLocator locator)
    {
        var cf = session.Automation.ConditionFactory;
        var controlType = ParseControlType(locator.ControlType!);
        return root.FindFirstDescendant(
            new AndCondition(
                cf.ByName(locator.Name!),
                cf.ByControlType(controlType)));
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

    /// <summary>
    /// Returns <c>true</c> when all non-XPath locator properties that are present
    /// (Name, AutomationId, ClassName, ControlType) match the element.
    /// Used by fast attribute search to verify candidates that were found by a
    /// single-attribute query.
    /// </summary>
    private static bool LocatorMatchesExceptXPath(AutomationElement element, UiLocator locator)
    {
        if (!string.IsNullOrWhiteSpace(locator.Name) &&
            !string.Equals(SafeElementName(element), locator.Name, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.AutomationId) &&
            !string.Equals(SafeElementAutomationId(element), locator.AutomationId, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ClassName) &&
            !string.Equals(SafeElementClassName(element), locator.ClassName, StringComparison.OrdinalIgnoreCase))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ControlType) &&
            !string.Equals(element.ControlType.ToString(), locator.ControlType, StringComparison.OrdinalIgnoreCase))
            return false;

        return true;
    }

    // =========================================================================
    // Win32 window discovery model
    // =========================================================================

    /// <summary>
    /// Lightweight record returned by the Win32 top-level window enumerator,
    /// carrying the handle, owning process identifier, and window title.
    /// </summary>
    private sealed class Win32WindowInfo
    {
        public IntPtr Hwnd { get; init; }
        public int ProcessId { get; init; }
        public string Title { get; init; } = string.Empty;
        public string ClassName { get; init; } = string.Empty;
    }

    /// <summary>
    /// Resolved criteria used to search for a Win32 top-level window.
    /// Built from a <see cref="UiRequest"/> by <c>BuildWin32SwitchCriteria</c>.
    /// </summary>
    private sealed class Win32SwitchCriteria
    {
        public string? TitleOrValue { get; init; }
        public IntPtr? Hwnd { get; init; }
        public string? ClassName { get; init; }
        public int? ProcessId { get; init; }
        public string MatchMode { get; init; } = "contains";
    }

    // =========================================================================
    // Locator search result
    // =========================================================================

    /// <summary>
    /// Carries the outcome of a smart locator search together with the strategy
    /// name that produced the result.  Strategy names are used in structured log
    /// messages so callers can tell at a glance which path was taken.
    /// </summary>
    internal sealed class LocatorSearchResult
    {
        public AutomationElement? Element { get; init; }

        /// <summary>
        /// Identifies which search path found (or failed to find) the element.
        /// Common values: cache, parent-scope, automationid-controltype, automationid,
        /// name-controltype, name-classname, controltype-classname, xpath-first,
        /// xpath-fallback, xpath-only, condition-fallback, not-found, attributes-exception.
        /// </summary>
        public string Strategy { get; init; } = "not-found";

        public static LocatorSearchResult Found(AutomationElement element, string strategy)
            => new() { Element = element, Strategy = strategy };

        public static LocatorSearchResult NotFound(string strategy)
            => new() { Element = null, Strategy = strategy };
    }

    // =========================================================================
    // pywinauto-style centralized resolver
    // =========================================================================

    private ElementResolution.ElementResolveResult PywinautoResolve(
        UiRequest request,
        ElementResolution.ResolveOptions options,
        CancellationToken cancellationToken = default)
    {
        var timeoutMs = request.TimeoutMs ?? 5000;
        var pollMs = request.PollIntervalMs ?? 200;

        if (UsePywinautoStyleResolver)
        {
            var result = _pywinautoResolver.ResolveWithRetry(
                request,
                request.Locator,
                options,
                TimeSpan.FromMilliseconds(timeoutMs),
                TimeSpan.FromMilliseconds(pollMs),
                cancellationToken);

            if (result.Success || !ShouldTryLegacyResolverFallback(request))
                return result;

            try
            {
                var legacy = _newResolver.ResolveOne(request.Locator ?? new UiLocator(), request, options.Purpose);
                if (legacy.Element != null)
                {
                    return new ElementResolution.ElementResolveResult
                    {
                        Success = true,
                        Element = legacy.Element,
                        Snapshot = ElementResolution.ElementSnapshot.FromResolutionSnapshot(
                            Resolution.ElementResolver.CreateSnapshot(legacy.Element)),
                        Strategy = "legacy-fallback",
                        FallbackUsed = true,
                        CandidateCount = legacy.Candidates.Count,
                        Criteria = ElementResolution.ElementSearchCriteria.FromLocator(request.Locator),
                        ElapsedMs = result.ElapsedMs
                    };
                }
            }
            catch
            {
                // return primary resolver diagnostics
            }

            return result;
        }

        var inner = _newResolver.ResolveOne(request.Locator ?? new UiLocator(), request, options.Purpose);
        return new ElementResolution.ElementResolveResult
        {
            Success = inner.Element != null,
            Element = inner.Element,
            Snapshot = inner.Element == null
                ? null
                : ElementResolution.ElementSnapshot.FromResolutionSnapshot(
                    Resolution.ElementResolver.CreateSnapshot(inner.Element)),
            Strategy = inner.Strategy,
            CandidateCount = inner.Candidates.Count,
            Criteria = ElementResolution.ElementSearchCriteria.FromLocator(request.Locator)
        };
    }

    private static bool ShouldTryLegacyResolverFallback(UiRequest request) =>
        request.Fast == true;

    private static ElementResolution.ResolveOptions QueryResolveOptions =>
        new()
        {
            AllowOffscreen = true,
            AllowDisabled = true,
            IncludeHidden = true,
            ReturnCandidates = true,
            Purpose = "query"
        };

    private static ElementResolution.ResolveOptions ReadResolveOptions =>
        new()
        {
            AllowOffscreen = true,
            AllowDisabled = true,
            IncludeHidden = true,
            ReturnCandidates = true,
            Purpose = "read"
        };

    private static ElementResolution.ResolveOptions ClickResolveOptions =>
        new()
        {
            AllowOffscreen = false,
            AllowDisabled = false,
            IncludeHidden = false,
            ReturnCandidates = true,
            ThrowIfNotFound = false,
            Purpose = "click"
        };

    private object BuildPywinautoQueryResponse(string operation, ElementResolution.ElementResolveResult result) =>
        new
        {
            operation,
            found = result.Success,
            exists = result.Success,
            strategy = result.Strategy,
            error = result.Error,
            snapshot = result.Snapshot,
            element = result.Snapshot,
            candidateCount = result.CandidateCount,
            candidates = result.Candidates,
            searchRoot = result.SearchRoot,
            criteria = result.Criteria,
            elapsedMs = result.ElapsedMs,
            ambiguous = result.Ambiguous,
            fallbackUsed = result.FallbackUsed
        };

    // =========================================================================
    // Resolved element for state-query pipeline
    // =========================================================================

    /// <summary>
    /// Resolves the target element for state-query operations (isenabled, isvisible,
    /// isfocused, iswindowactive, isclickable, iseditable, exists, wait).
    /// Unlike action-oriented lookup this resolver:
    /// <list type="bullet">
    ///   <item>Does not require a clickable point.</item>
    ///   <item>Allows offscreen elements.</item>
    ///   <item>Delegates to the central engine resolver so that all new locator
    ///         fields (value, text, matchMode, foundIndex, etc.) are supported.</item>
    /// </list>
    /// Throws <see cref="InvalidOperationException"/> when the element cannot be found.
    /// </summary>
    private ResolvedElement ResolveForStateQuery(UiRequest request)
    {
        var result = PywinautoResolve(request, QueryResolveOptions);
        if (!result.Success || result.Element == null)
        {
            throw new InvalidOperationException(
                result.Error ?? $"Element not found for state query. candidateCount={result.CandidateCount}");
        }

        return new ResolvedElement
        {
            Element = result.Element,
            Strategy = result.Strategy,
            RootStrategy = request.SearchRoot ?? "currentWindow"
        };
    }

    /// <summary>
    /// Walks the UIA parent chain looking for the nearest <see cref="ControlType.ComboBox"/>
    /// ancestor. Returns <c>null</c> when no ComboBox is found within
    /// <see cref="MaxMenuParentChainDepth"/> steps.
    /// </summary>
    private static AutomationElement? TryFindComboBoxAncestor(AutomationElement element)
    {
        var current = element.Parent;
        for (int i = 0; i < MaxMenuParentChainDepth && current != null; i++)
        {
            if (current.ControlType == ControlType.ComboBox)
                return current;
            current = current.Parent;
        }
        return null;
    }

    // =========================================================================
    // Wait operation — live state polling
    // =========================================================================

    /// <summary>
    /// Carries the outcome of a single state evaluation: whether the target
    /// condition is met and the raw payload to return on success.
    /// </summary>
    private sealed class StateResult
    {
        public bool Matched { get; init; }
        public object? Payload { get; init; }
    }

    private object Wait(UiRequest request, CancellationToken cancellationToken)
    {
        var state = string.IsNullOrWhiteSpace(request.State)
            ? "exists"
            : request.State.Trim().ToLowerInvariant();

        var timeoutMs      = request.TimeoutMs      ?? 10000;
        var pollIntervalMs = request.PollIntervalMs ?? 200;

        var start           = Stopwatch.StartNew();
        Exception? lastEx   = null;
        object? lastPayload = null;

        while (start.ElapsedMilliseconds < timeoutMs)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var stateResult = EvaluateState(request, state);
                lastPayload = stateResult.Payload;

                if (stateResult.Matched)
                {
                    return new
                    {
                        operation = "wait",
                        success   = true,
                        state,
                        elapsedMs = start.ElapsedMilliseconds,
                        result    = stateResult.Payload
                    };
                }
            }
            catch (Exception ex)
            {
                lastEx = ex;
            }

            WaitWithCancellation(pollIntervalMs, cancellationToken);
        }

        return new
        {
            operation  = "wait",
            success    = false,
            state,
            elapsedMs  = start.ElapsedMilliseconds,
            lastError  = lastEx?.Message,
            lastResult = lastPayload
        };
    }

    private StateResult EvaluateState(UiRequest request, string state)
    {
        return state switch
        {
            "exists"       => EvaluateExists(request),
            "enabled"      => EvaluateEnabled(request, expected: true),
            "disabled"     => EvaluateEnabled(request, expected: false),
            "visible"      => EvaluateVisible(request, expected: true),
            "hidden"       => EvaluateVisible(request, expected: false),
            "focused"      => EvaluateFocused(request, expected: true),
            "windowactive" => EvaluateWindowActive(request, expected: true),
            "clickable"    => EvaluateClickable(request),
            "editable"     => EvaluateEditable(request),
            "gone"         => EvaluateGone(request),
            "selected"     => EvaluateSelected(request, expected: true),
            "checked"      => EvaluateChecked(request, expected: true),
            "unchecked"    => EvaluateChecked(request, expected: false),
            _ => throw new ArgumentException(
                $"Unsupported wait state '{state}'. " +
                "Supported: exists, enabled, disabled, visible, hidden, focused, " +
                "windowactive, clickable, editable, gone, selected, checked, unchecked.")
        };
    }

    private StateResult EvaluateSelected(UiRequest request, bool expected)
    {
        var resolved = ResolveForStateQuery(request);
        bool isSelected = false;
        try
        {
            if (resolved.Element.Patterns.SelectionItem.IsSupported)
            {
                isSelected = resolved.Element.Patterns.SelectionItem.Pattern.IsSelected;
            }
        }
        catch { }
        return new StateResult
        {
            Matched = isSelected == expected,
            Payload = new { selected = isSelected, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateChecked(UiRequest request, bool expected)
    {
        var resolved = ResolveForStateQuery(request);
        bool isChecked = false;
        try
        {
            if (resolved.Element.Patterns.Toggle.IsSupported)
            {
                var toggleState = resolved.Element.Patterns.Toggle.Pattern.ToggleState;
                isChecked = toggleState == FlaUI.Core.Definitions.ToggleState.On;
            }
        }
        catch { }
        return new StateResult
        {
            Matched = isChecked == expected,
            Payload = new { @checked = isChecked, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateExists(UiRequest request)
    {
        try
        {
            var resolved = ResolveForStateQuery(request);
            return new StateResult
            {
                Matched = true,
                Payload = new
                {
                    exists       = true,
                    strategy     = resolved.Strategy,
                    automationId = SafeElementAutomationId(resolved.Element),
                    name         = SafeElementName(resolved.Element),
                    controlType  = SafeElementControlType(resolved.Element)
                }
            };
        }
        catch
        {
            return new StateResult { Matched = false };
        }
    }

    private StateResult EvaluateGone(UiRequest request)
    {
        try
        {
            ResolveForStateQuery(request);
            // Element still exists.
            return new StateResult { Matched = false };
        }
        catch
        {
            // Element not found — it is gone.
            return new StateResult { Matched = true, Payload = new { gone = true } };
        }
    }

    private StateResult EvaluateEnabled(UiRequest request, bool expected)
    {
        var resolved   = ResolveForStateQuery(request);
        var uiaEnabled = resolved.Element.IsEnabled;
        return new StateResult
        {
            Matched = uiaEnabled == expected,
            Payload = new { enabled = uiaEnabled, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateVisible(UiRequest request, bool expected)
    {
        var resolved    = ResolveForStateQuery(request);
        var rect        = resolved.Element.BoundingRectangle;
        var isOffscreen = resolved.Element.Properties.IsOffscreen.ValueOrDefault;
        var hasRect     = !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
        var visible     = hasRect && !isOffscreen;
        return new StateResult
        {
            Matched = visible == expected,
            Payload = new { visible, isOffscreen, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateFocused(UiRequest request, bool expected)
    {
        var resolved = ResolveForStateQuery(request);
        var focused  = resolved.Element.Properties.HasKeyboardFocus.ValueOrDefault;
        return new StateResult
        {
            Matched = focused == expected,
            Payload = new { focused, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateWindowActive(UiRequest request, bool expected)
    {
        var resolved   = ResolveForStateQuery(request);
        var hwnd       = SafeWindowHandle(resolved.Element);
        var rootHwnd   = hwnd != IntPtr.Zero ? GetAncestor(hwnd, GA_ROOT) : IntPtr.Zero;
        var foreground = GetForegroundWindow();
        var active     = rootHwnd != IntPtr.Zero && rootHwnd == foreground;
        return new StateResult
        {
            Matched = active == expected,
            Payload = new
            {
                active,
                hwnd           = rootHwnd.ToInt64(),
                foregroundHwnd = foreground.ToInt64(),
                strategy       = resolved.Strategy
            }
        };
    }

    private StateResult EvaluateClickable(UiRequest request)
    {
        var resolved = ResolveForStateQuery(request);
        var element  = resolved.Element;
        var enabled  = element.IsEnabled;
        var visible  = IsTargetPracticallyVisible(element, null, out _);
        var hasClickablePoint = TryGetElementClickablePoint(element, out _, out _);
        var clickable = enabled && visible && hasClickablePoint;
        return new StateResult
        {
            Matched = clickable,
            Payload = new { clickable, enabled, visible, hasClickablePoint, strategy = resolved.Strategy }
        };
    }

    private StateResult EvaluateEditable(UiRequest request)
    {
        var resolved     = ResolveForStateQuery(request);
        var element      = resolved.Element;
        var enabled      = element.IsEnabled;
        var valuePattern = element.Patterns.Value.PatternOrDefault;
        var hasValuePattern = valuePattern != null;
        var isReadOnly   = valuePattern?.IsReadOnly ?? true;
        var editable     = enabled && hasValuePattern && !isReadOnly;
        return new StateResult
        {
            Matched = editable,
            Payload = new { editable, enabled, hasValuePattern, isReadOnly, strategy = resolved.Strategy }
        };
    }

    /// <summary>
    /// Sleeps for at most <paramref name="ms"/> milliseconds, honouring
    /// <paramref name="cancellationToken"/> on each iteration.
    /// </summary>
    private static void WaitWithCancellation(int ms, CancellationToken cancellationToken)
    {
        const int SliceMs = 50;
        var elapsed = 0;
        while (elapsed < ms)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var slice = Math.Min(SliceMs, ms - elapsed);
            Thread.Sleep(slice);
            elapsed += slice;
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
    internal static AutomationElement? FindByXPath(
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

    internal static ControlType ParseControlType(string value)
    {
        if (Enum.TryParse<ControlType>(value, ignoreCase: true, out var ct))
            return ct;
        throw new ArgumentException(
            $"Unknown controlType '{value}'. " +
            "Valid values: Button, CheckBox, ComboBox, Custom, DataGrid, DataItem, Edit, " +
            "Group, Header, HeaderItem, List, ListItem, Menu, MenuItem, Pane, RadioButton, " +
            "Slider, Spinner, Tab, TabItem, Table, Text, ToolBar, Tree, TreeItem, Window.");
    }

    internal static string DescribeLocator(UiLocator l)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(l.XPath))         parts.Add($"xpath='{l.XPath}'");
        if (!string.IsNullOrWhiteSpace(l.Name))           parts.Add($"name='{l.Name}'");
        if (!string.IsNullOrWhiteSpace(l.AutomationId))   parts.Add($"automationId='{l.AutomationId}'");
        if (!string.IsNullOrWhiteSpace(l.ClassName))      parts.Add($"className='{l.ClassName}'");
        if (!string.IsNullOrWhiteSpace(l.ControlType))    parts.Add($"controlType='{l.ControlType}'");
        if (l.Hwnd.HasValue)                               parts.Add($"hwnd={l.Hwnd.Value}");
        if (l.ProcessId.HasValue)                          parts.Add($"processId={l.ProcessId.Value}");
        if (!string.IsNullOrWhiteSpace(l.Value))           parts.Add($"value='{l.Value}'");
        if (!string.IsNullOrWhiteSpace(l.Text))            parts.Add($"text='{l.Text}'");
        if (!string.IsNullOrWhiteSpace(l.MatchMode))       parts.Add($"matchMode='{l.MatchMode}'");
        if (l.FoundIndex.HasValue)                         parts.Add($"foundIndex={l.FoundIndex.Value}");
        if (l.CtrlIndex.HasValue)                          parts.Add($"ctrlIndex={l.CtrlIndex.Value}");
        if (l.Depth.HasValue)                              parts.Add($"depth={l.Depth.Value}");
        if (l.TopLevelOnly == true)                        parts.Add("topLevelOnly=true");
        return string.Join(", ", parts);
    }

    private static object? DescribeLocatorAsObject(UiLocator? locator)
    {
        if (locator == null)
            return null;

        return new
        {
            automationId = locator.AutomationId,
            name         = locator.Name,
            controlType  = locator.ControlType,
            className    = locator.ClassName,
            xpath        = locator.XPath,
            hwnd         = locator.Hwnd,
            processId    = locator.ProcessId,
            value        = locator.Value,
            text         = locator.Text,
            matchMode    = locator.MatchMode,
            foundIndex   = locator.FoundIndex,
            ctrlIndex    = locator.CtrlIndex,
            depth        = locator.Depth,
            topLevelOnly = locator.TopLevelOnly,
            frameworkId  = locator.FrameworkId
        };
    }

    private (ResolvedElement first, ResolvedElement second) ResolveTwoElementsForPosition(
        UiRequest request)
    {
        if (request.Locator == null)
            throw new ArgumentException("'locator' is required for position operations.");
        if (request.Locator2 == null)
            throw new ArgumentException("'locator2' is required for position operations.");

        var firstRequest = new UiRequest
        {
            Operation                                 = request.Operation,
            Locator                                   = request.Locator,
            ParentLocator                             = request.ParentLocator,
            TimeoutMs                                 = request.TimeoutMs,
            Fast                                      = request.Fast,
            DisableAutoFollow                         = request.DisableAutoFollow,
            UseCache                                  = request.UseCache,
            PreferXPath                               = request.PreferXPath,
            XPathOnly                                 = request.XPathOnly,
            PreferAttributes                          = request.PreferAttributes,
            FallbackToWindowRootIfParentChildNotFound = request.FallbackToWindowRootIfParentChildNotFound
        };

        var secondRequest = new UiRequest
        {
            Operation                                 = request.Operation,
            Locator                                   = request.Locator2,
            ParentLocator                             = request.ParentLocator,
            TimeoutMs                                 = request.TimeoutMs,
            Fast                                      = request.Fast,
            DisableAutoFollow                         = request.DisableAutoFollow,
            UseCache                                  = request.UseCache,
            PreferXPath                               = request.PreferXPath,
            XPathOnly                                 = request.XPathOnly,
            PreferAttributes                          = request.PreferAttributes,
            FallbackToWindowRootIfParentChildNotFound = request.FallbackToWindowRootIfParentChildNotFound
        };

        var first  = ResolveForStateQuery(firstRequest);
        var second = ResolveForStateQuery(secondRequest);

        return (first, second);
    }

    private (Rectangle r1, Rectangle r2, ResolvedElement e1, ResolvedElement e2) GetTwoRectsLive(
        UiRequest request)
    {
        var (e1, e2) = ResolveTwoElementsForPosition(request);

        var r1 = e1.Element.BoundingRectangle;
        var r2 = e2.Element.BoundingRectangle;

        if (r1.IsEmpty || r1.Width <= 0 || r1.Height <= 0)
            throw new InvalidOperationException(
                $"First element has empty bounding rectangle. locator={DescribeLocator(request.Locator!)}");

        if (r2.IsEmpty || r2.Width <= 0 || r2.Height <= 0)
            throw new InvalidOperationException(
                $"Second element has empty bounding rectangle. locator2={DescribeLocator(request.Locator2!)}");

        return (r1, r2, e1, e2);
    }

    private static double CenterX(Rectangle r) => r.Left + (r.Width / 2.0);
    private static double CenterY(Rectangle r) => r.Top  + (r.Height / 2.0);

    private static bool RectanglesOverlapVertically(Rectangle a, Rectangle b)
        => a.Top < b.Bottom && a.Bottom > b.Top;

    private static bool RectanglesOverlapHorizontally(Rectangle a, Rectangle b)
        => a.Left < b.Right && a.Right > b.Left;

    private static object RectObj(Rectangle r) => new
    {
        left    = r.Left,
        top     = r.Top,
        right   = r.Right,
        bottom  = r.Bottom,
        width   = r.Width,
        height  = r.Height,
        centerX = CenterX(r),
        centerY = CenterY(r),
        isEmpty = r.IsEmpty
    };

    private static object ElementInfoObj(AutomationElement element) => new
    {
        automationId = SafeElementAutomationId(element),
        name         = SafeElementName(element),
        controlType  = SafeElementControlType(element),
        className    = SafeElementClassName(element),
        isEnabled    = SafeIsEnabled(element),
        isOffscreen  = SafeIsOffscreen(element)
    };

    private void SetToggle(UiRequest req, bool wantChecked)
    {
        var element = ResolveElementForOperation(
            req,
            purpose: wantChecked ? "check" : "uncheck",
            action: true,
            allowOffscreen: false,
            requireClickable: true);

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

    /// <summary>
    /// Tries to obtain the clickable point for <paramref name="element"/>.
    /// Returns <c>true</c> and populates <paramref name="x"/>/<paramref name="y"/>
    /// when the element exposes a clickable point; returns <c>false</c> otherwise.
    /// </summary>
    private static bool TryGetElementClickablePoint(
        AutomationElement element, out double x, out double y)
    {
        try
        {
            var pt = element.GetClickablePoint();
            x = pt.X;
            y = pt.Y;
            return true;
        }
        catch
        {
            x = 0;
            y = 0;
            return false;
        }
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

    private static void SendDatePickerKey(VirtualKeyShort key)
    {
        Keyboard.Press(key);
        Thread.Sleep(25);
        Keyboard.Release(key);
    }

    private object? SelectComboBoxNativeUia(UiRequest request, CancellationToken cancellationToken)
    {
        var session = TryGetSessionOrNull();

        IntPtr? rootHwnd = null;
        int? processId = null;

        try
        {
            processId = session?.Application?.ProcessId;
        }
        catch
        {
            processId = null;
        }

        try
        {
            if (session?.ActiveWindow != null)
            {
                var hwnd = SafeWindowHandle(session.ActiveWindow);
                if (hwnd != IntPtr.Zero)
                    rootHwnd = hwnd;
            }
        }
        catch
        {
            rootHwnd = null;
        }

        var selector = new DesktopAutomationDriver.NativeUia.NativeUiaComboBoxSelector(_logger);
        return selector.Select(request, rootHwnd, processId, cancellationToken);
    }

    private object? FindComboBoxNativeUia(UiRequest request, CancellationToken cancellationToken)
    {
        var session = TryGetSessionOrNull();

        IntPtr? rootHwnd = null;
        int? processId = null;

        try
        {
            processId = session?.Application?.ProcessId;
        }
        catch
        {
            processId = null;
        }

        try
        {
            if (session?.ActiveWindow != null)
            {
                var hwnd = SafeWindowHandle(session.ActiveWindow);
                if (hwnd != IntPtr.Zero)
                    rootHwnd = hwnd;
            }
        }
        catch
        {
            rootHwnd = null;
        }

        var selector = new DesktopAutomationDriver.NativeUia.NativeUiaComboBoxSelector(_logger);
        return selector.FindOnly(request, rootHwnd, processId, cancellationToken);
    }

    private object? InspectComboBox(UiRequest req, CancellationToken cancellationToken = default)
    {
        var session = RequireSession();
        return _nativeUiaComboBoxService.InspectComboBox(
            req,
            GetActiveWindowHwndOrNull(session),
            GetActiveProcessIdOrNull(session),
            cancellationToken);
    }

    private static bool IsComboBoxSelectRequest(UiRequest req)
    {
        var controlType = req.Locator?.ControlType;
        if (string.IsNullOrWhiteSpace(controlType))
            return false;

        return string.Equals(controlType, "ComboBox", StringComparison.OrdinalIgnoreCase)
               || string.Equals(controlType, "Edit", StringComparison.OrdinalIgnoreCase);
    }

    private static IntPtr? GetActiveWindowHwndOrNull(AutomationSession session)
    {
        try
        {
            var activeWindow = session.ActiveWindow ?? session.Application?.GetMainWindow(session.Automation);
            if (activeWindow == null)
                return null;

            return activeWindow.Properties.NativeWindowHandle.Value;
        }
        catch
        {
            return null;
        }
    }

    private static int? GetActiveProcessIdOrNull(AutomationSession session)
    {
        try
        {
            return session.Application?.ProcessId;
        }
        catch
        {
            return null;
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

    // =========================================================================
    // Win32 native ComboBox strategies (pywinauto-style)
    // =========================================================================

    /// <summary>
    /// Attempts to select <paramref name="requestedValue"/> in a native Win32 ComboBox
    /// using CB_FINDSTRINGEXACT / CB_SETCURSEL (no dropdown open).
    /// Returns true and sets <paramref name="strategy"/> when the selection is confirmed.
    /// </summary>
    private bool TrySelectNativeWin32ComboBox(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        out string strategy)
    {
        strategy = "win32-native-not-used";

        try
        {
            var hwnd      = SafeWindowHandle(comboBox);
            var className = SafeElementClassName(comboBox);

            if (!Win32ComboBoxHelper.IsLikelyNativeWin32ComboBox(hwnd, className))
                return false;

            _logger.LogInformation(
                "Trying native Win32 ComboBox selection. hwnd=0x{Hwnd:X}, className={ClassName}, value={Value}",
                hwnd.ToInt64(), className, requestedValue);

            if (!Win32ComboBoxHelper.SelectByText(hwnd, requestedValue, _logger, "ui-win32-native-combobox"))
                return false;

            Win32ComboBoxHelper.HideDropdown(hwnd);
            Thread.Sleep(ComboBoxSelectionCommitDelayMs);

            if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, "ui-win32-native-combobox"))
            {
                _logger.LogInformation(
                    "ComboBox selection strategy selected: win32-native-combobox. combo={Combo}, value={Value}",
                    SafeElementName(comboBox), requestedValue);

                strategy = "win32-native-combobox";
                return true;
            }

            _logger.LogWarning(
                "Native Win32 ComboBox selection did not verify. requested={Requested}, actual={Actual}, hwnd=0x{Hwnd:X}",
                requestedValue, GetComboBoxCurrentValue(session, comboBox), hwnd.ToInt64());

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Native Win32 ComboBox selection failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox), requestedValue);

            return false;
        }
    }

    /// <summary>
    /// Attempts to select the item at <paramref name="index"/> in a native Win32 ComboBox
    /// using CB_SETCURSEL.  Verifies that the selected text matches the expected item text
    /// before returning success.
    /// Returns true and sets <paramref name="strategy"/> when the selection is verified.
    /// </summary>
    private bool TrySelectNativeWin32ComboBoxByIndex(
        AutomationSession session,
        AutomationElement comboBox,
        int index,
        out string strategy)
    {
        strategy = "win32-native-index-not-used";

        try
        {
            var hwnd      = SafeWindowHandle(comboBox);
            var className = SafeElementClassName(comboBox);

            if (!Win32ComboBoxHelper.IsLikelyNativeWin32ComboBox(hwnd, className))
                return false;

            // Read expected item text before changing the selection so we can verify afterward.
            var expectedText = Win32ComboBoxHelper.GetItemText(hwnd, index);
            if (string.IsNullOrWhiteSpace(expectedText))
            {
                _logger.LogWarning(
                    "Win32 ComboBox index selection aborted: expected item text is empty. hwnd=0x{Hwnd:X}, index={Index}",
                    hwnd.ToInt64(), index);
                return false;
            }

            _logger.LogInformation(
                "Trying native Win32 ComboBox index selection. hwnd=0x{Hwnd:X}, className={ClassName}, index={Index}, expected={Expected}",
                hwnd.ToInt64(), className, index, expectedText);

            if (!Win32ComboBoxHelper.SelectByIndex(hwnd, index, _logger, "ui-win32-native-combobox-index"))
                return false;

            Win32ComboBoxHelper.HideDropdown(hwnd);
            Thread.Sleep(ComboBoxSelectionCommitDelayMs);

            var selectedText = Win32ComboBoxHelper.GetSelectedText(hwnd);

            if (!string.Equals(
                    NormalizeMenuText(selectedText),
                    NormalizeMenuText(expectedText),
                    StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning(
                    "Win32 ComboBox index selection did not verify. hwnd=0x{Hwnd:X}, index={Index}, expected={Expected}, actual={Actual}",
                    hwnd.ToInt64(), index, expectedText, selectedText);
                return false;
            }

            _logger.LogInformation(
                "ComboBox selection verified: win32-native-combobox-index. combo={Combo}, index={Index}, actual={Actual}",
                SafeElementName(comboBox), index, selectedText);

            strategy = "win32-native-combobox-index";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Native Win32 ComboBox index selection failed. combo={Combo}, index={Index}",
                SafeElementName(comboBox), index);

            return false;
        }
    }

    private object SelectComboBoxByIndex(
        AutomationSession session,
        AutomationElement comboBox,
        int index,
        UiRequest req)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), "ComboBox index must be >= 0.");

        _logger.LogInformation(
            "Trying ComboBox select by index. index={Index}, comboBox={ComboBox}",
            index,
            SafeElementName(comboBox));

        // Strategy 1: Win32 native ComboBox index selection.
        if (TrySelectNativeWin32ComboBoxByIndex(session, comboBox, index, out var win32Strategy))
        {
            var hwnd   = SafeWindowHandle(comboBox);
            var actual = Win32ComboBoxHelper.GetSelectedText(hwnd);

            _logger.LogInformation(
                "ComboBox selected by index using Win32 native strategy. index={Index}, actual={Actual}, strategy={Strategy}",
                index,
                actual,
                win32Strategy);

            return new
            {
                operation = "select",
                selectedIndex = index,
                actual,
                comboBox = SafeElementName(comboBox),
                verified = true,
                strategy = win32Strategy
            };
        }

        // Strategy 2: UIA index selection.
        if (TrySelectComboBoxByUiaIndex(session, comboBox, index, out var uiaStrategy))
        {
            var actual = GetComboBoxCurrentValue(session, comboBox);

            _logger.LogInformation(
                "ComboBox selected by index using UIA strategy. index={Index}, actual={Actual}, strategy={Strategy}",
                index,
                actual,
                uiaStrategy);

            return new
            {
                operation = "select",
                selectedIndex = index,
                actual,
                comboBox = SafeElementName(comboBox),
                verified = true,
                strategy = uiaStrategy
            };
        }

        throw new InvalidOperationException(
            $"Unable to select ComboBox index {index}. ComboBox={SafeElementName(comboBox)}");
    }

    private bool TrySelectComboBoxByUiaIndex(
        AutomationSession session,
        AutomationElement comboBox,
        int index,
        out string strategy)
    {
        strategy = "uia-index-not-used";

        try
        {
            // Expand the dropdown so items are accessible.
            OpenComboBoxDropdown(session, comboBox);
            Thread.Sleep(MenuExpandDelayMs);

            // Collect all ListItem descendants.
            var cf    = session.Automation.ConditionFactory;
            var items = comboBox.FindAllDescendants(cf.ByControlType(ControlType.ListItem));

            if (items.Length == 0)
            {
                // Some combo boxes nest items inside a List child element.
                var listChild = comboBox.FindFirstDescendant(cf.ByControlType(ControlType.List));
                if (listChild != null)
                    items = listChild.FindAllChildren();
            }

            if (items.Length == 0)
            {
                // Last attempt: look for the popup list near the combo box.
                var dynamicList = FindDynamicComboBoxList(session, comboBox);
                if (dynamicList != null)
                    items = dynamicList.FindAllChildren(cf.ByControlType(ControlType.ListItem));
            }

            if (index >= items.Length)
            {
                _logger.LogInformation(
                    "UIA ComboBox index search failed. index={Index}, availableCount={Count}, comboBox={ComboBox}",
                    index,
                    items.Length,
                    SafeElementName(comboBox));
                return false;
            }

            var item     = items[index];
            var itemText = SafeElementName(item);

            // Ensure we have a non-empty item text to verify against after selection.
            if (string.IsNullOrWhiteSpace(itemText))
            {
                try { itemText = item.Properties.Name.Value ?? string.Empty; }
                catch (Exception nameEx)
                {
                    _logger.LogDebug(nameEx, "Could not read Properties.Name from ComboBox ListItem at index {Index}.", index);
                    itemText = string.Empty;
                }
            }

            // Try SelectionItem pattern first.
            item.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();
            Thread.Sleep(ComboBoxSelectionCommitDelayMs);

            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(ComboBoxSelectionCommitDelayMs);
                strategy = "uia-index-selection-item";
            }
            else if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(ComboBoxSelectionCommitDelayMs);
                strategy = "uia-index-invoke";
            }
            else
            {
                item.Click();
                Thread.Sleep(ComboBoxSelectionCommitDelayMs);
                strategy = "uia-index-click";
            }

            // Collapse the dropdown.
            try { comboBox.Patterns.ExpandCollapse.PatternOrDefault?.Collapse(); }
            catch (Exception collapseEx)
            {
                _logger.LogDebug(collapseEx, "ComboBox collapse failed after index selection. comboBox={ComboBox}", SafeElementName(comboBox));
            }

            WaitForComboBoxDropdownToClose(comboBox, timeoutMs: 1000);

            _logger.LogInformation(
                "UIA ComboBox index selection committed. index={Index}, item={Item}, strategy={Strategy}",
                index,
                itemText,
                strategy);

            // Verify that the final ComboBox value matches the selected item.
            // An empty item text means we have nothing to compare against, so we
            // cannot prove the selection succeeded — fail safe.
            if (string.IsNullOrWhiteSpace(itemText))
            {
                _logger.LogWarning(
                    "UIA ComboBox index selection cannot verify because item text is empty. index={Index}, strategy={Strategy}, comboBox={ComboBox}",
                    index,
                    strategy,
                    SafeElementName(comboBox));

                return false;
            }

            if (!VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, itemText, strategy))
            {
                _logger.LogWarning(
                    "UIA ComboBox index selection verification failed. strategy={Strategy}, index={Index}, expected={Expected}, comboBox={ComboBox}",
                    strategy,
                    index,
                    itemText,
                    SafeElementName(comboBox));

                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "UIA ComboBox index selection failed. comboBox={ComboBox}, index={Index}",
                SafeElementName(comboBox),
                index);

            return false;
        }
    }

    private bool TrySelectComboBoxByDirectUia(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline,
        out string strategy)
    {
        strategy = "direct-uia-not-used";

        try
        {
            if (string.IsNullOrWhiteSpace(requestedValue))
                return false;

            var logicalMatch = FindExactComboBoxListItemLogical(
                session,
                comboBox,
                requestedValue,
                maxItems: ComboBoxDirectUiaSearchLimit);

            if (logicalMatch != null)
            {
                if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                    !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "direct UIA logical commit"))
                {
                    return false;
                }

                var capabilities = DetectComboBoxItemPatternCapabilities(logicalMatch);

                if (capabilities.HasAnyUsefulPattern &&
                    TryCommitComboBoxItemUsingAvailablePatterns(
                        session,
                        comboBox,
                        logicalMatch,
                        requestedValue,
                        capabilities,
                        guard,
                        operationDeadline,
                        "direct-uia-logical-or-popup"))
                {
                    strategy = "direct-uia-pattern-based";
                    return true;
                }

                _logger.LogInformation(
                    "Direct UIA logical ListItem pattern commit did not verify. requested={Requested}, actual={Actual}, combo={Combo}, capabilities={Capabilities}",
                    requestedValue,
                    GetComboBoxCurrentValue(session, comboBox),
                    SafeElementName(comboBox),
                    capabilities.ToString());
            }

            if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) &&
                IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "direct UIA reopen") &&
                OpenComboBoxDropdown(session, comboBox))
            {
                Thread.Sleep(MenuExpandDelayMs);

                var popupMatch = FindExactComboBoxListItemFromPopup(
                    session,
                    comboBox,
                    requestedValue,
                    maxItems: ComboBoxDirectUiaSearchLimit);

                if (popupMatch != null)
                {
                    var capabilities = DetectComboBoxItemPatternCapabilities(popupMatch);

                    if (capabilities.HasAnyUsefulPattern &&
                        TryCommitComboBoxItemUsingAvailablePatterns(
                            session,
                            comboBox,
                            popupMatch,
                            requestedValue,
                            capabilities,
                            guard,
                            operationDeadline,
                            "direct-uia-logical-or-popup"))
                    {
                        strategy = "direct-uia-pattern-based";
                        return true;
                    }

                    _logger.LogInformation(
                        "Direct UIA popup ListItem pattern commit did not verify. requested={Requested}, actual={Actual}, combo={Combo}, capabilities={Capabilities}",
                        requestedValue,
                        GetComboBoxCurrentValue(session, comboBox),
                        SafeElementName(comboBox),
                        capabilities.ToString());
                }
            }

            if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) &&
                IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "direct UIA value-pattern commit") &&
                TrySetComboBoxValueByValuePattern(session, comboBox, requestedValue, "direct-uia-valuepattern"))
            {
                strategy = "direct-uia-valuepattern";
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Direct UIA ComboBox selection failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                requestedValue);

            return false;
        }
    }

    private AutomationElement? FindExactComboBoxListItemLogical(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        int maxItems)
    {
        try
        {
            var items = GetLogicalComboBoxItems(session, comboBox, maxItems);

            foreach (var item in items)
            {
                if (ComboBoxItemMatchesExactText(item, requestedValue))
                {
                    _logger.LogInformation(
                        "Direct UIA logical ComboBox item found. combo={Combo}, requested={Requested}, item={Item}",
                        SafeElementName(comboBox),
                        requestedValue,
                        SafeElementName(item));

                    return item;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Direct UIA logical ListItem search failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                requestedValue);
        }

        return null;
    }

    private AutomationElement? FindExactComboBoxListItemFromPopup(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        int maxItems)
    {
        try
        {
            var list = FindDynamicComboBoxList(session, comboBox);

            if (list == null)
                return null;

            var items = GetListItemsBounded(session, list, maxItems);

            foreach (var item in items)
            {
                if (ComboBoxItemMatchesExactText(item, requestedValue))
                {
                    _logger.LogInformation(
                        "Direct UIA popup ComboBox item found. combo={Combo}, requested={Requested}, item={Item}",
                        SafeElementName(comboBox),
                        requestedValue,
                        SafeElementName(item));

                    return item;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Direct UIA popup ListItem search failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                requestedValue);
        }

        return null;
    }

    private bool ComboBoxItemMatchesExactText(
        AutomationElement item,
        string requestedValue)
    {
        try
        {
            var requested = NormalizeMenuText(requestedValue);

            var name = NormalizeMenuText(SafeElementName(item));
            if (string.Equals(name, requested, StringComparison.OrdinalIgnoreCase))
                return true;

            var automationId = NormalizeMenuText(SafeElementAutomationId(item));
            if (string.Equals(automationId, requested, StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }
        catch
        {
            return false;
        }
    }

    private sealed class ComboBoxItemPatternCapabilities
    {
        public bool HasScrollItem { get; init; }
        public bool HasSelectionItem { get; init; }
        public bool HasInvoke { get; init; }

        public bool HasAnyUsefulPattern =>
            HasScrollItem || HasSelectionItem || HasInvoke;

        public override string ToString()
        {
            return $"ScrollItem={HasScrollItem}, SelectionItem={HasSelectionItem}, Invoke={HasInvoke}";
        }
    }

    /// <summary>
    /// Captures ComboBox identity so multi-step selection can abort if UI focus or UIA refresh points at another ComboBox.
    /// </summary>
    private sealed class ComboBoxTargetGuard
    {
        public string AutomationId { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string ControlType { get; init; } = string.Empty;
        public string RuntimeId { get; init; } = string.Empty;
        public Rectangle BoundingRectangle { get; init; }
    }

    private ComboBoxTargetGuard CaptureComboBoxTargetGuard(AutomationElement comboBox)
    {
        return new ComboBoxTargetGuard
        {
            AutomationId = SafeElementAutomationId(comboBox) ?? string.Empty,
            Name = SafeElementName(comboBox) ?? string.Empty,
            ControlType = comboBox.ControlType.ToString(),
            RuntimeId = SafeRuntimeId(comboBox),
            BoundingRectangle = comboBox.BoundingRectangle
        };
    }

    private string SafeRuntimeId(AutomationElement element)
    {
        try
        {
            var runtimeId = element.Properties.RuntimeId.ValueOrDefault;
            return runtimeId == null ? string.Empty : string.Join(".", runtimeId);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to read ComboBox RuntimeId.");
            return string.Empty;
        }
    }

    private bool IsSameComboBoxTarget(
        AutomationElement comboBox,
        ComboBoxTargetGuard guard)
    {
        try
        {
            var runtimeId = SafeRuntimeId(comboBox);
            if (!string.IsNullOrWhiteSpace(guard.RuntimeId) &&
                string.Equals(runtimeId, guard.RuntimeId, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var automationId = SafeElementAutomationId(comboBox);
            var name = SafeElementName(comboBox);
            var rect = comboBox.BoundingRectangle;

            var sameAutomationId =
                string.Equals(automationId, guard.AutomationId, StringComparison.OrdinalIgnoreCase);

            var sameName =
                string.Equals(name, guard.Name, StringComparison.OrdinalIgnoreCase);

            var closeRect =
                Math.Abs(rect.Left - guard.BoundingRectangle.Left) <= ComboBoxRefetchRectangleTolerancePx &&
                Math.Abs(rect.Top - guard.BoundingRectangle.Top) <= ComboBoxRefetchRectangleTolerancePx;

            return sameAutomationId && sameName && closeRect;
        }
        catch
        {
            return false;
        }
    }

    private bool IsComboBoxTargetGuardValid(
        AutomationElement comboBox,
        ComboBoxTargetGuard guard,
        string requestedValue,
        string action)
    {
        if (IsSameComboBoxTarget(comboBox, guard))
            return true;

        _logger.LogWarning(
            "ComboBox target changed before {Action}. Aborting to avoid opening another dropdown. requested={Requested}, originalName={OriginalName}, currentName={CurrentName}",
            action,
            requestedValue,
            guard.Name,
            SafeElementName(comboBox));

        return false;
    }

    private bool IsComboBoxOperationWithinDeadline(
        DateTime operationDeadline,
        AutomationElement comboBox,
        string requestedValue)
    {
        if (DateTime.UtcNow <= operationDeadline)
            return true;

        _logger.LogWarning(
            "ComboBox selection operation timed out before HTTP timeout. requested={Requested}, combo={Combo}",
            requestedValue,
            SafeElementName(comboBox));

        return false;
    }

    private static bool IsComboBoxTabBlurCommitFallbackAllowed()
    {
        return ComboBoxAllowTabBlurCommitFallback;
    }

    private ComboBoxItemPatternCapabilities DetectComboBoxItemPatternCapabilities(
        AutomationElement item)
    {
        try
        {
            var capabilities = new ComboBoxItemPatternCapabilities
            {
                HasScrollItem = item.Patterns.ScrollItem.IsSupported,
                HasSelectionItem = item.Patterns.SelectionItem.IsSupported,
                HasInvoke = item.Patterns.Invoke.IsSupported
            };

            _logger.LogInformation(
                "ComboBox ListItem pattern capabilities detected. item={Item}, capabilities={Capabilities}",
                SafeElementName(item),
                capabilities.ToString());

            return capabilities;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Failed to detect ComboBox ListItem pattern capabilities. item={Item}",
                SafeElementName(item));

            return new ComboBoxItemPatternCapabilities();
        }
    }

    private ComboBoxItemPatternCapabilities ProbeComboBoxDropdownItemCapabilities(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline)
    {
        try
        {
            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "capability probe reopen"))
            {
                return new ComboBoxItemPatternCapabilities();
            }

            if (!OpenComboBoxDropdown(session, comboBox))
                return new ComboBoxItemPatternCapabilities();

            Thread.Sleep(MenuExpandDelayMs);

            var list = FindDynamicComboBoxList(session, comboBox);
            if (list == null)
                return new ComboBoxItemPatternCapabilities();

            var items = GetListItemsBounded(session, list, maxItems: 5);

            var firstItem = items.FirstOrDefault(x =>
                !string.IsNullOrWhiteSpace(SafeElementName(x)) ||
                !string.IsNullOrWhiteSpace(SafeElementAutomationId(x)));

            if (firstItem == null)
                return new ComboBoxItemPatternCapabilities();

            var capabilities = DetectComboBoxItemPatternCapabilities(firstItem);
            _logger.LogInformation(
                "ComboBox pattern capability probe: Combo={Combo}, FirstItem={FirstItem}, ScrollItem={ScrollItem}, SelectionItem={SelectionItem}, Invoke={Invoke}, Decision={Decision}",
                SafeElementName(comboBox),
                SafeElementName(firstItem),
                capabilities.HasScrollItem,
                capabilities.HasSelectionItem,
                capabilities.HasInvoke,
                capabilities.HasAnyUsefulPattern
                    ? "Use pattern-based selection"
                    : "Fallback to visual search because no useful ListItem patterns were detected");

            return capabilities;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox dropdown pattern probing failed. combo={Combo}",
                SafeElementName(comboBox));

            return new ComboBoxItemPatternCapabilities();
        }
    }

    private bool TryCommitComboBoxItemUsingAvailablePatterns(
        AutomationSession session,
        AutomationElement comboBox,
        AutomationElement item,
        string requestedValue,
        ComboBoxItemPatternCapabilities capabilities,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline,
        string source)
    {
        try
        {
            _logger.LogInformation(
                "Trying ComboBox pattern-based commit. source={Source}, requested={Requested}, item={Item}, capabilities={Capabilities}",
                source,
                requestedValue,
                SafeElementName(item),
                capabilities.ToString());

            if (capabilities.HasScrollItem && item.Patterns.ScrollItem.IsSupported)
            {
                try
                {
                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "ScrollIntoView"))
                    {
                        return false;
                    }

                    item.Patterns.ScrollItem.Pattern.ScrollIntoView();
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    _logger.LogInformation(
                        "ComboBox item ScrollIntoView executed. source={Source}, requested={Requested}, item={Item}",
                        source,
                        requestedValue,
                        SafeElementName(item));
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "ComboBox ScrollIntoView failed. source={Source}, requested={Requested}, item={Item}",
                        source,
                        requestedValue,
                        SafeElementName(item));
                }
            }

            if (capabilities.HasSelectionItem && item.Patterns.SelectionItem.IsSupported)
            {
                try
                {
                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "SelectionItem commit"))
                    {
                        return false;
                    }

                    item.Patterns.SelectionItem.Pattern.Select();
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    // If the dropdown has already collapsed, verify via the normal stable-collapse path.
                    // If the dropdown is still expanded, SelectionItem.Select() only highlighted the item
                    // without closing the dropdown — skip the expensive collapse wait and immediately
                    // press Enter to finalise the highlighted selection.
                    if (!IsComboBoxExpanded(comboBox))
                    {
                        if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                        {
                            _logger.LogInformation(
                                "ComboBox item committed using SelectionItemPattern. source={Source}, requested={Requested}",
                                source,
                                requestedValue);

                            return true;
                        }
                    }
                    else
                    {
                        // Dropdown still open — check if item is highlighted and press Enter to commit.
                        var isItemHighlighted = false;
                        try
                        {
                            isItemHighlighted = item.Patterns.SelectionItem.IsSupported &&
                                                item.Patterns.SelectionItem.Pattern.IsSelected;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogDebug(
                                ex,
                                "ComboBox SelectionItem.IsSelected read failed. source={Source}, requested={Requested}",
                                source,
                                requestedValue);
                        }

                        if (isItemHighlighted &&
                            IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "SelectionItem highlight Enter"))
                        {
                            _logger.LogInformation(
                                "ComboBox dropdown still expanded after SelectionItem.Select with item highlighted — pressing Enter. source={Source}, requested={Requested}",
                                source,
                                requestedValue);

                            Keyboard.Press(VirtualKeyShort.RETURN);
                            Keyboard.Release(VirtualKeyShort.RETURN);
                            Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                            if (VerifyComboBoxValueAfterEnterWithDropdownStateCheck(session, comboBox, requestedValue, source + "-selectionitem-enter"))
                            {
                                _logger.LogInformation(
                                    "ComboBox item committed by SelectionItem highlight + Enter. source={Source}, requested={Requested}",
                                    source,
                                    requestedValue);

                                return true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "ComboBox SelectionItemPattern.Select failed. source={Source}, requested={Requested}, item={Item}",
                        source,
                        requestedValue,
                        SafeElementName(item));
                }
            }

            if (capabilities.HasInvoke && item.Patterns.Invoke.IsSupported)
            {
                try
                {
                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "Invoke commit"))
                    {
                        return false;
                    }

                    item.Patterns.Invoke.Pattern.Invoke();
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    // If the dropdown has already collapsed, verify via the normal stable-collapse path.
                    // If the dropdown is still expanded, Invoke() did not close it — skip the expensive
                    // collapse wait and immediately press Enter to finalise the selection.
                    if (!IsComboBoxExpanded(comboBox))
                    {
                        if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                        {
                            _logger.LogInformation(
                                "ComboBox item committed using InvokePattern. source={Source}, requested={Requested}",
                                source,
                                requestedValue);

                            return true;
                        }
                    }
                    else
                    {
                        // Dropdown still open after Invoke — press Enter to commit.
                        if (IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "Invoke Enter"))
                        {
                            _logger.LogInformation(
                                "ComboBox dropdown still expanded after InvokePattern — pressing Enter. source={Source}, requested={Requested}",
                                source,
                                requestedValue);

                            Keyboard.Press(VirtualKeyShort.RETURN);
                            Keyboard.Release(VirtualKeyShort.RETURN);
                            Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                            if (VerifyComboBoxValueAfterEnterWithDropdownStateCheck(session, comboBox, requestedValue, source + "-invoke-enter"))
                            {
                                _logger.LogInformation(
                                    "ComboBox item committed by Invoke + Enter. source={Source}, requested={Requested}",
                                    source,
                                    requestedValue);

                                return true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "ComboBox InvokePattern.Invoke failed. source={Source}, requested={Requested}, item={Item}",
                        source,
                        requestedValue,
                        SafeElementName(item));
                }
            }

            if (capabilities.HasScrollItem)
            {
                if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                    !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "physical click"))
                {
                    return false;
                }

                if (TryPhysicalClickComboBoxListItem(item, requestedValue, source))
                {
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                    {
                        _logger.LogInformation(
                            "ComboBox item committed using physical click after ScrollIntoView. source={Source}, requested={Requested}",
                            source,
                            requestedValue);

                        return true;
                    }
                }
            }

            // Final fallback: if dropdown is still expanded after all pattern attempts, focus the
            // visible item and press Enter to commit the highlighted selection.
            if (IsComboBoxExpanded(comboBox) &&
                IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) &&
                IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "pattern final Enter"))
            {
                var itemVisible = IsElementVisibleAndClickableEnough(item);

                _logger.LogInformation(
                    "ComboBox dropdown still expanded after all pattern attempts — attempting final Enter fallback. source={Source}, requested={Requested}, itemVisible={ItemVisible}",
                    source,
                    requestedValue,
                    itemVisible);

                if (itemVisible)
                {
                    TryFocusComboBoxListItem(item, requestedValue, source + "-final-enter");
                    Thread.Sleep(ComboBoxAnchorMoveDelayMs);
                }

                if (IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) &&
                    IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "pattern final Enter press"))
                {
                    Keyboard.Press(VirtualKeyShort.RETURN);
                    Keyboard.Release(VirtualKeyShort.RETURN);
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxValueAfterEnterWithDropdownStateCheck(session, comboBox, requestedValue, source + "-final-enter"))
                    {
                        _logger.LogInformation(
                            "ComboBox item committed by final Enter fallback. source={Source}, requested={Requested}",
                            source,
                            requestedValue);

                        return true;
                    }
                }
            }

            _logger.LogWarning(
                "ComboBox pattern-based commit did not verify. source={Source}, requested={Requested}, actual={Actual}, item={Item}, capabilities={Capabilities}",
                source,
                requestedValue,
                GetComboBoxCurrentValue(session, comboBox),
                SafeElementName(item),
                capabilities.ToString());

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox pattern-based commit failed. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
    }

    private bool TryActivateComboBoxItemByUiaPattern(
        AutomationSession session,
        AutomationElement comboBox,
        AutomationElement item,
        string requestedValue,
        string source)
    {
        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();

                _logger.LogInformation(
                    "ComboBox item activated using SelectionItemPattern. source={Source}, requested={Requested}, item={Item}",
                    source,
                    requestedValue,
                    SafeElementName(item));

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                return VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source);
            }

            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();

                _logger.LogInformation(
                    "ComboBox item activated using InvokePattern. source={Source}, requested={Requested}, item={Item}",
                    source,
                    requestedValue,
                    SafeElementName(item));

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                return VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source);
            }

            _logger.LogInformation(
                "ComboBox item does not support SelectionItemPattern or InvokePattern. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Direct UIA ComboBox item activation failed. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
    }

    private bool CommitSmallComboBoxItemUserLike(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline)
    {
        const string source = "small-combobox-userlike-click";

        try
        {
            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "small ComboBox reopen"))
            {
                return false;
            }

            if (!OpenComboBoxDropdown(session, comboBox))
                return false;

            Thread.Sleep(MenuExpandDelayMs);

            var freshItem = FindComboBoxItemByTextWithScroll(
                session,
                comboBox,
                requestedValue,
                maxScrollAttempts: SmallComboBoxMaxScrollAttempts);

            if (freshItem == null)
                return false;

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "physical click"))
            {
                return false;
            }

            if (TryPhysicalClickComboBoxListItem(
                    freshItem,
                    requestedValue,
                    source))
            {
                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                if (VerifyComboBoxSelectedValueStableAfterCollapse(
                        session,
                        comboBox,
                        requestedValue,
                        source))
                {
                    return true;
                }
            }

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "small ComboBox pattern fallback"))
            {
                return false;
            }

            return CommitExactVisibleComboBoxItem(
                session,
                comboBox,
                freshItem,
                requestedValue,
                guard,
                operationDeadline,
                "small-combobox-pattern-fallback");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Small ComboBox user-like commit failed. requested={Requested}, combo={Combo}",
                requestedValue,
                SafeElementName(comboBox));

            return false;
        }
    }

    private bool CommitExactVisibleComboBoxItem(
        AutomationSession session,
        AutomationElement comboBox,
        AutomationElement item,
        string requestedValue,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline,
        string source)
    {
        try
        {
            _logger.LogInformation(
                "Committing exact ComboBox item. source={Source}, combo={Combo}, requested={Requested}, item={Item}",
                source,
                SafeElementName(comboBox),
                requestedValue,
                SafeElementName(item));

            if (item.Patterns.SelectionItem.IsSupported)
            {
                try
                {
                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "SelectionItem commit"))
                    {
                        return false;
                    }

                    item.Patterns.SelectionItem.Pattern.Select();
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                    {
                        _logger.LogInformation(
                            "ComboBox item committed by SelectionItemPattern. source={Source}, requested={Requested}",
                            source,
                            requestedValue);

                        return true;
                    }

                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "ENTER commit"))
                    {
                        return false;
                    }

                    Keyboard.Press(VirtualKeyShort.RETURN);
                    Keyboard.Release(VirtualKeyShort.RETURN);
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source + "-selectionitem-enter"))
                    {
                        _logger.LogInformation(
                            "ComboBox item committed by SelectionItemPattern plus ENTER. source={Source}, requested={Requested}",
                            source,
                            requestedValue);

                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "SelectionItemPattern commit failed. source={Source}, requested={Requested}",
                        source,
                        requestedValue);
                }
            }

            if (item.Patterns.Invoke.IsSupported)
            {
                try
                {
                    if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                        !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "Invoke commit"))
                    {
                        return false;
                    }

                    item.Patterns.Invoke.Pattern.Invoke();
                    Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                    if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                    {
                        _logger.LogInformation(
                            "ComboBox item committed by InvokePattern. source={Source}, requested={Requested}",
                            source,
                            requestedValue);

                        return true;
                    }

                    if (IsComboBoxTabBlurCommitFallbackAllowed())
                    {
                        if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                            !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "TAB blur fallback"))
                        {
                            return false;
                        }

                        Keyboard.Press(VirtualKeyShort.TAB);
                        Keyboard.Release(VirtualKeyShort.TAB);
                        Thread.Sleep(ComboBoxBlurCommitDelayMs);

                        if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source + "-tab-blur"))
                        {
                            _logger.LogInformation(
                                "ComboBox item committed by optional TAB blur fallback. source={Source}, requested={Requested}",
                                source,
                                requestedValue);

                            return true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(
                        ex,
                        "InvokePattern commit failed. source={Source}, requested={Requested}",
                        source,
                        requestedValue);
                }
            }

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "physical click"))
            {
                return false;
            }

            if (TryPhysicalClickComboBoxListItem(item, requestedValue, source))
            {
                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                {
                    _logger.LogInformation(
                        "ComboBox item committed by physical click. source={Source}, requested={Requested}",
                        source,
                        requestedValue);

                    return true;
                }
            }

            if (TryFocusComboBoxListItem(item, requestedValue, source))
            {
                Thread.Sleep(ComboBoxAnchorMoveDelayMs);

                if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                    !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "ENTER commit"))
                {
                    return false;
                }

                Keyboard.Press(VirtualKeyShort.RETURN);
                Keyboard.Release(VirtualKeyShort.RETURN);

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                {
                    _logger.LogInformation(
                        "ComboBox item committed by ENTER. source={Source}, requested={Requested}",
                        source,
                        requestedValue);

                    return true;
                }
            }

            _logger.LogWarning(
                "ComboBox exact item was visible but commit did not verify. source={Source}, requested={Requested}, actual={Actual}, item={Item}",
                source,
                requestedValue,
                GetComboBoxCurrentValue(session, comboBox),
                SafeElementName(item));

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Commit exact ComboBox item failed. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
    }

    private bool TryPhysicalClickComboBoxListItem(
        AutomationElement item,
        string requestedValue,
        string source)
    {
        try
        {
            var rect = item.BoundingRectangle;

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            {
                _logger.LogInformation(
                    "ComboBox item physical click skipped because item rectangle is invalid. source={Source}, requested={Requested}, item={Item}",
                    source,
                    requestedValue,
                    SafeElementName(item));

                return false;
            }

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            _logger.LogInformation(
                "Physical clicking exact ComboBox ListItem. source={Source}, requested={Requested}, item={Item}, point={Point}",
                source,
                requestedValue,
                SafeElementName(item),
                point);

            return TryPhysicalClickPoint(point, $"ComboBox Commit {source}");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Physical click exact ComboBox ListItem failed. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
    }

    private bool TryFocusComboBoxListItem(
        AutomationElement item,
        string requestedValue,
        string source)
    {
        try
        {
            item.Focus();

            _logger.LogInformation(
                "Focused exact ComboBox ListItem before ENTER commit. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "Focus exact ComboBox ListItem failed. source={Source}, requested={Requested}, item={Item}",
                source,
                requestedValue,
                SafeElementName(item));

            return false;
        }
    }

    private bool TrySetComboBoxValueByValuePattern(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        string source)
    {
        try
        {
            if (comboBox.Patterns.Value.IsSupported)
            {
                comboBox.Patterns.Value.Pattern.SetValue(requestedValue);

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                Keyboard.Press(VirtualKeyShort.RETURN);
                Keyboard.Release(VirtualKeyShort.RETURN);

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                {
                    _logger.LogInformation(
                        "ComboBox value selected using ComboBox ValuePattern.SetValue. combo={Combo}, value={Value}",
                        SafeElementName(comboBox),
                        requestedValue);

                    return true;
                }
            }

            var childEdit = FindComboBoxChildEdit(comboBox);

            if (childEdit != null && childEdit.Patterns.Value.IsSupported)
            {
                childEdit.Patterns.Value.Pattern.SetValue(requestedValue);

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                Keyboard.Press(VirtualKeyShort.RETURN);
                Keyboard.Release(VirtualKeyShort.RETURN);

                Thread.Sleep(ComboBoxSelectionCommitDelayMs);

                if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, source))
                {
                    _logger.LogInformation(
                        "ComboBox value selected using child Edit ValuePattern.SetValue. combo={Combo}, edit={Edit}, value={Value}",
                        SafeElementName(comboBox),
                        SafeElementName(childEdit),
                        requestedValue);

                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox ValuePattern.SetValue strategy failed. combo={Combo}, value={Value}",
                SafeElementName(comboBox),
                requestedValue);

            return false;
        }
    }

    private AutomationElement? FindComboBoxChildEdit(AutomationElement comboBox)
    {
        try
        {
            var children = comboBox.FindAllChildren();

            foreach (var child in children)
            {
                try
                {
                    if (child.ControlType == ControlType.Edit)
                        return child;
                }
                catch
                {
                    // ignore stale child
                }
            }
        }
        catch
        {
            // ignore
        }

        return null;
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
        string itemName,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline)
    {
        var requested = NormalizeMenuText(itemName);
        if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
            !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "paged visible search reset"))
        {
            return false;
        }

        if (!ResetComboBoxDropdownToTop(session, comboBox))
            return false;

        for (var page = 0; page < ComboBoxPagedSearchMaxPages; page++)
        {
            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "paged visible search"))
            {
                return false;
            }

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

                    var matchCapabilities = DetectComboBoxItemPatternCapabilities(item);
                    if (matchCapabilities.HasAnyUsefulPattern &&
                        TryCommitComboBoxItemUsingAvailablePatterns(
                            session,
                            comboBox,
                            item,
                            itemName,
                            matchCapabilities,
                            guard,
                            operationDeadline,
                            "visible-search-pattern-based"))
                    {
                        return true;
                    }

                    return CommitExactVisibleComboBoxItem(
                        session,
                        comboBox,
                        item,
                        itemName,
                        guard,
                        operationDeadline,
                        "huge-list-visible-search");
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
        string itemName,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(itemName))
                return false;

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "keyboard focus"))
            {
                return false;
            }

            BringElementWindowToForeground(comboBox);
            Thread.Sleep(WindowActivationDelayMs);

            if (!IsComboBoxTargetGuardValid(comboBox, guard, itemName, "keyboard open"))
                return false;

            if (!FocusElementForKeyboardInput(comboBox, "ComboBoxTypeAhead"))
            {
                _logger.LogWarning(
                    "ComboBox keyboard type-ahead skipped because focus could not be confirmed. combo={Combo}, item={Item}",
                    SafeElementName(comboBox),
                    itemName);

                return false;
            }

            Thread.Sleep(ComboBoxTypeAheadFocusDelayMs);

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "keyboard reopen") ||
                !OpenComboBoxDropdown(session, comboBox))
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

            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, itemName) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, itemName, "keyboard type-ahead"))
            {
                return false;
            }

            // Use Keyboard.Type for literal character-by-character input so that special
            // characters in item names (e.g. "+", "^", "%") are never misinterpreted as
            // modifier keys the way SendKeysString would treat them.
            Keyboard.Type(itemName);

            Thread.Sleep(ComboBoxTypeAheadCommitDelayMs);

            if (!IsComboBoxTargetGuardValid(comboBox, guard, itemName, "ENTER commit"))
                return false;

            Keyboard.Press(VirtualKeyShort.RETURN);
            Keyboard.Release(VirtualKeyShort.RETURN);

            Thread.Sleep(ComboBoxTypeAheadCommitDelayMs);

            var verified = VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, itemName, "keyboard-typeahead");

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
        string requestedValue,
        ComboBoxTargetGuard guard,
        DateTime operationDeadline)
    {
        try
        {
            if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "visible anchor reset"))
            {
                return false;
            }

            if (!ResetComboBoxDropdownToTop(session, comboBox))
                return false;

            string? previousSignature = null;

            for (var window = 0; window < ComboBoxAnchorWindowSearchMaxWindows; window++)
            {
                if (!IsComboBoxOperationWithinDeadline(operationDeadline, comboBox, requestedValue) ||
                    !IsComboBoxTargetGuardValid(comboBox, guard, requestedValue, "visible anchor reopen"))
                {
                    return false;
                }

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

                    var matchCapabilities = DetectComboBoxItemPatternCapabilities(exactItem);
                    if (matchCapabilities.HasAnyUsefulPattern &&
                        TryCommitComboBoxItemUsingAvailablePatterns(
                            session,
                            comboBox,
                            exactItem,
                            requestedValue,
                            matchCapabilities,
                            guard,
                            operationDeadline,
                            "visible-search-pattern-based"))
                    {
                        return true;
                    }

                    return CommitExactVisibleComboBoxItem(
                        session,
                        comboBox,
                        exactItem,
                        requestedValue,
                        guard,
                        operationDeadline,
                        "huge-list-visible-search");
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

                    if (VerifyComboBoxSelectedValueStableAfterCollapse(session, comboBox, requestedValue, "keyboard-step-search"))
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

    private bool WaitForComboBoxDropdownCollapsed(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        string source)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(ComboBoxPostCommitCollapseTimeoutMs);

        _logger.LogInformation(
            "Waiting for ComboBox dropdown collapse. source={Source}, requested={Requested}, combo={Combo}",
            source,
            requestedValue,
            SafeElementName(comboBox));

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var isExpanded = false;

                if (comboBox.Patterns.ExpandCollapse.IsSupported)
                {
                    var state = comboBox.Patterns.ExpandCollapse.Pattern.ExpandCollapseState;
                    isExpanded = state == ExpandCollapseState.Expanded;
                }

                var popupList = FindDynamicComboBoxList(session, comboBox);
                var popupStillVisible = popupList != null && IsElementVisibleOnScreen(popupList);

                if (!isExpanded && !popupStillVisible)
                {
                    _logger.LogInformation(
                        "ComboBox dropdown collapsed after commit. source={Source}, requested={Requested}, combo={Combo}",
                        source,
                        requestedValue,
                        SafeElementName(comboBox));

                    return true;
                }
            }
            catch
            {
                // Stale popup/combo is okay during collapse. Continue polling.
            }

            Thread.Sleep(ComboBoxPostCommitPollDelayMs);
        }

        _logger.LogWarning(
            "Timed out waiting for ComboBox dropdown to collapse. source={Source}, requested={Requested}, combo={Combo}",
            source,
            requestedValue,
            SafeElementName(comboBox));

        return false;
    }

    /// <summary>
    /// Checks whether the combobox committed the requested value after an Enter press,
    /// without requiring the dropdown to have collapsed first.  The current dropdown
    /// expand/collapse state is read and logged so callers can diagnose stuck-open dropdowns.
    /// </summary>
    private bool VerifyComboBoxValueAfterEnterWithDropdownStateCheck(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        string source)
    {
        try
        {
            Thread.Sleep(ComboBoxPostCommitStableDelayMs);

            var expandState = GetComboBoxExpandState(comboBox);
            var freshComboBox = RefreshComboBoxElement(session, comboBox) ?? comboBox;
            var actual = GetComboBoxCurrentValue(session, freshComboBox);
            var matched = ComboBoxValueMatches(actual, requestedValue);

            _logger.LogInformation(
                "ComboBox value check after Enter commit. source={Source}, requested={Requested}, actual={Actual}, matched={Matched}, dropdownState={DropdownState}",
                source,
                requestedValue,
                actual,
                matched,
                expandState);

            if (matched && string.Equals(expandState, "Expanded", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogInformation(
                    "ComboBox value matched but dropdown still expanded after Enter. source={Source}, requested={Requested}",
                    source,
                    requestedValue);
            }

            return matched;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox value verification after Enter failed. source={Source}, requested={Requested}",
                source,
                requestedValue);

            return false;
        }
    }

    private bool IsElementVisibleOnScreen(AutomationElement element)
    {
        try
        {
            var rect = element.BoundingRectangle;

            return !rect.IsEmpty &&
                   rect.Width > 0 &&
                   rect.Height > 0 &&
                   rect.Right > 0 &&
                   rect.Bottom > 0;
        }
        catch
        {
            return false;
        }
    }

    private bool VerifyComboBoxSelectedValueStableAfterCollapse(
        AutomationSession session,
        AutomationElement comboBox,
        string requestedValue,
        string source)
    {
        try
        {
            if (!WaitForComboBoxDropdownCollapsed(session, comboBox, requestedValue, source))
                return false;

            // Important: allow app event handlers to finish after collapse.
            Thread.Sleep(ComboBoxPostCommitStableDelayMs);

            var freshComboBox = RefreshComboBoxElement(session, comboBox) ?? comboBox;
            var firstActual = GetComboBoxCurrentValue(session, freshComboBox);
            var firstMatched = ComboBoxValueMatches(firstActual, requestedValue);

            _logger.LogInformation(
                "ComboBox post-collapse verify first read. source={Source}, requested={Requested}, actual={Actual}, matched={Matched}",
                source,
                requestedValue,
                firstActual,
                firstMatched);

            if (!firstMatched)
            {
                _logger.LogWarning(
                    "ComboBox rollback/default detected after dropdown collapse. source={Source}, requested={Requested}, actual={Actual}, combo={Combo}",
                    source,
                    requestedValue,
                    GetComboBoxCurrentValue(session, comboBox),
                    SafeElementName(comboBox));

                return false;
            }

            // Read again after a short delay to catch rollback/default reset.
            Thread.Sleep(ComboBoxPostCommitStableDelayMs);

            freshComboBox = RefreshComboBoxElement(session, freshComboBox) ?? freshComboBox;
            var secondActual = GetComboBoxCurrentValue(session, freshComboBox);
            var secondMatched = ComboBoxValueMatches(secondActual, requestedValue);

            _logger.LogInformation(
                "ComboBox post-collapse verify second read. source={Source}, requested={Requested}, actual={Actual}, matched={Matched}",
                source,
                requestedValue,
                secondActual,
                secondMatched);

            if (secondMatched)
            {
                _logger.LogInformation(
                    "ComboBox selection committed and stable. No further fallback will run. source={Source}, combo={Combo}, requested={Requested}",
                    source,
                    SafeElementName(comboBox),
                    requestedValue);

                return true;
            }

            _logger.LogWarning(
                "ComboBox rollback/default detected after dropdown collapse. source={Source}, requested={Requested}, actual={Actual}, combo={Combo}",
                source,
                requestedValue,
                GetComboBoxCurrentValue(session, comboBox),
                SafeElementName(comboBox));

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "ComboBox stable post-collapse verification failed. source={Source}, requested={Requested}, combo={Combo}",
                source,
                requestedValue,
                SafeElementName(comboBox));

            return false;
        }
    }

    private AutomationElement? RefreshComboBoxElement(AutomationSession session, AutomationElement comboBox)
    {
        try
        {
            var originalName = SafeElementName(comboBox);
            var originalAutomationId = SafeElementAutomationId(comboBox);
            var originalClassName = SafeElementClassName(comboBox);
            var originalProcessId = SafeProcessId(comboBox);
            var originalRect = comboBox.BoundingRectangle;
            var root = FindWindowAncestorOrSelf(comboBox) ?? session.Automation.GetDesktop();
            var cf = session.Automation.ConditionFactory;

            foreach (var candidate in root.FindAllDescendants(cf.ByControlType(ControlType.ComboBox)))
            {
                if (originalProcessId.HasValue && SafeProcessId(candidate) != originalProcessId)
                    continue;

                var automationId = SafeElementAutomationId(candidate);
                var name = SafeElementName(candidate);
                var className = SafeElementClassName(candidate);

                if (!string.IsNullOrWhiteSpace(originalAutomationId) &&
                    string.Equals(automationId, originalAutomationId, StringComparison.Ordinal) &&
                    (string.IsNullOrWhiteSpace(originalClassName) ||
                     string.Equals(className, originalClassName, StringComparison.Ordinal)))
                {
                    return candidate;
                }

                if (!string.IsNullOrWhiteSpace(originalName) &&
                    string.Equals(name, originalName, StringComparison.Ordinal) &&
                    IsSameApproximateRectangle(candidate, originalRect))
                {
                    return candidate;
                }
            }
        }
        catch
        {
            // Best effort only; callers can still use the original element.
        }

        return null;
    }

    private static bool IsSameApproximateRectangle(AutomationElement element, Rectangle originalRect)
    {
        try
        {
            var rect = element.BoundingRectangle;

            return !rect.IsEmpty &&
                   Math.Abs(rect.Left - originalRect.Left) <= ComboBoxRefetchRectangleTolerancePx &&
                   Math.Abs(rect.Top - originalRect.Top) <= ComboBoxRefetchRectangleTolerancePx &&
                   Math.Abs(rect.Width - originalRect.Width) <= ComboBoxRefetchRectangleTolerancePx &&
                   Math.Abs(rect.Height - originalRect.Height) <= ComboBoxRefetchRectangleTolerancePx;
        }
        catch
        {
            return false;
        }
    }

    private bool ComboBoxValueMatches(string? actualValue, string requestedValue)
    {
        var actual = NormalizeMenuText(actualValue ?? string.Empty);
        var expected = NormalizeMenuText(requestedValue ?? string.Empty);

        return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
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

    /// <summary>
    /// Score model used to rank context-menu popup candidates by proximity and ownership.
    /// </summary>
    private sealed class ContextMenuCandidateScore
    {
        public required AutomationElement Element { get; init; }
        public int Score { get; init; }
        public string Reason { get; init; } = string.Empty;
    }

    /// <summary>
    /// Returns a numeric score for <paramref name="candidate"/> based on how likely it is to be
    /// the context menu that appeared after a right-click at <paramref name="rightClickPoint"/>
    /// inside the window identified by <paramref name="targetHwnd"/>.
    /// Higher scores are better. Returns -1000 for candidates with no valid bounding rectangle
    /// or when an exception occurs; candidates with a score at or below
    /// <see cref="ContextMenuMinimumCandidateScore"/> are dropped by the caller.
    /// </summary>
    private int ScoreContextMenuCandidate(
        AutomationElement candidate,
        Point rightClickPoint,
        IntPtr targetHwnd,
        out string reason)
    {
        reason = string.Empty;

        try
        {
            var score = 0;
            var rect = candidate.BoundingRectangle;
            var hwnd = SafeWindowHandle(candidate);

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                return -1000;

            var nearClick =
                rightClickPoint.X >= rect.Left - ContextMenuNearClickMarginLeading &&
                rightClickPoint.X <= rect.Right + ContextMenuNearClickMarginTrailing &&
                rightClickPoint.Y >= rect.Top - ContextMenuNearClickMarginLeading &&
                rightClickPoint.Y <= rect.Bottom + ContextMenuNearClickMarginTrailing;

            if (nearClick)
            {
                score += 100;
                reason += "near-click;";
            }

            var dx = Math.Abs(rect.Left - rightClickPoint.X);
            var dy = Math.Abs(rect.Top - rightClickPoint.Y);

            if (dx <= ContextMenuTopLeftProximityThreshold && dy <= ContextMenuTopLeftProximityThreshold)
            {
                score += 80;
                reason += "top-left-near-click;";
            }

            if (targetHwnd != IntPtr.Zero)
            {
                if (hwnd == targetHwnd)
                {
                    score += 80;
                    reason += "same-hwnd;";
                }
                else
                {
                    try
                    {
                        var root = hwnd != IntPtr.Zero ? GetAncestor(hwnd, GA_ROOT) : IntPtr.Zero;
                        if (root != IntPtr.Zero && root == targetHwnd)
                        {
                            score += 80;
                            reason += "same-root-hwnd;";
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            if (!nearClick && dx > ContextMenuFarFromClickThreshold && dy > ContextMenuFarFromClickThreshold)
            {
                score -= 100;
                reason += "far-from-click;";
            }

            var name = SafeElementName(candidate);
            var className = SafeElementClassName(candidate);

            if (name.Contains("Start", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Volume", StringComparison.OrdinalIgnoreCase) ||
                className.Contains("Shell", StringComparison.OrdinalIgnoreCase) ||
                className.Contains("Tray", StringComparison.OrdinalIgnoreCase))
            {
                score -= 200;
                reason += "system-shell-penalty;";
            }

            var itemCount = GetContextMenuItems(candidate, maxItems: 5).Count;
            if (itemCount > 0)
            {
                score += 20;
                reason += $"items={itemCount};";
            }

            return score;
        }
        catch
        {
            reason = "exception;";
            return -1000;
        }
    }

    private List<AutomationElement> GetContextMenuItems(
        AutomationElement menuRoot,
        int maxItems)
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

    private AutomationElement? FindActiveContextMenuPopup(
        AutomationSession session,
        Point rightClickPoint,
        IntPtr targetHwnd)
    {
        try
        {
            var candidates = GetContextMenuPopupCandidates(session, rightClickPoint, targetHwnd);
            var scored = new List<ContextMenuCandidateScore>();

            foreach (var candidate in candidates)
            {
                var items = GetContextMenuItems(session, candidate, maxItems: 1);
                if (items.Count == 0)
                    continue;

                var score = ScoreContextMenuCandidate(
                    candidate,
                    rightClickPoint,
                    targetHwnd,
                    out var reason);

                scored.Add(new ContextMenuCandidateScore
                {
                    Element = candidate,
                    Score = score,
                    Reason = reason
                });
            }

            var best = scored.OrderByDescending(x => x.Score).FirstOrDefault();

            if (best == null)
                return null;

            _logger.LogInformation(
                "Selected context menu popup candidate. score={Score}, reason={Reason}, name={Name}, rect={Rect}, targetHwnd=0x{TargetHwnd:X}, click={ClickPoint}",
                best.Score,
                best.Reason,
                SafeElementName(best.Element),
                best.Element.BoundingRectangle,
                targetHwnd.ToInt64(),
                rightClickPoint);

            return best.Element;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to detect scoped active context menu popup.");
            return null;
        }
    }

    private AutomationElement? FindContextSubMenuPopup(
        AutomationSession session,
        AutomationElement submenuItem,
        Point rightClickPoint,
        IntPtr targetHwnd)
    {
        try
        {
            var itemRect = submenuItem.BoundingRectangle;

            foreach (var candidate in GetContextMenuPopupCandidates(session, rightClickPoint, targetHwnd))
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

    private List<AutomationElement> GetContextMenuPopupCandidates(
        AutomationSession session,
        Point rightClickPoint,
        IntPtr targetHwnd)
    {
        var results = new List<AutomationElement>();

        try
        {
            var driverPid = Environment.ProcessId;
            var targetPid = session.Application.ProcessId;
            var desktop = session.Automation.GetDesktop();
            var queue = new Queue<(AutomationElement Element, int Depth)>();

            foreach (var child in desktop.FindAllChildren())
            {
                try
                {
                    var childPid = child.Properties.ProcessId.Value;

                    if (childPid == driverPid)
                        continue;

                    if (childPid != targetPid)
                        continue;
                }
                catch
                {
                    continue;
                }

                queue.Enqueue((child, 1));
            }

            while (queue.Count > 0 && results.Count < MaxApplicationContextMenuCandidates)
            {
                var (element, depth) = queue.Dequeue();

                try
                {
                    var ct = element.ControlType;
                    var rect = element.BoundingRectangle;

                    if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                        continue;

                    if (IsContextMenuContainerType(ct))
                    {
                        var score = ScoreContextMenuCandidate(
                            element,
                            rightClickPoint,
                            targetHwnd,
                            out _);

                        // Drop obviously unrelated candidates (system-shell / far + penalised).
                        if (score > ContextMenuMinimumCandidateScore)
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
            _logger.LogWarning(ex, "Failed collecting scoped context menu popup candidates.");
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
        // 1. Physical click first.
        // This matches real user behavior and avoids blocking InvokePattern
        // when the menu item opens a modal/new window/export screen.
        try
        {
            var rect = item.BoundingRectangle;

            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                var x = (int)Math.Round(rect.Left + rect.Width / 2.0);
                var y = (int)Math.Round(rect.Top + rect.Height / 2.0);
                var point = new Point(x, y);

                _logger.LogInformation(
                    "Activating context menu item by physical click first. item={Item}, x={X}, y={Y}",
                    itemName,
                    x,
                    y);

                if (SendInstantLeftClick(point, $"Context menu item: {itemName}"))
                {
                    Thread.Sleep(MenuActionDelayMs);

                    _logger.LogInformation(
                        "Context menu item activated by physical click. item={Item}",
                        itemName);

                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Physical click failed for context menu item {Item}; trying pattern fallback.",
                itemName);
        }

        // 2. InvokePattern fallback only.
        // Do not use Invoke first because it can block if the action opens a new window/modal.
        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                _logger.LogInformation(
                    "Activating context menu item by InvokePattern fallback. item={Item}",
                    itemName);

                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "InvokePattern fallback failed for context menu item {Item}.",
                itemName);
        }

        // 3. SelectionItem fallback.
        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                _logger.LogInformation(
                    "Activating context menu item by SelectionItem fallback. item={Item}",
                    itemName);

                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(MenuActionDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "SelectionItem fallback failed for context menu item {Item}.",
                itemName);
        }

        return false;
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
        return OpenHeaderDropdownAndFindContainer(header, region);
    }

    private AutomationElement? OpenHeaderDropdownAndFindContainer(
        AutomationElement header,
        HeaderDropdownRegion region)
    {
        BringElementWindowToForeground(header);
        Thread.Sleep(WindowActivationDelayMs);

        var rect = header.BoundingRectangle;

        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            throw new InvalidOperationException("Header has invalid bounding rectangle.");

        AutomationElement? container = null;
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

            container = FindOpenDropdownContainerNear(header);
            if (container != null)
                return container;
        }

        Thread.Sleep(GridHeaderDropdownHelper.DropdownOpenDelayMs);
        return FindOpenDropdownContainerNear(header);
    }

    private AutomationElement? FindRecentlyOpenedListNearHeader(AutomationElement header) =>
        FindOpenDropdownContainerNear(header);

    private AutomationElement? FindOpenDropdownContainerNear(AutomationElement? nearElement)
    {
        try
        {
            var session = RequireSession();
            var nearRect = nearElement?.BoundingRectangle;
            var desktop = session.Automation.GetDesktop();
            var cf = session.Automation.ConditionFactory;

            var containerTypes = new[]
            {
                ControlType.List,
                ControlType.Menu,
                ControlType.Pane,
                ControlType.Window,
                ControlType.Custom,
                ControlType.Tree,
                ControlType.DataGrid
            };

            AutomationElement? best = null;
            var bestScore = int.MinValue;

            foreach (var containerType in containerTypes)
            {
                foreach (var candidate in desktop.FindAllDescendants(cf.ByControlType(containerType)))
                {
                    try
                    {
                        var rect = candidate.BoundingRectangle;
                        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                            continue;

                        if (nearRect.HasValue && !GridHeaderDropdownHelper.IsListNearHeader(rect, nearRect.Value))
                            continue;

                        var score = (int)rect.Width + (int)rect.Height;
                        if (GetDropdownSelectableItems(candidate).Count == 0)
                            score -= 1000;

                        if (score > bestScore)
                        {
                            bestScore = score;
                            best = candidate;
                        }
                    }
                    catch
                    {
                        // ignore unstable popup elements
                    }
                }
            }

            if (best != null)
            {
                _logger.LogInformation(
                    "Found open dropdown container. container={Container}, controlType={ControlType}, bounds={Bounds}",
                    SafeElementName(best),
                    best.ControlType,
                    best.BoundingRectangle);
            }

            return best;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindOpenDropdownContainerNear failed");
            return null;
        }
    }

    private static readonly ControlType[] DropdownSelectableItemTypes =
    [
        ControlType.ListItem,
        ControlType.CheckBox,
        ControlType.RadioButton,
        ControlType.MenuItem,
        ControlType.Text,
        ControlType.DataItem,
        ControlType.TreeItem,
        ControlType.Custom,
        ControlType.Button
    ];

    private List<AutomationElement> GetDropdownSelectableItems(AutomationElement container)
    {
        try
        {
            var cf = RequireSession().Automation.ConditionFactory;
            var results = new List<AutomationElement>();

            foreach (var itemType in DropdownSelectableItemTypes)
            {
                results.AddRange(container.FindAllDescendants(cf.ByControlType(itemType)));
            }

            return results
                .Where(e =>
                {
                    try
                    {
                        return SafeIsOffscreen(e) != true;
                    }
                    catch
                    {
                        return false;
                    }
                })
                .DistinctBy(SafeRuntimeIdString)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private AutomationElement? FindDropdownSelectableItem(
        IReadOnlyList<AutomationElement> items,
        string? value,
        int? index,
        string matchMode)
    {
        if (index.HasValue)
        {
            return index.Value >= 0 && index.Value < items.Count ? items[index.Value] : null;
        }

        if (string.IsNullOrWhiteSpace(value))
            return null;

        return items.FirstOrDefault(item =>
            MatchesAnyText(
                SafeElementName(item),
                SafeElementValue(item),
                SafeElementText(item),
                value,
                matchMode));
    }

    private List<(string name, string controlType, object rectangle, object patterns)> BuildDropdownItemSummaries(
        IReadOnlyList<AutomationElement> items)
    {
        return items.Select(item => (
            name: SafeElementName(item),
            controlType: SafeElementControlType(item),
            rectangle: SafeBoundingRectangleObject(item),
            patterns: (object)new
            {
                toggle = item.Patterns.Toggle.IsSupported,
                selectionItem = item.Patterns.SelectionItem.IsSupported,
                invoke = item.Patterns.Invoke.IsSupported
            })).ToList();
    }

    private sealed record OpenDropdownActivationResult(
        bool Success,
        string MatchedText,
        string ActivationStrategy,
        bool Verified,
        string VerificationReason,
        object Container,
        object MatchedItem);

    private OpenDropdownActivationResult ActivateOpenDropdownItemSoft(
        AutomationElement item,
        string itemName,
        DropdownItemClickRegion region)
    {
        var container = item.Parent ?? item;
        var matchedSummary = new
        {
            name = SafeElementName(item),
            controlType = SafeElementControlType(item),
            rectangle = SafeBoundingRectangleObject(item)
        };
        var containerSummary = new
        {
            controlType = SafeElementControlType(container),
            className = SafeElementClassName(container),
            rectangle = SafeBoundingRectangleObject(container)
        };

        if (item.Patterns.Toggle.IsSupported)
        {
            try
            {
                item.Patterns.Toggle.Pattern.Toggle();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                return new OpenDropdownActivationResult(
                    true,
                    itemName,
                    "toggle-pattern",
                    item.Patterns.Toggle.Pattern.ToggleState == ToggleState.On,
                    "ToggleState changed",
                    containerSummary,
                    matchedSummary);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ActivateOpenDropdownItemSoft toggle failed.");
            }
        }

        if (item.Patterns.SelectionItem.IsSupported)
        {
            try
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                var verified = item.Patterns.SelectionItem.Pattern.IsSelected;
                return new OpenDropdownActivationResult(
                    true,
                    itemName,
                    "selection-item-pattern",
                    verified,
                    verified ? "SelectionItem.IsSelected" : "SelectionItem select sent",
                    containerSummary,
                    matchedSummary);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ActivateOpenDropdownItemSoft SelectionItem failed.");
            }
        }

        if (item.Patterns.Invoke.IsSupported)
        {
            try
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                return new OpenDropdownActivationResult(
                    true,
                    itemName,
                    "invoke-pattern",
                    false,
                    "Invoke sent; UIA state not exposed",
                    containerSummary,
                    matchedSummary);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ActivateOpenDropdownItemSoft Invoke failed.");
            }
        }

        foreach (var candidateRegion in GetDropdownItemRegionOrder(region))
        {
            try
            {
                var point = GetDropdownItemClickPoint(item.BoundingRectangle, candidateRegion);
                if (SendInstantLeftClick(point, $"SelectOpenDropdownItem {SanitizeValue(itemName)} at {candidateRegion}"))
                {
                    Thread.Sleep(DropdownItemPhysicalClickSettleMs);
                    return new OpenDropdownActivationResult(
                        true,
                        itemName,
                        $"physical-click-{candidateRegion.ToString().ToLowerInvariant()}",
                        false,
                        "physical click sent; UIA state not exposed",
                        containerSummary,
                        matchedSummary);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ActivateOpenDropdownItemSoft physical click failed.");
            }
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.SPACE);
            Thread.Sleep(DropdownItemFallbackDelayMs);
            return new OpenDropdownActivationResult(
                true,
                itemName,
                "focus-space",
                false,
                "Focus+Space sent; UIA state not exposed",
                containerSummary,
                matchedSummary);
        }
        catch
        {
            // continue
        }

        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.RETURN);
            Thread.Sleep(DropdownItemFallbackDelayMs);
            return new OpenDropdownActivationResult(
                true,
                itemName,
                "focus-enter",
                false,
                "Focus+Enter sent; UIA state not exposed",
                containerSummary,
                matchedSummary);
        }
        catch
        {
            // continue
        }

        return new OpenDropdownActivationResult(
            false,
            itemName,
            "none",
            false,
            "All activation strategies failed",
            containerSummary,
            matchedSummary);
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

    private bool TryPhysicalClickPoint(Point point, string actionName)
    {
        return SendInstantLeftClick(point, actionName);
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

    // =========================================================================
    // Scroll helpers
    // =========================================================================

    private static string NormalizeScrollMode(UiRequest request)
    {
        var op = request.Operation?.Trim().ToLowerInvariant();

        if (op is "wheelscroll" or "mousescroll")
            return "wheel";

        var mode = string.IsNullOrWhiteSpace(request.Mode)
            ? "auto"
            : request.Mode.Trim().ToLowerInvariant();

        return mode switch
        {
            "auto"    => "auto",
            "wheel"   => "wheel",
            "mouse"   => "wheel",
            "pattern" => "pattern",
            "uia"     => "pattern",
            _ => throw new ArgumentException(
                $"Unsupported scroll mode '{request.Mode}'. Supported: auto, wheel, pattern.")
        };
    }

    private static string NormalizeScrollDirection(string? direction)
    {
        if (string.IsNullOrWhiteSpace(direction))
            return "down";

        var normalized = direction.Trim().ToLowerInvariant();

        return normalized switch
        {
            "up"    => "up",
            "down"  => "down",
            "left"  => "left",
            "right" => "right",
            _ => throw new ArgumentException(
                $"Unsupported scroll direction '{direction}'. Supported: up, down, left, right.")
        };
    }

    private static string ReverseScrollDirection(string direction) =>
        direction switch
        {
            "up"    => "down",
            "down"  => "up",
            "left"  => "right",
            "right" => "left",
            _       => direction
        };

    private bool TryScrollByPattern(
        AutomationElement element,
        string direction,
        int amount,
        out string strategy)
    {
        strategy = "uia-scrollpattern";

        try
        {
            var scrollPattern = element.Patterns.Scroll.PatternOrDefault;

            if (scrollPattern == null)
            {
                strategy = "uia-scrollpattern-not-supported";
                return false;
            }

            for (var i = 0; i < amount; i++)
            {
                switch (direction)
                {
                    case "up":
                        if (!scrollPattern.VerticallyScrollable)
                            return false;
                        scrollPattern.Scroll(ScrollAmount.NoAmount, ScrollAmount.SmallDecrement);
                        break;

                    case "down":
                        if (!scrollPattern.VerticallyScrollable)
                            return false;
                        scrollPattern.Scroll(ScrollAmount.NoAmount, ScrollAmount.SmallIncrement);
                        break;

                    case "left":
                        if (!scrollPattern.HorizontallyScrollable)
                            return false;
                        scrollPattern.Scroll(ScrollAmount.SmallDecrement, ScrollAmount.NoAmount);
                        break;

                    case "right":
                        if (!scrollPattern.HorizontallyScrollable)
                            return false;
                        scrollPattern.Scroll(ScrollAmount.SmallIncrement, ScrollAmount.NoAmount);
                        break;
                }

                Thread.Sleep(50);
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(
                ex,
                "UIA ScrollPattern failed. element={Element}, direction={Direction}, amount={Amount}",
                SafeElementName(element),
                SanitizeValue(direction),
                amount);

            strategy = "uia-scrollpattern-failed";
            return false;
        }
    }

    private bool TryMouseWheelScrollOnElement(
        AutomationElement element,
        string direction,
        int amount,
        CancellationToken cancellationToken,
        out string strategy)
    {
        strategy = "mouse-wheel-element";

        try
        {
            var rect = element.BoundingRectangle;

            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            {
                strategy = "mouse-wheel-empty-rectangle";
                return false;
            }

            var point = new Point(
                (int)Math.Round(rect.Left + rect.Width / 2.0),
                (int)Math.Round(rect.Top + rect.Height / 2.0));

            try
            {
                element.Focus();
            }
            catch
            {
                // Some controls cannot focus. Mouse wheel may still work.
            }

            return TryMouseWheelScrollAtPoint(point, direction, amount, cancellationToken, out strategy);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Mouse wheel scroll on element failed. element={Element}, direction={Direction}, amount={Amount}",
                SafeElementName(element),
                SanitizeValue(direction),
                amount);

            strategy = "mouse-wheel-element-failed";
            return false;
        }
    }

    private bool TryMouseWheelScrollAtPoint(
        Point point,
        string direction,
        int amount,
        CancellationToken cancellationToken,
        out string strategy)
    {
        strategy = "mouse-wheel-coordinates";

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            SetCursorPos(point.X, point.Y);
            Thread.Sleep(CursorPositionStabilityDelayMs);

            var rawDelta   = ToWheelDelta(direction, amount);
            var horizontal = direction is "left" or "right";

            cancellationToken.ThrowIfCancellationRequested();
            if (!SendMouseWheelRawDelta(rawDelta, horizontal))
            {
                strategy = "mouse-wheel-coordinates-failed";
                return false;
            }

            _logger.LogInformation(
                "Mouse wheel scroll sent. point=({X},{Y}), direction={Direction}, amount={Amount}, rawDelta={RawDelta}",
                point.X,
                point.Y,
                SanitizeValue(direction),
                amount,
                rawDelta);

            return true;
        }
        catch (OperationCanceledException)
        {
            strategy = "mouse-wheel-cancelled";
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Mouse wheel scroll failed. point=({X},{Y}), direction={Direction}, amount={Amount}",
                point.X,
                point.Y,
                SanitizeValue(direction),
                amount);

            strategy = "mouse-wheel-coordinates-failed";
            return false;
        }
    }

    private static int ToWheelDelta(string direction, int amount)
    {
        const int WheelDeltaUnit = 120;

        return direction switch
        {
            "up"    =>  WheelDeltaUnit * amount,
            "down"  => -WheelDeltaUnit * amount,
            "left"  => -WheelDeltaUnit * amount,
            "right" =>  WheelDeltaUnit * amount,
            _       => -WheelDeltaUnit * amount
        };
    }

    /// <summary>
    /// Returns true when the mouse action's wheel delta corresponds to a horizontal scroll
    /// (left or right direction), so that MOUSEEVENTF_HWHEEL is used instead of MOUSEEVENTF_WHEEL.
    /// </summary>
    private static bool IsHorizontalScrollDelta(string? direction, int wheelDelta)
    {
        if (string.IsNullOrWhiteSpace(direction))
            return false;

        var normalized = direction.Trim().ToLowerInvariant();
        return normalized is "left" or "right";
    }

    /// <summary>
    /// Sends a raw wheel delta value via SendInput.
    /// Use <paramref name="horizontal"/> = true for left/right (MOUSEEVENTF_HWHEEL).
    /// </summary>
    private bool SendMouseWheelRawDelta(int rawDelta, bool horizontal = false)
    {
        try
        {
            if (rawDelta == 0)
                return true;

            var input = new INPUT
            {
                type = INPUT_MOUSE,
                U    = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx          = 0,
                        dy          = 0,
                        // Negative deltas are represented as two's-complement unsigned values.
                        mouseData   = unchecked((uint)rawDelta),
                        dwFlags     = horizontal ? MOUSEEVENTF_HWHEEL : MOUSEEVENTF_WHEEL,
                        time        = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            var sent = SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
            return sent == 1;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SendMouseWheelRawDelta failed. rawDelta={RawDelta}, horizontal={Horizontal}", rawDelta, horizontal);
            return false;
        }
    }

    private string? CaptureScrollState(AutomationElement element)
    {
        try
        {
            var sp = element.Patterns.Scroll.PatternOrDefault;

            if (sp != null)
            {
                return string.Join("|",
                    sp.HorizontalScrollPercent,
                    sp.VerticalScrollPercent,
                    sp.HorizontallyScrollable,
                    sp.VerticallyScrollable);
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            return element.BoundingRectangle.ToString();
        }
        catch
        {
            return null;
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

    internal static string SafeElementName(AutomationElement element)
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

    internal static string SafeElementAutomationId(AutomationElement element)
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

    internal static string SafeElementClassName(AutomationElement element)
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

    internal static int? SafeProcessId(AutomationElement element)
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

    internal static bool? SafeIsOffscreen(AutomationElement element)
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

    internal static bool? SafeIsEnabled(AutomationElement element)
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

    internal static string SafeElementControlType(AutomationElement element)
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

    internal static object? SafeBoundingRectangleObject(AutomationElement element)
    {
        try
        {
            var rect = element.BoundingRectangle;
            return new
            {
                left    = rect.Left,
                top     = rect.Top,
                right   = rect.Right,
                bottom  = rect.Bottom,
                width   = rect.Width,
                height  = rect.Height,
                isEmpty = rect.IsEmpty
            };
        }
        catch
        {
            return null;
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

    /// <summary>
    /// Returns a sanitized, human-readable description of a <see cref="UiLocator"/>
    /// safe for use in log messages (control characters stripped to prevent log injection).
    /// </summary>
    private static string SafeDescribeLocator(UiLocator? locator)
    {
        if (locator == null)
            return "(none)";

        var parts = new System.Text.StringBuilder();

        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            parts.Append("automationId=").Append(SanitizeValue(locator.AutomationId)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.ControlType))
            parts.Append("controlType=").Append(SanitizeValue(locator.ControlType)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.Name))
            parts.Append("name=").Append(SanitizeValue(locator.Name)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.ClassName))
            parts.Append("className=").Append(SanitizeValue(locator.ClassName)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.XPath))
            parts.Append("xpath=").Append(SanitizeValue(locator.XPath)).Append(' ');

        if (locator.Hwnd.HasValue)
            parts.Append("hwnd=").Append(locator.Hwnd.Value).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.Value))
            parts.Append("value=").Append(SanitizeValue(locator.Value)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.Text))
            parts.Append("text=").Append(SanitizeValue(locator.Text)).Append(' ');

        if (!string.IsNullOrWhiteSpace(locator.MatchMode))
            parts.Append("matchMode=").Append(SanitizeValue(locator.MatchMode)).Append(' ');

        return parts.Length > 0 ? parts.ToString().TrimEnd() : "(empty)";
    }

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
    private const uint MOUSEEVENTF_LEFTDOWN   = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP     = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN  = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP    = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP   = 0x0040;
    private const uint MOUSEEVENTF_MOVE       = 0x0001;
    private const uint MOUSEEVENTF_WHEEL      = 0x0800;
    private const uint MOUSEEVENTF_HWHEEL     = 0x1000;
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

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr hWnd);

    // =========================================================================
    // Win32 window enumeration P/Invoke
    // =========================================================================

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetShellWindow();

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    private const int SW_RESTORE = 9;

    private object DumpTree(UiRequest request)
    {
        var session = RequireSession();
        var root = _newResolver.ResolveSearchRoot(request);
        
        var depthLimit = request.Depth ?? 20;
        var limit = request.Limit ?? 1000;
        var includeOffscreen = request.IncludeOffscreen != false; // default true for dumptree

        var itemsList = new List<object>();
        var flatList = new List<(AutomationElement Element, int Depth, int CtrlIndex)>();

        void Traverse(AutomationElement current, int currentDepth, ref int ctrlIndex)
        {
            if (currentDepth > depthLimit) return;
            if (flatList.Count >= limit) return;

            flatList.Add((current, currentDepth, ctrlIndex++));

            try
            {
                var children = current.FindAllChildren();
                foreach (var child in children)
                {
                    if (!includeOffscreen && SafeIsOffscreen(child) == true)
                        continue;
                    Traverse(child, currentDepth + 1, ref ctrlIndex);
                }
            }
            catch { }
        }

        int startCtrlIndex = 0;
        Traverse(root, 0, ref startCtrlIndex);

        var baseCounts = new Dictionary<string, int>();
        var identifierMap = new Dictionary<AutomationElement, string>();
        var foundIndexMap = new Dictionary<AutomationElement, int>();

        foreach (var (el, d, ctrlIdx) in flatList)
        {
            var baseId = GetBaseIdentifier(el);
            if (!baseCounts.ContainsKey(baseId))
            {
                baseCounts[baseId] = 0;
            }
            var idx = baseCounts[baseId];
            baseCounts[baseId]++;

            foundIndexMap[el] = idx;

            if (idx == 0)
            {
                identifierMap[el] = baseId;
            }
            else
            {
                // When idx is 1 (the first duplicate), we want it numbered 0 to produce the requested sequence: Edit, Edit0, Edit1, Edit2
                identifierMap[el] = $"{baseId}{idx - 1}";
            }
        }

        var includeIdentifiers = request.IncludeIdentifiers != false;

        foreach (var (el, d, ctrlIdx) in flatList)
        {
            var id = identifierMap[el];
            var fIdx = foundIndexMap[el];
            var baseId = GetBaseIdentifier(el);
            var controlTypeStr = SafeElementControlType(el);
            var name = SafeElementName(el);
            var aid = SafeElementAutomationId(el);
            var className = SafeElementClassName(el);
            var rect = el.BoundingRectangle;

            var betterSuggestions = new List<object>();
            if (!string.IsNullOrEmpty(aid))
            {
                betterSuggestions.Add(new { automationId = aid, controlType = controlTypeStr });
            }
            if (!string.IsNullOrEmpty(name))
            {
                betterSuggestions.Add(new { name = name, controlType = controlTypeStr });
            }
            if (!string.IsNullOrEmpty(className))
            {
                betterSuggestions.Add(new { className = className, controlType = controlTypeStr, foundIndex = fIdx });
            }
            betterSuggestions.Add(new { controlType = controlTypeStr, foundIndex = fIdx });
            betterSuggestions.Add(new { rectangle = new { left = (int)rect.Left, top = (int)rect.Top, width = (int)rect.Width, height = (int)rect.Height }, controlType = controlTypeStr });

            itemsList.Add(new
            {
                identifier = includeIdentifiers ? id : null,
                name = name,
                automationId = aid,
                controlType = controlTypeStr,
                className = className,
                foundIndex = fIdx,
                ctrlIndex = ctrlIdx,
                depth = d,
                rectangle = new
                {
                    left = (int)rect.Left,
                    top = (int)rect.Top,
                    right = (int)rect.Right,
                    bottom = (int)rect.Bottom,
                    width = (int)rect.Width,
                    height = (int)rect.Height
                },
                patterns = new
                {
                    value = el.Patterns.Value.IsSupported,
                    text = el.Patterns.Text.IsSupported,
                    invoke = el.Patterns.Invoke.IsSupported,
                    selectionItem = el.Patterns.SelectionItem.IsSupported,
                    toggle = el.Patterns.Toggle.IsSupported
                },
                locatorSuggestion = new
                {
                    controlType = controlTypeStr,
                    foundIndex = fIdx
                },
                betterLocatorSuggestions = betterSuggestions
            });
        }

        return new
        {
            operation = request.Operation,
            searchRoot = request.SearchRoot ?? "currentWindow",
            treeView = request.TreeView ?? "control",
            depth = depthLimit,
            count = flatList.Count,
            items = itemsList
        };
    }

    private string GetBaseIdentifier(AutomationElement element)
    {
        var name = SafeElementName(element);
        if (!string.IsNullOrWhiteSpace(name))
            return CleanIdentifier(name);

        var aid = SafeElementAutomationId(element);
        if (!string.IsNullOrWhiteSpace(aid))
            return CleanIdentifier(aid);

        var ct = SafeElementControlType(element);
        if (!string.IsNullOrWhiteSpace(ct))
            return CleanIdentifier(ct);

        var cn = SafeElementClassName(element);
        if (!string.IsNullOrWhiteSpace(cn))
            return CleanIdentifier(cn);

        return "Control";
    }

    private string CleanIdentifier(string raw)
    {
        var cleaned = new string(raw.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());
        return string.IsNullOrEmpty(cleaned) ? "Control" : cleaned;
    }

    private string GetSearchRootName(UiRequest request)
    {
        return request.SearchRoot ?? (request.UseDesktopRoot == true ? "desktop" : (request.UseActiveWindowRoot == true ? "foreground" : "current"));
    }

    private object FindElement(UiRequest request)
    {
        var locator = request.Locator ?? new UiLocator();
        var matchResult = _newResolver.ResolveOne(locator, request, "findelement");
        var firstCand = matchResult.Candidates.FirstOrDefault();
        var snapshot = DesktopAutomationDriver.Services.Resolution.ElementResolver.CreateSnapshot(matchResult.Element);
        return new
        {
            found = true,
            index = 0,
            score = firstCand?.Score ?? 100,
            strategy = matchResult.Strategy,
            reason = firstCand?.MatchReasons.FirstOrDefault() ?? "Matched",
            element = new
            {
                name = snapshot.Name,
                automationId = snapshot.AutomationId,
                controlType = snapshot.ControlType,
                className = snapshot.ClassName,
                rectangle = snapshot.Rectangle,
                isEnabled = snapshot.IsEnabled,
                isOffscreen = snapshot.IsOffscreen
            }
        };
    }

    private object FindElements(UiRequest request)
    {
        var candidates = _newResolver.ResolveAll(request, "findall");
        var matchesList = candidates.Select((c, idx) => new
        {
            index = idx,
            score = c.Score,
            strategy = "unified-resolver-findall",
            reason = c.MatchReasons.FirstOrDefault() ?? "Matched",
            element = new
            {
                name = c.Snapshot.Name,
                automationId = c.Snapshot.AutomationId,
                controlType = c.Snapshot.ControlType,
                className = c.Snapshot.ClassName,
                rectangle = c.Snapshot.Rectangle,
                isEnabled = c.Snapshot.IsEnabled,
                isOffscreen = c.Snapshot.IsOffscreen
            }
        }).ToList();

        return new
        {
            count = candidates.Count,
            matches = matchesList
        };
    }

    private object InspectElement(UiRequest request)
    {
        var locator = request.Locator ?? new UiLocator();
        var matchResult = _newResolver.ResolveOne(locator, request, "inspectelement");
        var snapshot = DesktopAutomationDriver.Services.Resolution.ElementResolver.CreateSnapshot(matchResult.Element);
        return new
        {
            found = true,
            snapshot = snapshot
        };
    }

    private object ResolveDebug(UiRequest request)
    {
        return FindElements(request);
    }

    private object PrintControlIdentifiers(UiRequest request)
    {
        var session = TryGetSessionOrNull();
        if (session == null)
            throw new InvalidOperationException("No active automation session.");

        AutomationElement rootEl;
        if (request.Locator != null)
        {
            var locator = request.Locator;
            var matchResult = _newResolver.ResolveOne(locator, request, "printcontrolidentifiers");
            rootEl = matchResult.Element;
        }
        else
        {
            rootEl = GetWindowRoot(session, false);
        }

        var tree = new List<object>();
        WalkTreeForIdentifiers(rootEl, 0, request.Depth ?? 5, tree);

        return new
        {
            root = new
            {
                name = SafeElementName(rootEl),
                controlType = SafeElementControlType(rootEl),
                automationId = SafeElementAutomationId(rootEl),
                className = SafeElementClassName(rootEl),
                rectangle = SafeBoundingRectangleObject(rootEl)
            },
            tree = tree
        };
    }

    private void WalkTreeForIdentifiers(AutomationElement element, int depth, int maxDepth, List<object> tree)
    {
        if (depth > maxDepth) return;

        tree.Add(new
        {
            depth = depth,
            name = SafeElementName(element) ?? string.Empty,
            controlType = SafeElementControlType(element) ?? string.Empty,
            automationId = SafeElementAutomationId(element) ?? string.Empty,
            className = SafeElementClassName(element) ?? string.Empty,
            rectangle = SafeBoundingRectangleObject(element)
        });

        try
        {
            var children = element.FindAllChildren();
            foreach (var child in children)
            {
                WalkTreeForIdentifiers(child, depth + 1, maxDepth, tree);
            }
        }
        catch { }
    }

    private object? ListHeaderDropdownItems(UiRequest req)
    {
        var header = ResolveElementForOperation(
            req,
            purpose: "listheaderdropdownitems",
            action: true,
            allowOffscreen: false,
            requireClickable: true);
        var region = GridHeaderDropdownHelper.ParseRegion(req.Value ?? req.ClickRegion);

        var container = OpenHeaderDropdownAndFindList(header, region);
        if (container == null)
        {
            throw new InvalidOperationException("Failed to open header dropdown or find dropdown container.");
        }

        var items = GetDropdownSelectableItems(container)
            .Select(i => SafeElementName(i))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
        return new
        {
            opened = true,
            container = new
            {
                controlType = SafeElementControlType(container),
                className = SafeElementClassName(container),
                rectangle = SafeBoundingRectangleObject(container)
            },
            items
        };
    }
}
