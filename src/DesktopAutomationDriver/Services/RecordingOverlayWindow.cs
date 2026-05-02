using System.Runtime.InteropServices;
using DesktopAutomationDriver.Models.Recording;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

// Aliases to resolve ambiguity with identically-named FlaUI types
using WinLabel = System.Windows.Forms.Label;
using WinApplication = System.Windows.Forms.Application;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// A borderless, always-on-top, semi-transparent status bar displayed at the top
/// of the primary screen during a recording session.
///
/// This form installs and owns the low-level Win32 keyboard / mouse hooks.
/// All hook callbacks fire on this form's STA message-pump thread.
/// </summary>
public sealed class RecordingOverlayWindow : Form
{
    // ── Win32 hook constants ────────────────────────────────────────────────
    private const int WH_KEYBOARD_LL = 13;
    private const int WH_MOUSE_LL = 14;

    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;
    private const int WM_LBUTTONDBLCLK = 0x0203;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_RBUTTONUP   = 0x0205;
    private const int WM_MOUSEMOVE = 0x0200;

    private const byte VK_CONTROL = 0x11;
    private const byte VK_P = 0x50;
    private const byte VK_A = 0x41;
    private const byte VK_S = 0x53;
    private const byte VK_W = 0x57;

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public System.Drawing.Point pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    // ── extended window-style P/Invokes (click-through) ────────────────────
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WS_EX_LAYERED = 0x00080000;
    private const int WS_EX_TOPMOST = 0x00000008;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_NOACTIVATE = 0x08000000;

    /// <summary>Maximum number of child elements shown in the "Children ▶" submenu.</summary>
    private const int MaxChildrenToDisplay = 30;
    private const int MaxTableAncestorDepth = 5;
    private const int MaxWindowSearchDepth = 5;
    private const int MaxMenuAncestorDepth = 10;

    /// <summary>
    /// Minimum pixel distance the mouse must travel while the left button is held before a
    /// mouse-down + mouse-up sequence is classified as a drag rather than a click.
    /// 8 px is large enough to ignore normal hand tremor and sub-pixel high-DPI jitter
    /// while still being small enough to catch deliberate short drags.
    /// </summary>
    private const int DragThresholdPixels = 8;

    /// <summary>Pre-computed squared threshold to avoid a multiplication on every WM_LBUTTONUP.</summary>
    private const int DragThresholdPixelsSq = DragThresholdPixels * DragThresholdPixels;

    /// <summary>
    /// Delay in milliseconds between intercepting a right-click and showing the assistive
    /// context menu.  A one-tick timer is used instead of BeginInvoke so that any
    /// window-activation or native-menu-dismissal messages triggered by the right-click
    /// are fully processed before the menu is built, ensuring the correct UIA element is
    /// identified on the first right-click over MenuBar / MenuItem elements.
    /// </summary>
    private const int ContextMenuDelayMs = 1;

    /// <summary>
    /// Delay in milliseconds inserted after <c>SetForegroundWindow</c> to let the OS
    /// finish window activation before firing a physical mouse click.  Without this
    /// pause, the click can land on whatever window still holds the input focus
    /// (typically the IDE that launched the driver) rather than the target window.
    /// </summary>
    private const int WindowActivationDelayMs = 100;

    /// <summary>
    /// Delay in milliseconds inserted between each step when re-navigating a menu
    /// hierarchy for a popup sub-MenuItem click.  Allows time for each intermediate
    /// dropdown to materialise in the UIA tree before the next item is searched.
    /// </summary>
    private const int MenuNavigationDelayMs = 300;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    // ── foreground-window P/Invokes ─────────────────────────────────────────
    /// <summary>
    /// Walks up the HWND tree to the root (top-level) window.
    /// GA_ROOT (2) returns the root window that doesn't have a parent.
    /// </summary>
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    private const uint GA_ROOT = 2;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    // ── fields ──────────────────────────────────────────────────────────────
    private readonly IRecordingService _service;
    private readonly ILogger _logger;

    private IntPtr _keyboardHook = IntPtr.Zero;
    private IntPtr _mouseHook = IntPtr.Zero;

    // Keep hard references to hook delegates so GC never collects them
    private HookProc? _keyboardProc;
    private HookProc? _mouseProc;

    private UIA3Automation? _automation;
    private System.Windows.Forms.Timer? _cursorTimer;

    private WinLabel _statusLabel = null!;

    // Guard against showing multiple context menus simultaneously (e.g. rapid right-clicks).
    private bool _menuOpen;

    // One-shot 10 s fallback timer started whenever _menuOpen is set to true.
    // Disposed/stopped in ReapplyClickThroughStyle() (called from every menu.Closed handler)
    // so it does not outlive the menu it guards.
    private System.Windows.Forms.Timer? _menuSafetyTimer;

    // Set to true after we suppress WM_RBUTTONDOWN so we can also suppress the
    // matching WM_RBUTTONUP (preventing the target app from receiving WM_CONTEXTMENU
    // which would open the app's own context menu and immediately close ours).
    private bool _suppressNextRButtonUp;

    // Set to true by the Assistive-mode "Right Click" action item just before it calls
    // Mouse.RightClick().  The next WM_RBUTTONDOWN from the hook is then passed through
    // rather than intercepted so that the simulated click reaches the target window.
    private bool _suppressNextRButtonDown;

    // ── Passive mode drag detection state ────────────────────────────────────
    // Set on WM_LBUTTONDOWN; cleared on WM_LBUTTONUP or WM_LBUTTONDBLCLK.
    private bool _leftButtonDown;
    private System.Drawing.Point _mouseDownPoint;
    private ElementInfo? _mouseDownEagerInfo;
    // Squared maximum distance (avoids a sqrt per MOUSEMOVE event).
    private int _maxDragDistanceSq;

    // ── Assistive-mode drag-and-drop two-step selection state ─────────────────
    // When true the user has picked a drag source and the next right-click selects the target.
    private bool _awaitingDragTarget;
    private ElementInfo? _dragSourceInfo;

    // ── Assistive-mode popup detection cache ──────────────────────────────────
    private AutomationElement? _lastDetectedPopupWindow;
    private IntPtr _lastDetectedPopupHwnd = IntPtr.Zero;
    private DateTime _lastPopupDetectionUtc = DateTime.MinValue;

    // ── Drag-tracking helpers ────────────────────────────────────────────────
    /// <summary>
    /// Initialises passive-mode drag-detection state for a new left-button-down event.
    /// </summary>
    private void BeginDragTracking(System.Drawing.Point pt, ElementInfo? eagerInfo)
    {
        _mouseDownPoint = pt;
        _mouseDownEagerInfo = eagerInfo;
        _leftButtonDown = true;
        _maxDragDistanceSq = 0;
    }

    // ── constructor ─────────────────────────────────────────────────────────
    public RecordingOverlayWindow(IRecordingService service, ILogger logger)
    {
        _service = service;
        _logger = logger;
        BuildUi();
    }

    // ── UI setup ─────────────────────────────────────────────────────────────
    private void BuildUi()
    {
        var screen = Screen.PrimaryScreen ?? Screen.AllScreens[0];

        const int OverlayWidth  = 420;
        const int OverlayHeight = 52;

        FormBorderStyle = FormBorderStyle.None;
        TopMost = true;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        Opacity = 0.88;
        BackColor = Color.FromArgb(20, 20, 20);
        ForeColor = Color.White;

        // Small widget pinned to the top-right corner of the primary screen
        Bounds = new Rectangle(
            screen.Bounds.Right - OverlayWidth,
            screen.Bounds.Top,
            OverlayWidth,
            OverlayHeight);

        _statusLabel = new WinLabel
        {
            Dock = DockStyle.Fill,
            ForeColor = Color.White,
            BackColor = Color.Transparent,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Text = "● REC  P=Passive  A=Assistive  S=Stop"
        };

        Controls.Add(_statusLabel);
    }

    // ── Form lifecycle ───────────────────────────────────────────────────────
    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);

        // Make the status bar click-through so it doesn't block underlying windows
        var style = GetWindowLong(Handle, GWL_EXSTYLE);
        SetWindowLong(Handle, GWL_EXSTYLE, style | WS_EX_TRANSPARENT | WS_EX_LAYERED);

        // Create the UIA3 automation backend on this STA thread
        try
        {
            _automation = new UIA3Automation();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not initialize UIA3 automation for recording");
        }

        // Install global low-level hooks (must be installed on a thread with a message pump)
        _keyboardProc = KeyboardHookCallback;
        _mouseProc = MouseHookCallback;

        var hMod = GetModuleHandle(null);
        _keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, hMod, 0);
        _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, hMod, 0);

        if (_keyboardHook == IntPtr.Zero || _mouseHook == IntPtr.Zero)
            _logger.LogWarning("Failed to install one or more low-level hooks (keyboard={Kb}, mouse={Ms})",
                _keyboardHook, _mouseHook);

        // Timer to refresh the "cursor on: …" label in Assistive mode
        _cursorTimer = new System.Windows.Forms.Timer { Interval = 150 };
        _cursorTimer.Tick += OnCursorTimerTick;
        _cursorTimer.Start();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _cursorTimer?.Stop();
        _cursorTimer?.Dispose();

        _menuSafetyTimer?.Stop();
        _menuSafetyTimer?.Dispose();
        _menuSafetyTimer = null;

        if (_keyboardHook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_keyboardHook);
            _keyboardHook = IntPtr.Zero;
        }
        if (_mouseHook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_mouseHook);
            _mouseHook = IntPtr.Zero;
        }

        _automation?.Dispose();
        _automation = null;

        _service.OnOverlayClosed();
        base.OnFormClosed(e);
    }

    // ── Cursor timer (Assistive mode status update) ──────────────────────────
    private void OnCursorTimerTick(object? sender, EventArgs e)
    {
        if (_service.CurrentMode != RecordingMode.Assistive) return;

        // Proactively scan for a child popup window so the overlay updates without
        // requiring the user to move the mouse over the popup first.
        var popup = DetectActivePopupWindow();
        if (popup != null)
        {
            _lastDetectedPopupWindow = popup;
            _lastDetectedPopupHwnd = AssistivePopupResolver.SafeWindowHandle(popup);
            _lastPopupDetectionUtc = DateTime.UtcNow;

            var popupName = popup.Name ?? popup.AutomationId ?? "popup";
            var popupText = $"● [Popup] {popupName}  Right-click = Popup Actions";

            if (_statusLabel.Text != popupText)
            {
                _statusLabel.Text = popupText;
                _logger.LogInformation(
                    "Overlay popup status updated: {Name}, hwnd=0x{Hwnd:X}",
                    popupName,
                    _lastDetectedPopupHwnd.ToInt64());
            }

            return;
        }

        var pt = Cursor.Position;
        var info = _service.GetElementAtPoint(pt);
        var name = info?.Name ?? info?.AutomationId ?? info?.ControlType ?? "unknown";
        var ct = info?.ControlType ?? string.Empty;

        // Truncate the element name so the label always fits in the 420 px corner widget
        // (≈ 9.5 pt Segoe UI → ~18 px/char → 22 chars keeps the label under ~400 px).
        const int MaxNameLen = 22;
        var displayName = name.Length > MaxNameLen ? name[..MaxNameLen] + "…" : name;

        var text = $"● [{ct}] {displayName}  Right-click  │  Ctrl+RC = Window";
        if (_statusLabel.Text != text)
            _statusLabel.Text = text;
    }

    // ── Keyboard hook ────────────────────────────────────────────────────────
    private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
        {
            var kbStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

            if (ctrlDown)
            {
                if (kbStruct.vkCode == VK_P)
                {
                    BeginInvoke(new Action(() => ActivateMode(RecordingMode.Passive)));
                    return (IntPtr)1; // suppress
                }
                if (kbStruct.vkCode == VK_A)
                {
                    BeginInvoke(new Action(() => ActivateMode(RecordingMode.Assistive)));
                    return (IntPtr)1; // suppress
                }
                if (kbStruct.vkCode == VK_S)
                {
                    BeginInvoke(new Action(RequestStop));
                    return (IntPtr)1; // suppress
                }
                if (kbStruct.vkCode == VK_W &&
                    _service.CurrentMode == RecordingMode.Assistive)
                {
                    BeginInvoke(new Action(() =>
                    {
                        var pt = Cursor.Position;
                        _statusLabel.Text = $"CTRL+W popup menu at {pt.X},{pt.Y}";
                        try
                        {
                            ShowWindowContextMenu(pt);
                        }
                        catch (Exception ex)
                        {
                            _statusLabel.Text = "CTRL+W popup menu failed: " + ex.Message;
                            _logger.LogError(ex, "CTRL+W popup menu failed");
                        }
                    }));
                    return (IntPtr)1; // suppress
                }
            }
        }

        return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private void ActivateMode(RecordingMode mode)
    {
        // Reset drag-related state so stale captures from the previous mode are discarded.
        _leftButtonDown = false;
        _awaitingDragTarget = false;
        _dragSourceInfo = null;

        _service.SetMode(mode);

        // When switching to Assistive mode, bring the session's application window to
        // the foreground so that UIA FromPoint returns elements from the target app
        // (not from IntelliJ or another IDE that may be behind/overlapping it).
        if (mode == RecordingMode.Assistive)
            _service.BringApplicationWindowToFront();

        var label = mode == RecordingMode.Passive
            ? "● PASSIVE  Recording clicks & keys  S=Stop"
            : "● ASSISTIVE  Right-click element  │  Ctrl+Right-click for window actions  S=Stop";
        _statusLabel.Text = label;
    }

    private void RequestStop()
    {
        Close(); // OnFormClosed → _service.OnOverlayClosed() handles export
    }

    // ── Mouse hook ───────────────────────────────────────────────────────────
    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode < 0)
            return CallNextHookEx(_mouseHook, nCode, wParam, lParam);

        var ms = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

        if (_service.CurrentMode == RecordingMode.Passive)
        {
            if (wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                var pt = ms.pt;
                // Capture element info eagerly in the hook callback, before the hook
                // returns and the application processes the click (which may close a
                // dialog before BeginInvoke executes, e.g. a Save As "Save"/"Cancel" button).
                ElementInfo? eagerInfo = null;
                try { eagerInfo = _service.GetElementAtPoint(pt); } catch { /* best effort */ }

                // Store state for drag detection. The click is recorded now; if the
                // mouse moves beyond DragThresholdPixels before button-up, the click
                // action will be replaced by a DragAndDrop action.
                BeginDragTracking(pt, eagerInfo);

                BeginInvoke(new Action(() => RecordPassiveClick(pt, eagerInfo, ActionType.Click)));
            }
            else if (wParam == (IntPtr)WM_LBUTTONDBLCLK)
            {
                // Cancel drag tracking — the button-up for a double-click should not
                // trigger drag detection since WM_LBUTTONDBLCLK already recorded the action.
                _leftButtonDown = false;
                var pt = ms.pt;
                ElementInfo? eagerInfo = null;
                try { eagerInfo = _service.GetElementAtPoint(pt); } catch { /* best effort */ }
                BeginInvoke(new Action(() => RecordPassiveClick(pt, eagerInfo, ActionType.DoubleClick)));
            }
            else if (wParam == (IntPtr)WM_MOUSEMOVE && _leftButtonDown)
            {
                // Accumulate the maximum squared distance from the mouse-down point.
                // Using squared distance avoids a sqrt on every mouse-move event.
                var dx = ms.pt.X - _mouseDownPoint.X;
                var dy = ms.pt.Y - _mouseDownPoint.Y;
                var distSq = dx * dx + dy * dy;
                if (distSq > _maxDragDistanceSq)
                    _maxDragDistanceSq = distSq;
            }
            else if (wParam == (IntPtr)WM_LBUTTONUP && _leftButtonDown)
            {
                _leftButtonDown = false;
                if (_maxDragDistanceSq >= DragThresholdPixelsSq)
                {
                    // Significant movement: upgrade the earlier Click to a DragAndDrop.
                    var upPt = ms.pt;
                    var sourceInfo = _mouseDownEagerInfo;
                    ElementInfo? targetInfo = null;
                    try { targetInfo = _service.GetElementAtPoint(upPt); } catch { /* best effort */ }
                    BeginInvoke(new Action(() => RecordPassiveDrag(sourceInfo, targetInfo)));
                }
                // Small movement: the Click already recorded on mouse-down is correct; nothing to do.
            }
            else if (wParam == (IntPtr)WM_RBUTTONDOWN)
            {
                var pt = ms.pt;
                // Capture element info eagerly before the hook returns (same reason as left-click).
                ElementInfo? eagerInfo = null;
                try { eagerInfo = _service.GetElementAtPoint(pt); } catch { /* best effort */ }
                BeginInvoke(new Action(() => RecordPassiveRightClick(pt, eagerInfo)));
                // Intentionally NOT suppressed — the native right-click is passed through so
                // the application can display its own context menu if it has one.
            }
        }
        else if (_service.CurrentMode == RecordingMode.Assistive)
        {
            if (wParam == (IntPtr)WM_RBUTTONDOWN)
            {
                var pt = ms.pt;

                // Ctrl+Right Click OR normal right-click when a popup is active →
                // show the window-level context menu instead of the normal element context menu.
                bool ctrlHeld = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool popupActive = IsPopupCurrentlyActive();
                if (ctrlHeld || popupActive)
                {
                    BeginInvoke(new Action(() =>
                    {
                        if (popupActive)
                            _statusLabel.Text = $"Popup detected — showing popup actions at {pt.X},{pt.Y}";
                        else
                            _statusLabel.Text = $"CTRL+RC captured at {pt.X},{pt.Y}";
                        try
                        {
                            ShowWindowContextMenu(pt);
                        }
                        catch (Exception ex)
                        {
                            _statusLabel.Text = "Popup/window menu failed: " + ex.Message;
                            _logger.LogError(ex, "Popup/window context menu failed");
                        }
                    }));
                    _suppressNextRButtonUp = true;
                    return (IntPtr)1; // suppress the native right-click
                }

                // A simulated right-click fired by the "Right Click" action item sets
                // _suppressNextRButtonDown so the hook does not intercept it.
                if (_suppressNextRButtonDown)
                {
                    _suppressNextRButtonDown = false;
                    return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
                }

                if (_awaitingDragTarget)
                {
                    // Second step of assistive drag: the user right-clicked the drop target.
                    _awaitingDragTarget = false;
                    var sourceInfo = _dragSourceInfo;
                    _dragSourceInfo = null;

                    // Use a one-tick timer for the same reason as normal right-click handling
                    // (let native menu/activation messages settle before accessing the UIA tree).
                    var timer = new System.Windows.Forms.Timer { Interval = ContextMenuDelayMs };
                    timer.Tick += (_, _) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        RecordAssistiveDragTarget(pt, sourceInfo);
                    };
                    timer.Start();
                }
                else
                {
                    // Normal: show the assistive context menu.
                    // Use a one-tick timer instead of BeginInvoke so that any native
                    // menu dismissal or window-activation messages triggered by the
                    // right-click are fully processed before we build and show the
                    // assistive context menu.  This ensures the correct element is
                    // identified on the first right-click over MenuBar / MenuItem
                    // elements where the native menu may still be transitioning.
                    var timer = new System.Windows.Forms.Timer { Interval = ContextMenuDelayMs };
                    timer.Tick += (_, _) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        ShowAssistiveContextMenu(pt);
                    };
                    timer.Start();
                }

                _suppressNextRButtonUp = true;
                return (IntPtr)1; // suppress the native right-click
            }
            if (wParam == (IntPtr)WM_RBUTTONUP && _suppressNextRButtonUp)
            {
                _suppressNextRButtonUp = false;
                return (IntPtr)1; // suppress the matching RBUTTONUP so the target app doesn't open its own context menu
            }
        }

        return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    // ── Passive recording helpers ────────────────────────────────────────────
    private void RecordPassiveClick(System.Drawing.Point pt, ElementInfo? eagerInfo, ActionType actionType)
    {
        // Prefer the eagerly captured info (grabbed in the hook callback while the
        // element was still on screen, before the application processed the click).
        // Fall back to a fresh point-lookup in case the eager capture failed.
        var info = eagerInfo ?? _service.GetElementAtPoint(pt);

        // When a ListItem is left-clicked, check whether it lives inside an editable
        // ComboBox that already contains typed text — if so, record as TypeAndSelect
        // rather than a plain click so the automation script can replay the filter+pick
        // interaction correctly.
        if (actionType == ActionType.Click
            && info?.ControlType == "ListItem"
            && _automation != null)
        {
            try
            {
                var element = _automation.FromPoint(pt);
                if (element != null)
                {
                    var (comboInfo, typedText) = FindEditableComboBoxAncestor(element);
                    if (comboInfo != null && !string.IsNullOrEmpty(typedText))
                    {
                        _service.AddAction(new RecordedAction
                        {
                            ActionType = ActionType.TypeAndSelect,
                            Mode = RecordingMode.Passive,
                            Element = comboInfo,
                            Value = typedText,
                            Description = $"Type and select '{info.Name ?? typedText}' in {ElementInfo.GetLabel(comboInfo)}"
                        });
                        return;
                    }
                }
            }
            catch { /* best effort */ }
        }

        var actionLabel = actionType == ActionType.DoubleClick ? "Double Click" : "Click";
        _service.AddAction(new RecordedAction
        {
            ActionType = actionType,
            Mode = RecordingMode.Passive,
            Element = info,
            Description = BuildDescription(actionLabel, info)
        });
    }

    /// <summary>
    /// Called (via BeginInvoke) when passive-mode drag detection determines that the left
    /// button was held and the cursor moved at least <see cref="DragThresholdPixels"/> pixels.
    /// Replaces the Click that was optimistically recorded on mouse-down with a DragAndDrop.
    /// </summary>
    private void RecordPassiveDrag(ElementInfo? sourceInfo, ElementInfo? targetInfo)
    {
        var sourceLabel = ElementInfo.GetLabel(sourceInfo);
        var targetLabel = ElementInfo.GetLabel(targetInfo);
        _service.ReplaceLastAction(new RecordedAction
        {
            ActionType = ActionType.DragAndDrop,
            Element = sourceInfo,
            TargetElement = targetInfo,
            Description = $"Drag from {sourceLabel} to {targetLabel}"
        });
    }

    /// <summary>
    /// Called (via BeginInvoke) when a right-button-down event is observed in Passive mode.
    /// Records the action without suppressing the event so the application can still display
    /// its own context menu if it has one.
    /// </summary>
    private void RecordPassiveRightClick(System.Drawing.Point pt, ElementInfo? eagerInfo)
    {
        var info = eagerInfo ?? _service.GetElementAtPoint(pt);
        _service.AddAction(new RecordedAction
        {
            ActionType = ActionType.RightClick,
            Mode = RecordingMode.Passive,
            Element = info,
            Description = BuildDescription("Right Click", info)
        });
    }

    /// <summary>
    /// Called (via a one-tick timer) when the user right-clicks the drop target while the
    /// assistive-mode drag-source selection is pending (<see cref="_awaitingDragTarget"/> was true).
    /// Identifies the element at <paramref name="pt"/> and records the DragAndDrop action.
    /// </summary>
    private void RecordAssistiveDragTarget(System.Drawing.Point pt, ElementInfo? sourceInfo)
    {
        if (_automation == null) return;

        ElementInfo? targetInfo = null;
        try
        {
            var element = _automation.FromPoint(pt);
            if (element != null)
                targetInfo = BuildElementInfo(element);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not get drop-target element at point {Pt}", pt);
        }

        var sourceLabel = ElementInfo.GetLabel(sourceInfo);
        var targetLabel = ElementInfo.GetLabel(targetInfo);
        _service.AddAction(new RecordedAction
        {
            ActionType = ActionType.DragAndDrop,
            Mode = RecordingMode.Assistive,
            Element = sourceInfo,
            TargetElement = targetInfo,
            Description = $"Drag from {sourceLabel} to {targetLabel}"
        });
        UpdateStatusAfterAction(
            $"Drag [{sourceInfo?.ControlType}] {sourceLabel} → [{targetInfo?.ControlType}] {targetLabel}");
    }

    /// <summary>
    /// Walks the UIA control tree upward from <paramref name="element"/> (up to 5 levels)
    /// looking for an editable ComboBox ancestor that currently contains non-empty text.
    /// Returns the ComboBox's <see cref="ElementInfo"/> and its current value, or
    /// <c>(null, null)</c> when no such ancestor is found.
    /// </summary>
    private (ElementInfo? info, string? text) FindEditableComboBoxAncestor(AutomationElement element)
    {
        try
        {
            var walker = _automation!.TreeWalkerFactory.GetControlViewWalker();
            var current = element;

            for (int i = 0; i < 5; i++)
            {
                AutomationElement? parent;
                try { parent = walker.GetParent(current); }
                catch { break; }

                if (parent == null) break;

                try
                {
                    if (parent.ControlType == FlaUI.Core.Definitions.ControlType.ComboBox)
                    {
                        var valuePat = parent.Patterns.Value.PatternOrDefault;
                        if (valuePat != null && !valuePat.IsReadOnly)
                        {
                            var text = valuePat.Value;
                            if (!string.IsNullOrEmpty(text))
                                return (BuildElementInfo(parent), text);
                        }
                    }
                }
                catch { /* ignore inaccessible element */ }

                current = parent;
            }
        }
        catch { /* best effort */ }

        return (null, null);
    }

    // ── Assistive mode: context menu ─────────────────────────────────────────
    private void ShowAssistiveContextMenu(System.Drawing.Point pt)
    {
        // Ignore rapid right-clicks while a menu is already visible.
        if (_menuOpen) return;
        if (_automation == null) return;

        _menuOpen = true;

        // Safety fallback: reset _menuOpen after 10 s in case the menu's Closed event
        // never fires (e.g. edge-case focus or disposal race).
        StartMenuSafetyTimer();

        try
        {
            ShowAssistiveContextMenuCore(pt);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to show assistive context menu at {Pt}", pt);
            _menuOpen = false;
        }
    }

    private void ShowAssistiveContextMenuCore(System.Drawing.Point pt)
    {
        AutomationElement? element = null;
        try { element = _automation!.FromPoint(pt); }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not get element at point {Pt} for assistive menu", pt);
        }

        // For container controls (e.g. WinForms MenuStrip/MenuBar), FromPoint may return
        // the parent instead of the specific child under the cursor. Drill down one level.
        if (element != null)
            element = DrillDownToElementAtPoint(element, pt);

        // Save the original element before any promotion so that the "Type…" and
        // "Is Editable" menu items are still available when an inner Edit control is
        // promoted to its parent ComboBox (e.g. a username/password field with history).
        var originalElement = element;

        // If FromPoint returned a structural child of a ComboBox (e.g. the inner Edit
        // text-box, the dropdown Button, or a ListItem / List from the expanded dropdown),
        // promote to the ComboBox parent so that the "Options ▶" submenu and
        // "Type and Select…" items are available.
        // A ListItem may be nested two levels deep: ComboBox → List → ListItem, so walk
        // up to 3 levels to find the ComboBox ancestor.
        if (element != null &&
            (element.ControlType == ControlType.Edit ||
             element.ControlType == ControlType.Button ||
             element.ControlType == ControlType.ListItem ||
             element.ControlType == ControlType.List))
        {
            try
            {
                var current = element;
                for (int i = 0; i < 3; i++)
                {
                    var parent = current.Parent;
                    if (parent == null) break;
                    if (parent.ControlType == ControlType.ComboBox)
                    {
                        element = parent;
                        break;
                    }
                    current = parent;
                }
            }
            catch { /* best effort */ }
        }

        var elementInfo = element != null ? BuildElementInfo(element) : null;

        // Walk up to the nearest Window ancestor so we can later scan for child popup
        // windows (used for both post-action detection and the "Popup Windows ▶" menu).
        AutomationElement? windowAncestor = null;
        if (element != null && _automation != null)
        {
            try
            {
                var cur = element;
                for (int depth = 0; depth < MaxWindowSearchDepth && cur != null; depth++)
                {
                    if (cur.ControlType == ControlType.Window)
                    {
                        windowAncestor = cur;
                        break;
                    }
                    cur = cur.Parent;
                }
            }
            catch { /* best effort */ }
        }

        // Snapshot the current child Window names so we can detect newly opened popup
        // windows once the context menu is dismissed (popup-detection notification).
        var preActionChildWindowTitles = new HashSet<string>(StringComparer.Ordinal);
        if (windowAncestor != null && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                foreach (var childWindow in windowAncestor.FindAllChildren(cf.ByControlType(ControlType.Window)))
                {
                    try { preActionChildWindowTitles.Add(childWindow.Name ?? string.Empty); }
                    catch { /* best effort */ }
                }
            }
            catch { /* best effort */ }
        }

        // Capture the element's top-level window HWND before the assistive context menu
        // is shown.  NoActivateContextMenuStrip (WS_EX_NOACTIVATE) prevents
        // activation-based dismissal of modal popup windows (IsModal = true), but some
        // popups also close on click-based dismissal (e.g. WPF Popup with
        // StaysOpen = False).  Using the pre-captured HWND in the Right Click handler
        // avoids stale-element exceptions when the element's properties are queried
        // after the popup has closed.
        IntPtr capturedHwnd = IntPtr.Zero;
        if (element != null)
        {
            try
            {
                var hwnd = element.Properties.NativeWindowHandle.Value;
                if (hwnd != IntPtr.Zero)
                {
                    var root = GetAncestor(hwnd, GA_ROOT);
                    capturedHwnd = root != IntPtr.Zero ? root : hwnd;
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not capture HWND for element at {Pt}; Right Click will fall back to element reference", pt);
            }
        }

        // For sub-MenuItems that live inside a ToolStripDropDown popup: even with
        // NoActivateContextMenuStrip preventing activation-based dismissal, some
        // WinForms popups may still collapse when focus is redirected.  The captured
        // 'element' reference then points to a destroyed UIA element, so Invoke() /
        // Click() on it fail silently.  Build a re-navigation path right now (while
        // the popup is still open) so the Click handler can re-open the menu chain
        // and locate fresh UIA elements.
        AutomationElement? popupMenuBarRef = null;
        var popupMenuPath = new List<string>(); // top-down MenuItem names, e.g. ["QA", "Level 1"]
        if (element?.ControlType == ControlType.MenuItem)
        {
            try
            {
                var pathNames = new List<string>();
                bool insidePopup = false;
                var cur = element;
                for (int depth = 0; depth < MaxMenuAncestorDepth; depth++)
                {
                    if (cur.ControlType == ControlType.MenuItem)
                    {
                        // Only include items with a non-empty name; nameless items
                        // cannot be reliably re-located by FindFirstDescendant(ByName).
                        var itemName = cur.Name ?? string.Empty;
                        if (!string.IsNullOrEmpty(itemName))
                            pathNames.Add(itemName);
                    }
                    else if (cur.ControlType == ControlType.Menu)
                        insidePopup = true;
                    else if (cur.ControlType == ControlType.MenuBar)
                    {
                        // Paths with only one entry are top-level MenuItems (direct children
                        // of the MenuBar) — their element reference stays valid and the normal
                        // Invoke() path works fine.  Re-navigation is only needed for items
                        // that are two or more levels deep (inside a popup dropdown).
                        if (insidePopup && pathNames.Count > 1)
                        {
                            popupMenuBarRef = cur;
                            pathNames.Reverse(); // collected bottom-up; reverse to top-down
                            popupMenuPath = pathNames;
                        }
                        break;
                    }
                    var par = cur.Parent;
                    if (par == null) break;
                    cur = par;
                }
            }
            catch { /* best effort */ }
        }

        // When an inner Edit was promoted to its parent ComboBox, keep a reference to the
        // original Edit so "Type…" and "Is Editable" can still target the actual text field.
        var innerEditElement = (originalElement != null &&
                                !ReferenceEquals(originalElement, element) &&
                                originalElement.ControlType == ControlType.Edit)
            ? originalElement
            : null;
        var innerEditInfo = innerEditElement != null ? BuildElementInfo(innerEditElement) : null;

        var menu = new NoActivateContextMenuStrip { ShowImageMargin = false };
        menu.Font = new Font("Segoe UI", 10f);

        // Re-apply click-through styles when the menu closes so rapid right-clicks
        // cannot leave the overlay in a non-transparent state.
        menu.Closed += (_, _) =>
        {
            _menuOpen = false;
            ReapplyClickThroughStyle();

            // After the menu closes, check if a new child Window has appeared under the
            // same window ancestor (e.g. a "Working Environment" popup opened because
            // the user clicked OK).  If so, nudge the user via the status label.
            if (windowAncestor != null && _automation != null)
            {
                try
                {
                    var cf = _automation.ConditionFactory;
                    var currentChildWindows = windowAncestor.FindAllChildren(
                        cf.ByControlType(ControlType.Window));
                    var newTitles = currentChildWindows
                        .Select(w => { try { return w.Name ?? string.Empty; } catch { return string.Empty; } })
                        .Where(t => !preActionChildWindowTitles.Contains(t))
                        .ToList();
                    if (newTitles.Count > 0)
                    {
                        var newTitle = newTitles[0];
                        _statusLabel.Text =
                            $"  ⚡ New window: '{newTitle}'  — Ctrl+Right-click to target it  ";
                    }
                }
                catch { /* best effort */ }
            }
        };

        // ── Header (element details, not selectable) ────────────────────────
        var headerName = elementInfo?.Name ?? elementInfo?.AutomationId ?? "(no name)";
        var headerType = elementInfo?.ControlType ?? "Unknown";
        var header = new ToolStripMenuItem($"[{headerType}]  {headerName}")
        {
            Enabled = false,
            ForeColor = Color.DimGray
        };
        menu.Items.Add(header);
        menu.Items.Add(new ToolStripSeparator());

        // ── Interactive actions ──────────────────────────────────────────────
        AddActionItem(menu, "Click", element, elementInfo, ActionType.Click,
            () =>
            {
                if (popupMenuBarRef != null && popupMenuPath.Count > 1)
                {
                    // The element is a sub-MenuItem inside a popup/dropdown that has
                    // already collapsed (focus moved to this assistive context menu).
                    // Re-navigate the menu chain from the stable MenuBar reference,
                    // finding fresh UIA elements at each level.
                    BringElementWindowToForeground(popupMenuBarRef);
                    Thread.Sleep(WindowActivationDelayMs);

                    AutomationElement? stepAnchor = popupMenuBarRef;
                    for (int i = 0; i < popupMenuPath.Count; i++)
                    {
                        var stepName = popupMenuPath[i];
                        if (stepAnchor == null) break;

                        AutomationElement? stepItem = null;
                        try
                        {
                            var cf = _automation!.ConditionFactory;
                            stepItem = stepAnchor.FindFirstDescendant(
                                cf.ByControlType(ControlType.MenuItem).And(cf.ByName(stepName)));
                        }
                        catch { break; }

                        if (stepItem == null) break;

                        bool isLast = i == popupMenuPath.Count - 1;
                        try
                        {
                            if (stepItem.Patterns.Invoke.IsSupported)
                                stepItem.Patterns.Invoke.Pattern.Invoke();
                            else
                                stepItem.Click();
                        }
                        catch { break; }

                        if (!isLast)
                        {
                            // Wait for the intermediate dropdown to materialise before
                            // searching for the next item.
                            Thread.Sleep(MenuNavigationDelayMs);
                            stepAnchor = stepItem;
                        }
                    }
                }
                else if (element?.Patterns.Invoke.IsSupported == true)
                {
                    element.Patterns.Invoke.Pattern.Invoke();
                }
                else
                {
                    // Bring the element's root window to the foreground via Win32
                    // SetForegroundWindow before firing a physical mouse click.
                    // This is more reliable than UIA element.Focus() alone: when the
                    // assistive context menu is dismissed, Windows can hand focus back
                    // to IntelliJ (the process that launched the driver) before the
                    // action handler executes — causing the click to land on IntelliJ.
                    // The driver process still holds the foreground lock immediately
                    // after the menu closes, so SetForegroundWindow is allowed here.
                    BringElementWindowToForeground(element);
                    // Brief pause to let the window activation take effect before
                    // the physical mouse click fires; without this, Windows may not
                    // have finished processing the SetForegroundWindow call and the
                    // click can still land on the previously-focused application.
                    Thread.Sleep(WindowActivationDelayMs);
                    // Click at the exact point the user right-clicked rather than
                    // re-querying GetClickablePoint() from the element.  This is
                    // essential for Unknown-framework elements (e.g. Java Swing top-level
                    // MenuItems) whose UIA BoundingRectangle / ClickablePoint can be the
                    // full menu-bar row rather than the specific item under the cursor,
                    // causing clicks to land at the wrong position.
                    Mouse.Click(pt);
                }
            });
        AddActionItem(menu, "Double Click", element, elementInfo, ActionType.DoubleClick,
            () =>
            {
                BringElementWindowToForeground(element);
                Thread.Sleep(WindowActivationDelayMs);
                Mouse.DoubleClick(pt);
            });
        AddActionItem(menu, "Right Click", element, elementInfo, ActionType.RightClick,
            () =>
            {
                // Prefer the HWND captured before the menu appeared so that the correct
                // window receives focus even when the element becomes stale (e.g. a
                // modal popup closed on click-based dismissal).  Fall back to the
                // standard BringElementWindowToForeground when no HWND was captured.
                if (capturedHwnd != IntPtr.Zero)
                    SetForegroundWindow(capturedHwnd);
                else
                    BringElementWindowToForeground(element);
                Thread.Sleep(WindowActivationDelayMs);
                // Set the flag so the global hook does not intercept this simulated
                // right-click and show the assistive context menu a second time.
                _suppressNextRButtonDown = true;
                Mouse.RightClick(pt);
            });
        AddActionItem(menu, "Hover", element, elementInfo, ActionType.Hover,
            () => { /* cursor is already there */ });

        // Select — only for selectable elements
        if (element != null && IsSelectable(element))
        {
            AddActionItem(menu, "Select", element, elementInfo, ActionType.Select,
                () =>
                {
                    try { element.Patterns.SelectionItem.Pattern.Select(); }
                    catch { element.Click(); }
                });
        }

        // Drag and Drop — available for every element, regardless of control type.
        // Selecting this item sets the current element as the drag source; the user then
        // right-clicks the drop target to complete the recorded action.
        var dragDropItem = new ToolStripMenuItem("Drag and Drop…");
        dragDropItem.Click += (_, _) =>
        {
            _awaitingDragTarget = true;
            _dragSourceInfo = elementInfo;
            _statusLabel.Text = "  ● Drag source set — Right-click the drop target element  ";
        };
        menu.Items.Add(dragDropItem);

        // Sub-Menu Items — when the right-clicked element is a MenuItem that has child
        // MenuItems in the UIA tree (e.g. "QA" with children "1", "2", "3").
        //
        // Background: clicking a top-level MenuItem (e.g. "QA") causes WinForms to create a
        // floating ControlType=Menu element ("QADropDown") at the desktop level.  That element
        // has no UIA children — the sub-MenuItems remain children of "QA" in the UIA tree — but
        // the dynamically-created popup makes it hard to click sub-items via a second right-click.
        //
        // This flyout lets the user pick the sub-item directly from the assistive context menu
        // shown on the *first* right-click (on the parent "QA" item), before any popup appears.
        // The click handler navigates the full path from the stable MenuBar anchor, re-fetching
        // fresh UIA elements at each step — the same approach used by the existing top-level
        // Click action for deep sub-menu paths (popupMenuBarRef / popupMenuPath).
        if (element?.ControlType == ControlType.MenuItem && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                var subItems = element.FindAllChildren(cf.ByControlType(ControlType.MenuItem));

                if (subItems.Length > 0)
                {
                    // Walk up from element to the MenuBar to build a stable anchor and a
                    // top-down path of MenuItem names (e.g. element="QA" → ["QA"]).
                    AutomationElement? subMenuBarRef = null;
                    var parentMenuPath = new List<string>();
                    try
                    {
                        var pathNames = new List<string>();
                        var cur = element;
                        for (int d = 0; d < MaxMenuAncestorDepth; d++)
                        {
                            if (cur.ControlType == ControlType.MenuItem)
                            {
                                var n = cur.Name ?? string.Empty;
                                if (!string.IsNullOrEmpty(n))
                                    pathNames.Add(n);
                            }
                            else if (cur.ControlType == ControlType.MenuBar)
                            {
                                subMenuBarRef = cur;
                                pathNames.Reverse(); // collected bottom-up; reverse to top-down
                                parentMenuPath = pathNames;
                                break;
                            }
                            var par = cur.Parent;
                            if (par == null) break;
                            cur = par;
                        }
                    }
                    catch { /* best effort */ }

                    // Only offer the flyout when we have a valid MenuBar anchor; without it
                    // we cannot re-locate fresh elements after the popup closes.
                    if (subMenuBarRef != null)
                    {
                        var capturedMenuBarRef = subMenuBarRef;
                        var capturedParentMenuPath = parentMenuPath;

                        var subMenuFlyout = new ToolStripMenuItem("Sub-Menu Items ▶");
                        foreach (var subItem in subItems.Take(MaxChildrenToDisplay))
                        {
                            var subItemInfo = BuildElementInfo(subItem);
                            var subItemName = subItemInfo.Name ?? subItemInfo.AutomationId ?? "(unnamed)";
                            var subItemEntry = new ToolStripMenuItem(subItemName);
                            var capturedSubItemInfo = subItemInfo;
                            var capturedSubItemName = subItemName;

                            subItemEntry.Click += (_, _) =>
                            {
                                // Navigate: click each ancestor MenuItem to expand the chain,
                                // then click the target sub-item.  Re-fetch fresh UIA elements
                                // at every step so stale references never cause silent failures.
                                BringElementWindowToForeground(capturedMenuBarRef);
                                Thread.Sleep(WindowActivationDelayMs);

                                // Full path from MenuBar anchor to sub-item, e.g. ["QA", "1"].
                                var fullPath = new List<string>(capturedParentMenuPath) { capturedSubItemName };

                                AutomationElement? stepAnchor = capturedMenuBarRef;
                                for (int i = 0; i < fullPath.Count; i++)
                                {
                                    var stepName = fullPath[i];
                                    if (stepAnchor == null) break;

                                    AutomationElement? stepItem = null;
                                    try
                                    {
                                        var cf2 = _automation!.ConditionFactory;
                                        stepItem = stepAnchor.FindFirstDescendant(
                                            cf2.ByControlType(ControlType.MenuItem).And(cf2.ByName(stepName)));
                                    }
                                    catch { break; }

                                    if (stepItem == null) break;

                                    bool isLast = i == fullPath.Count - 1;
                                    try
                                    {
                                        if (stepItem.Patterns.Invoke.IsSupported)
                                            stepItem.Patterns.Invoke.Pattern.Invoke();
                                        else
                                            stepItem.Click();
                                    }
                                    catch { break; }

                                    if (!isLast)
                                    {
                                        // Wait for the intermediate dropdown to materialise
                                        // before searching for the next item in the chain.
                                        Thread.Sleep(MenuNavigationDelayMs);
                                        stepAnchor = stepItem;
                                    }
                                }

                                // Record two Click actions: one for the parent MenuItem that
                                // was right-clicked (e.g. "QA"), and one for the child item
                                // that was selected from the flyout (e.g. "Level 35").
                                // This mirrors the two-step navigation needed for automation
                                // (expand parent, then click child) and makes the exported
                                // JSON self-contained and replay-friendly.
                                var parentLabel = ElementInfo.GetLabel(elementInfo);
                                _service.AddAction(new RecordedAction
                                {
                                    ActionType = ActionType.Click,
                                    Mode = RecordingMode.Assistive,
                                    Element = elementInfo,
                                    Description = $"Click on {parentLabel}"
                                });
                                _service.AddAction(new RecordedAction
                                {
                                    ActionType = ActionType.Click,
                                    Mode = RecordingMode.Assistive,
                                    Element = capturedSubItemInfo,
                                    Description = $"Click on {ElementInfo.GetLabel(capturedSubItemInfo)}"
                                });
                                UpdateStatusAfterAction($"Click [{capturedSubItemInfo.ControlType}] {capturedSubItemName}");
                            };
                            subMenuFlyout.DropDownItems.Add(subItemEntry);
                        }

                        if (subItems.Length > MaxChildrenToDisplay)
                            subMenuFlyout.DropDownItems.Add(
                                new ToolStripMenuItem($"… and {subItems.Length - MaxChildrenToDisplay} more") { Enabled = false });

                        menu.Items.Add(subMenuFlyout);
                    }
                }
            }
            catch { /* best effort */ }
        }

        menu.Items.Add(new ToolStripSeparator());

        // ── Query actions ────────────────────────────────────────────────────
        AddQueryItem(menu, "Is Visible", element, elementInfo, ActionType.IsVisible,
            () => element != null && !element.IsOffscreen);
        AddQueryItem(menu, "Is Clickable", element, elementInfo, ActionType.IsClickable,
            () => element != null && !element.IsOffscreen && element.IsEnabled &&
                  (element.Patterns.Invoke.IsSupported || element.Patterns.Toggle.IsSupported));
        AddQueryItem(menu, "Is Enabled", element, elementInfo, ActionType.IsEnabled,
            () => element?.IsEnabled ?? false);
        AddQueryItem(menu, "Is Disabled", element, elementInfo, ActionType.IsDisabled,
            () => element != null && !element.IsEnabled);

        // Resolve the Edit element to use for "Is Editable" and "Type…":
        // - if the resolved element is itself an Edit or Document, use it directly;
        // - if it exposes a writable Value pattern (covers custom/Java text fields that do
        //   not use ControlType.Edit, e.g. IntelliJ Swing components via Java Access Bridge),
        //   use it directly;
        // - otherwise fall back to innerEditElement (the original Edit that was promoted
        //   to its parent ComboBox, e.g. a username/password field with credential history).
        bool isDirectlyEditable = element != null &&
            (element.ControlType == ControlType.Edit ||
             element.ControlType == ControlType.Document ||
             (element.Patterns.Value.IsSupported &&
              element.Patterns.Value.PatternOrDefault?.IsReadOnly == false));
        var editTarget = isDirectlyEditable ? element : innerEditElement;
        var editTargetInfo = isDirectlyEditable ? elementInfo : innerEditInfo;

        // Final fallback: if no editable target was found yet, look for an Edit or Document
        // child of the current element.  FromPoint() in some frameworks (e.g. Java Access
        // Bridge used by IntelliJ) returns the outer wrapper Pane/Group rather than the
        // inner text-input control, so we drill down one level here.
        if (editTarget == null && element != null && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                var editChild = element.FindFirstDescendant(
                    cf.ByControlType(ControlType.Edit).Or(cf.ByControlType(ControlType.Document)));
                if (editChild != null)
                {
                    editTarget = editChild;
                    editTargetInfo = BuildElementInfo(editChild);
                }
            }
            catch { /* best effort */ }
        }

        // Is Editable — for Edit controls (text boxes), or the inner Edit of a promoted ComboBox.
        if (editTarget != null)
        {
            AddQueryItem(menu, "Is Editable", editTarget, editTargetInfo, ActionType.IsEditable,
                () => editTarget.Patterns.Value.IsSupported && editTarget.Patterns.Value.Pattern?.IsReadOnly == false);
        }

        // Type — for Edit controls (text boxes), or the inner Edit of a promoted ComboBox.
        // Exception: do not show "Type" when the target is a disabled Edit control.
        bool isDisabledEdit = false;
        try
        {
            isDisabledEdit = editTarget != null &&
                             editTarget.ControlType == ControlType.Edit &&
                             !editTarget.IsEnabled;
        }
        catch { /* best effort; if we can't determine, show Type */ }

        if (editTarget != null && !isDisabledEdit)
        {
            menu.Items.Add(new ToolStripSeparator());
            var typeItem = new ToolStripMenuItem("Type…");
            typeItem.Click += (_, _) =>
            {
                var text = ShowTypePrompt(editTargetInfo?.Name ?? editTargetInfo?.AutomationId ?? "element");
                if (text == null) return; // user cancelled

                var elementLabel = ElementInfo.GetLabel(editTargetInfo);
                try
                {
                    if (editTarget.Patterns.Value.IsSupported)
                        editTarget.Patterns.Value.Pattern.SetValue(text);
                    else
                    {
                        editTarget.Focus();
                        System.Windows.Forms.SendKeys.SendWait(EscapeForSendKeys(text));
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Type action failed for element \"{Label}\"", elementLabel);
                }

                // Use single quotes in the description to avoid breaking the string when text contains double quotes
                var displayText = text.Replace("'", "\\'", StringComparison.Ordinal);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.Type,
                    Mode = RecordingMode.Assistive,
                    Element = editTargetInfo,
                    Value = text,
                    Description = $"Type '{displayText}' into {elementLabel}"
                });
                UpdateStatusAfterAction($"Type into [{editTargetInfo?.ControlType}] {elementLabel}");
            };
            menu.Items.Add(typeItem);

            // Clear — only for writable Edit controls
            if (editTarget.Patterns.Value.IsSupported &&
                editTarget.Patterns.Value.PatternOrDefault?.IsReadOnly == false)
            {
                AddActionItem(menu, "Clear", editTarget, editTargetInfo, ActionType.ClearText,
                    () =>
                    {
                        var valuePat = editTarget.Patterns.Value.PatternOrDefault;
                        valuePat?.SetValue(string.Empty);
                    });
            }

            // Get Text — read the current text / value of the Edit field
            AddQueryItem(menu, "Get Text", editTarget, editTargetInfo, ActionType.GetValue,
                () =>
                {
                    try
                    {
                        var valuePat = editTarget.Patterns.Value.PatternOrDefault;
                        if (valuePat != null)
                            return !string.IsNullOrEmpty(valuePat.Value);
                    }
                    catch { /* best effort */ }
                    return false;
                });
        }

        // Type and Select — only for editable ComboBox controls (autocomplete / filter pattern)
        if (element != null && element.ControlType == ControlType.ComboBox)
        {
            var valuePat = element.Patterns.Value.PatternOrDefault;
            if (valuePat != null && !valuePat.IsReadOnly)
            {
                menu.Items.Add(new ToolStripSeparator());
                var typeAndSelectItem = new ToolStripMenuItem("Type and Select…");
                typeAndSelectItem.Click += (_, _) =>
                {
                    var text = ShowTypePrompt(elementInfo?.Name ?? elementInfo?.AutomationId ?? "ComboBox");
                    if (text == null) return; // user cancelled

                    var elementLabel = ElementInfo.GetLabel(elementInfo);
                    try
                    {
                        // Focus the combo and type the filter text.
                        element.Focus();
                        try { valuePat.SetValue(text); }
                        catch
                        {
                            // Value pattern may throw even when it reports writable;
                            // fall back to simulated keystrokes.
                            try { valuePat.SetValue(string.Empty); } catch { }
                            System.Windows.Forms.SendKeys.SendWait(EscapeForSendKeys(text));
                        }

                        Thread.Sleep(500); // allow the dropdown to populate

                        // Find and select the matching item.
                        if (_automation != null)
                        {
                            var cf = _automation.ConditionFactory;
                            var items = element.FindAllDescendants(
                                cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem));

                            if (items.Length == 0)
                            {
                                var listChild = element.FindFirstDescendant(
                                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.List));
                                if (listChild != null)
                                    items = listChild.FindAllChildren();
                            }

                            if (items.Length > 0)
                            {
                                var target = items.FirstOrDefault(i =>
                                    i.Name.Equals(text, StringComparison.OrdinalIgnoreCase))
                                    ?? items[0];

                                target.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();
                                var selPat = target.Patterns.SelectionItem.PatternOrDefault;
                                if (selPat != null)
                                    selPat.Select();
                                else
                                    target.Click();

                                Thread.Sleep(100);
                                element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Type and Select failed for element \"{Label}\"", elementLabel);
                    }

                    var displayText = text.Replace("'", "\\'", StringComparison.Ordinal);
                    _service.AddAction(new RecordedAction
                    {
                        ActionType = ActionType.TypeAndSelect,
                        Mode = RecordingMode.Assistive,
                        Element = elementInfo,
                        Value = text,
                        Description = $"Type and select '{displayText}' in {elementLabel}"
                    });
                    UpdateStatusAfterAction($"Type and select into [{elementInfo?.ControlType}] {elementLabel}");
                };
                menu.Items.Add(typeAndSelectItem);
            }

            // Get Selected Value — read the currently selected / displayed value of the ComboBox
            menu.Items.Add(new ToolStripSeparator());
            AddQueryItem(menu, "Get Selected Value", element, elementInfo, ActionType.GetValue,
                () =>
                {
                    try
                    {
                        if (element.Patterns.Value.IsSupported)
                            return !string.IsNullOrEmpty(element.Patterns.Value.Pattern.Value);
                        if (element.Patterns.Selection.IsSupported)
                        {
                            var sel = element.Patterns.Selection.Pattern.Selection.Value;
                            return sel.Length > 0;
                        }
                    }
                    catch { /* best effort */ }
                    return false;
                });
        }

        // Get Table Headers — only for HeaderItem controls (column headers inside a DataGrid / Table)
        if (element != null && element.ControlType == ControlType.HeaderItem)
        {
            menu.Items.Add(new ToolStripSeparator());
            var getHeadersItem = new ToolStripMenuItem("Get Table Headers");
            getHeadersItem.Click += (_, _) =>
            {
                // Walk up the tree to find the parent DataGrid / Table / Header element
                // so the recorded action targets the table container, not the individual
                // header cell.
                AutomationElement tableElement = element;
                ElementInfo? tableInfo = elementInfo;
                try
                {
                    var current = element;
                    for (int i = 0; i < MaxTableAncestorDepth; i++)
                    {
                        var parent = current.Parent;
                        if (parent == null) break;
                        if (parent.ControlType == ControlType.DataGrid ||
                            parent.ControlType == ControlType.Table)
                        {
                            tableElement = parent;
                            tableInfo = BuildElementInfo(parent);
                            break;
                        }
                        current = parent;
                    }
                }
                catch { /* best effort – fall back to the header item itself */ }

                var tableLabel = ElementInfo.GetLabel(tableInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.GetTableHeaders,
                    Mode = RecordingMode.Assistive,
                    Element = tableInfo,
                    Description = $"Get table headers from {tableLabel}"
                });
                UpdateStatusAfterAction($"Get Table Headers on [{tableInfo?.ControlType}] {tableLabel}");
            };
            menu.Items.Add(getHeadersItem);
        }

        // Get Table Headers + Get Table Data — directly on DataGrid / Table elements
        if (element != null &&
            (element.ControlType == ControlType.DataGrid ||
             element.ControlType == ControlType.Table))
        {
            menu.Items.Add(new ToolStripSeparator());
            var directHeadersItem = new ToolStripMenuItem("Get Table Headers");
            directHeadersItem.Click += (_, _) =>
            {
                var label = ElementInfo.GetLabel(elementInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.GetTableHeaders,
                    Mode = RecordingMode.Assistive,
                    Element = elementInfo,
                    Description = $"Get table headers from {label}"
                });
                UpdateStatusAfterAction($"Get Table Headers on [{elementInfo?.ControlType}] {label}");
            };
            menu.Items.Add(directHeadersItem);

            var tableDataItem = new ToolStripMenuItem("Get Table Data");
            tableDataItem.Click += (_, _) =>
            {
                var label = ElementInfo.GetLabel(elementInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.GetTableData,
                    Mode = RecordingMode.Assistive,
                    Element = elementInfo,
                    Description = $"Get table data from {label}"
                });
                UpdateStatusAfterAction($"Get Table Data on [{elementInfo?.ControlType}] {label}");
            };
            menu.Items.Add(tableDataItem);
        }

        // Is Checked / Check / Uncheck — only for CheckBox controls
        if (element != null && element.ControlType == ControlType.CheckBox)
        {
            menu.Items.Add(new ToolStripSeparator());
            AddQueryItem(menu, "Is Checked", element, elementInfo, ActionType.IsChecked,
                () => element.Patterns.Toggle.IsSupported &&
                      element.Patterns.Toggle.Pattern.ToggleState == FlaUI.Core.Definitions.ToggleState.On);

            // Determine the current checked state so the label reflects what will happen.
            bool currentlyChecked = false;
            try
            {
                currentlyChecked = element.Patterns.Toggle.IsSupported &&
                    element.Patterns.Toggle.Pattern.ToggleState == FlaUI.Core.Definitions.ToggleState.On;
            }
            catch { /* best effort; default to showing "Check" */ }

            var checkActionLabel = currentlyChecked ? "Uncheck" : "Check";
            // Capture the desired state as a bool so the action lambda doesn't depend
            // on the label text (avoids coupling UI copy to control-flow logic).
            bool wantChecked = !currentlyChecked;
            AddActionItem(menu, checkActionLabel, element, elementInfo, ActionType.SelectCheckBox,
                () =>
                {
                    if (element.Patterns.Toggle.IsSupported)
                    {
                        var state = element.Patterns.Toggle.Pattern.ToggleState;
                        if ((wantChecked && state != FlaUI.Core.Definitions.ToggleState.On) ||
                            (!wantChecked && state == FlaUI.Core.Definitions.ToggleState.On))
                        {
                            element.Patterns.Toggle.Pattern.Toggle();
                        }
                    }
                    else
                    {
                        element.Click();
                    }
                });
        }

        // Assert — only for Text controls (static text labels)
        if (element != null && element.ControlType == ControlType.Text)
        {
            menu.Items.Add(new ToolStripSeparator());
            var assertItem = new ToolStripMenuItem("Assert");
            assertItem.Click += (_, _) =>
            {
                var textValue = element.Name ?? string.Empty;
                var elementLabel = ElementInfo.GetLabel(elementInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.Assert,
                    Mode = RecordingMode.Assistive,
                    Element = elementInfo,
                    Value = textValue,
                    Description = $"Assert text '{textValue}' on {elementLabel}"
                });
                UpdateStatusAfterAction($"Assert '{textValue}' on [{elementInfo?.ControlType}] {elementLabel}");
            };
            menu.Items.Add(assertItem);
        }

        // Window operations — only for Window controls
        if (element != null && element.ControlType == ControlType.Window &&
            element.Patterns.Window.IsSupported)
        {
            menu.Items.Add(new ToolStripSeparator());
            AddActionItem(menu, "Maximize", element, elementInfo, ActionType.Maximize,
                () => element.Patterns.Window.Pattern.SetWindowVisualState(
                    FlaUI.Core.Definitions.WindowVisualState.Maximized));
            AddActionItem(menu, "Minimize", element, elementInfo, ActionType.Minimize,
                () => element.Patterns.Window.Pattern.SetWindowVisualState(
                    FlaUI.Core.Definitions.WindowVisualState.Minimized));
            AddActionItem(menu, "Close Window", element, elementInfo, ActionType.CloseWindow,
                () => element.Patterns.Window.Pattern.Close());
        }

        // Switch Window — available when right-clicking a TitleBar element.
        // Walks up the UIA tree to locate the parent Window so its title can be used.
        if (element != null && element.ControlType == ControlType.TitleBar)
        {
            // Find the parent Window element (at most 5 levels up).
            AutomationElement? windowElement = null;
            ElementInfo? windowInfo = null;
            try
            {
                var current = element;
                for (int i = 0; i < MaxWindowSearchDepth; i++)
                {
                    var parent = current.Parent;
                    if (parent == null) break;
                    if (parent.ControlType == ControlType.Window)
                    {
                        windowElement = parent;
                        windowInfo = BuildElementInfo(parent);
                        break;
                    }
                    current = parent;
                }
            }
            catch { /* best effort */ }

            // Fall back to the TitleBar itself if no Window parent was found.
            var switchTarget = windowElement ?? element;
            var switchInfo = windowInfo ?? elementInfo;
            var windowTitle = switchInfo?.Name ?? string.Empty;

            menu.Items.Add(new ToolStripSeparator());
            var switchItem = new ToolStripMenuItem("Switch Window");
            switchItem.Click += (_, _) =>
            {
                var label = ElementInfo.GetLabel(switchInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.SwitchWindow,
                    Mode = RecordingMode.Assistive,
                    Element = switchInfo,
                    Value = windowTitle,
                    Description = $"Switch window to '{windowTitle}'"
                });
                UpdateStatusAfterAction($"Switch Window [{switchInfo?.ControlType}] {label}");
            };
            menu.Items.Add(switchItem);
        }

        // Expand / Collapse — for TreeItem, Tree children, MenuItem with submenus, and any
        // element with ExpandCollapsePattern (excluding ComboBox, which is handled by Options ▶).
        if (element != null &&
            element.ControlType != ControlType.ComboBox &&
            element.Patterns.ExpandCollapse.IsSupported)
        {
            menu.Items.Add(new ToolStripSeparator());
            AddActionItem(menu, "Expand", element, elementInfo, ActionType.Expand,
                () => element.Patterns.ExpandCollapse.Pattern.Expand());
            AddActionItem(menu, "Collapse", element, elementInfo, ActionType.Collapse,
                () => element.Patterns.ExpandCollapse.Pattern.Collapse());
        }

        // RangeValue controls (Slider, ProgressBar, ScrollBar, Spinner)
        if (element != null && element.Patterns.RangeValue.IsSupported)
        {
            menu.Items.Add(new ToolStripSeparator());

            // Get Numeric Value — reads the current numeric value; always available when the pattern is supported
            AddQueryItem(menu, "Get Numeric Value", element, elementInfo, ActionType.GetValue,
                () => true);

            // Set Value — prompts the user for a numeric value (only for writable controls)
            if (!element.Patterns.RangeValue.Pattern.IsReadOnly)
            {
                var setValueItem = new ToolStripMenuItem("Set Value…");
                setValueItem.Click += (_, _) =>
                {
                    var elemLabel = ElementInfo.GetLabel(elementInfo);
                    var input = ShowValuePrompt(elemLabel,
                        element.Patterns.RangeValue.Pattern.Minimum,
                        element.Patterns.RangeValue.Pattern.Maximum);
                    if (input == null) return; // user cancelled

                    try
                    {
                        if (double.TryParse(input, out var val))
                            element.Patterns.RangeValue.Pattern.SetValue(val);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Set Value failed for element \"{Label}\"", elemLabel);
                    }

                    _service.AddAction(new RecordedAction
                    {
                        ActionType = ActionType.SetValue,
                        Mode = RecordingMode.Assistive,
                        Element = elementInfo,
                        Value = input,
                        Description = $"Set value to '{input}' on {elemLabel}"
                    });
                    UpdateStatusAfterAction($"Set Value '{input}' on [{elementInfo?.ControlType}] {elemLabel}");
                };
                menu.Items.Add(setValueItem);
            }
        }

        // Scroll — for elements that support the Scroll pattern (List, ScrollBar, etc.)
        if (element != null && element.Patterns.Scroll.IsSupported)
        {
            menu.Items.Add(new ToolStripSeparator());
            AddActionItem(menu, "Scroll Up", element, elementInfo, ActionType.Scroll,
                () => element.Patterns.Scroll.Pattern.Scroll(
                    FlaUI.Core.Definitions.ScrollAmount.NoAmount,
                    FlaUI.Core.Definitions.ScrollAmount.SmallDecrement));
            AddActionItem(menu, "Scroll Down", element, elementInfo, ActionType.Scroll,
                () => element.Patterns.Scroll.Pattern.Scroll(
                    FlaUI.Core.Definitions.ScrollAmount.NoAmount,
                    FlaUI.Core.Definitions.ScrollAmount.SmallIncrement));
        }

        // ── Children submenu (for container controls and expandable edit/combo fields) ───
        // Note: ComboBox is checked explicitly here rather than being added to IsContainer()
        // because IsContainer() is also used by DrillDownToElementAtPoint() — drilling into
        // a ComboBox's structural children (Edit, Button, List) would break element identification.
        if (element != null && (IsContainer(element.ControlType) || element.ControlType == ControlType.ComboBox || element.Patterns.ExpandCollapse.IsSupported))
        {
            try
            {
                // For ComboBox controls expand the dropdown to materialise list items,
                // then collapse immediately so the user does not see the dropdown open
                // while the context menu is being built.
                // Do NOT expand MenuItem / Menu / MenuBar elements: expanding them
                // opens the native submenu popup on the message-pump thread which causes
                // a 10-15 s freeze and visual artefacts.  Their children are accessible
                // via FindAllChildren() without expansion.
                bool isComboBox = element.ControlType == ControlType.ComboBox;
                bool isMenuRelated = element.ControlType == ControlType.MenuItem
                    || element.ControlType == ControlType.Menu
                    || element.ControlType == ControlType.MenuBar;

                if (isComboBox && element.Patterns.ExpandCollapse.IsSupported)
                {
                    try { element.Patterns.ExpandCollapse.Pattern.Expand(); }
                    catch { /* best effort */ }
                    Thread.Sleep(300); // brief pause so list items materialize
                }

                // For ComboBox elements, retrieve the ListItem descendants (the actual
                // dropdown options) rather than the immediate structural children
                // (Edit, Button, List) which are not useful to the user.
                AutomationElement[] children;
                string childrenMenuLabel;
                if (isComboBox && _automation != null)
                {
                    var cf = _automation.ConditionFactory;

                    // First try: ListItem descendants directly under the ComboBox.
                    children = element.FindAllDescendants(cf.ByControlType(ControlType.ListItem));
                    childrenMenuLabel = children.Length > 0 ? "Options ▶" : "Children ▶";

                    if (children.Length == 0)
                    {
                        // Second try: find the List child and enumerate its children.
                        // Many combo implementations nest items inside a List container.
                        var listChild = element.FindFirstDescendant(cf.ByControlType(ControlType.List));
                        if (listChild != null)
                        {
                            children = listChild.FindAllChildren();
                            if (children.Length > 0)
                                childrenMenuLabel = "Options ▶";
                        }
                    }

                    // Last resort: immediate children of the ComboBox.
                    if (children.Length == 0)
                        children = element.FindAllChildren();

                    // Deduplicate: some ComboBox UIA implementations expose the same
                    // logical option at multiple levels of the accessibility tree
                    // (e.g. once as a ListItem descendant and again as a child of a
                    // nested List container). Remove subsequent occurrences of any
                    // (Name, AutomationId) pair so the menu shows each option once.
                    // Use ordinal (case-sensitive) comparison to preserve distinct
                    // options whose names differ only in letter case.
                    if (children.Length > 0)
                    {
                        var seen = new HashSet<(string, string)>();
                        children = children.Where(c =>
                        {
                            var key = (c.Name ?? string.Empty, c.AutomationId ?? string.Empty);
                            return seen.Add(key);
                        }).ToArray();
                    }

                    // Collapse the dropdown so it does not remain open while the context
                    // menu is visible (the user would see an expanded combo behind the menu).
                    try { element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse(); }
                    catch { /* best effort */ }
                }
                else if (isMenuRelated && _automation != null)
                {
                    // For menu elements enumerate children without expanding so that no
                    // submenu popup is opened.  The UIA tree exposes sub-items even while
                    // the menu is visually collapsed.
                    children = element.FindAllChildren();
                    childrenMenuLabel = "Children ▶";
                }
                else
                {
                    children = element.FindAllChildren();
                    childrenMenuLabel = "Children ▶";
                }

                if (children.Length > 0)
                {
                    menu.Items.Add(new ToolStripSeparator());
                    var childrenMenu = new ToolStripMenuItem(childrenMenuLabel);
                    foreach (var child in children.Take(MaxChildrenToDisplay))
                    {
                        var childInfo = BuildElementInfo(child);
                        var label = $"[{childInfo.ControlType}]  {childInfo.Name ?? childInfo.AutomationId ?? "(unnamed)"}";
                        var childItem = new ToolStripMenuItem(label);
                        var capturedChildInfo = childInfo;

                        childItem.Click += (_, _) =>
                        {
                            // Re-expand the parent so that fresh list items are materialised
                            // in the UIA tree before we attempt to select one.  The item
                            // reference captured during menu construction is stale once the
                            // dropdown has been collapsed, so we re-locate it by name or
                            // AutomationId after expanding.
                            try { element.Patterns.ExpandCollapse.PatternOrDefault?.Expand(); }
                            catch { /* best effort */ }

                            Thread.Sleep(200); // 200 ms lets UIA child elements re-materialise after Expand()

                            bool selected = false;
                            if (_automation != null)
                            {
                                try
                                {
                                    var cf = _automation.ConditionFactory;

                                    // Try to find a fresh reference by AutomationId first,
                                    // then fall back to matching by Name.
                                    AutomationElement? freshItem = null;
                                    if (!string.IsNullOrEmpty(capturedChildInfo.AutomationId))
                                        freshItem = element.FindFirstDescendant(
                                            cf.ByAutomationId(capturedChildInfo.AutomationId));

                                    if (freshItem == null && !string.IsNullOrEmpty(capturedChildInfo.Name))
                                        freshItem = element.FindFirstDescendant(
                                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.ListItem)
                                              .And(cf.ByName(capturedChildInfo.Name)));

                                    if (freshItem != null)
                                    {
                                        freshItem.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();

                                        // Strategy 1: Click() — reliably triggers ComboBox
                                        // selection events and updates the displayed value.
                                        freshItem.Click();

                                        // Strategy 2: Also call SelectionItem.Select() to
                                        // ensure the selection state is committed in case
                                        // Click() alone did not fully populate the ComboBox
                                        // (e.g. WPF or custom combo implementations).
                                        try { freshItem.Patterns.SelectionItem.PatternOrDefault?.Select(); }
                                        catch { /* best effort */ }

                                        Thread.Sleep(100);
                                        selected = true;
                                    }
                                }
                                catch { /* best effort */ }
                            }

                            if (!selected)
                            {
                                // Fallback: use the (possibly stale) original reference.
                                try
                                {
                                    // Try SelectionItem.Select() first on the original ref.
                                    if (child.Patterns.SelectionItem.IsSupported)
                                        child.Patterns.SelectionItem.Pattern.Select();
                                    else
                                        child.Click();

                                    Thread.Sleep(100);
                                    selected = true;
                                }
                                catch
                                {
                                    // Last resort: click the stale reference.
                                    try
                                    {
                                        child.Click();
                                        selected = true;
                                    }
                                    catch { /* best effort */ }
                                }
                            }

                            // Final fallback: if the ComboBox supports the Value pattern
                            // and no prior strategy confirmed success, write the selected
                            // item's name directly into the displayed value. This covers
                            // custom / owner-drawn combos where neither Click() nor
                            // SelectionItem.Select() update the text.
                            if (!selected && isComboBox && !string.IsNullOrEmpty(capturedChildInfo.Name))
                            {
                                try
                                {
                                    var valuePattern = element.Patterns.Value.PatternOrDefault;
                                    if (valuePattern != null && !valuePattern.IsReadOnly.Value)
                                        valuePattern.SetValue(capturedChildInfo.Name);
                                }
                                catch { /* best effort — Value pattern may not be writable */ }
                            }

                            // Ensure the dropdown is collapsed regardless of the path taken.
                            try { element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse(); }
                            catch { /* best effort */ }

                            // Record against the parent element (e.g. the ComboBox) so the
                            // locator targets the container, not the transient list item.
                            // Use AutomationId or ClassName to describe the parent combo box
                            // because its Name may reflect the currently displayed value and
                            // would be confusing in the recorded description.
                            var itemLabel = capturedChildInfo.Name ?? capturedChildInfo.AutomationId ?? "(item)";
                            string parentLabel;
                            if (!string.IsNullOrEmpty(elementInfo?.AutomationId))
                                parentLabel = elementInfo.AutomationId;
                            else if (!string.IsNullOrEmpty(elementInfo?.ClassName))
                                parentLabel = elementInfo.ClassName;
                            else
                                parentLabel = ElementInfo.GetLabel(elementInfo);
                            _service.AddAction(new RecordedAction
                            {
                                ActionType = ActionType.Select,
                                Mode = RecordingMode.Assistive,
                                Element = elementInfo,
                                Value = capturedChildInfo.Name,
                                Description = $"Select '{itemLabel}' from {parentLabel}"
                            });
                            UpdateStatusAfterAction($"Select '{itemLabel}' on [{elementInfo?.ControlType}] {parentLabel}");
                        };
                        childrenMenu.DropDownItems.Add(childItem);
                    }
                    if (children.Length > MaxChildrenToDisplay)
                        childrenMenu.DropDownItems.Add(new ToolStripMenuItem($"… and {children.Length - MaxChildrenToDisplay} more") { Enabled = false });
                    menu.Items.Add(childrenMenu);
                }
            }
            catch { /* best effort */ }
        }

        // ── Popup Windows ▶ — sibling child Window elements of the current window ─────
        // This section is always shown when child Window siblings exist (e.g. a "Working
        // Environment" dialog that opened inside the parent "Login" window after the user
        // clicked OK).  It offers a "Switch Window" action for each sibling window and a
        // "Buttons ▶" sub-flyout so the user can directly interact with its buttons.
        if (windowAncestor != null && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                var siblingWindows = windowAncestor.FindAllChildren(cf.ByControlType(ControlType.Window));
                if (siblingWindows.Length > 0)
                {
                    menu.Items.Add(new ToolStripSeparator());
                    var popupWindowsFlyout = new ToolStripMenuItem("Popup Windows ▶");

                    foreach (var sibWin in siblingWindows.Take(MaxChildrenToDisplay))
                    {
                        var sibWinInfo = BuildElementInfo(sibWin);
                        var sibWinTitle = sibWinInfo.Name ?? sibWinInfo.AutomationId ?? "(unnamed)";
                        var sibWinItem = new ToolStripMenuItem($"[Window]  {sibWinTitle}");

                        var capturedSibWinInfo = sibWinInfo;
                        var capturedSibWinTitle = sibWinTitle;
                        var capturedSibWin = sibWin;
                        sibWinItem.Click += (_, _) =>
                        {
                            _service.AddAction(new RecordedAction
                            {
                                ActionType = ActionType.SwitchWindow,
                                Mode = RecordingMode.Assistive,
                                Element = capturedSibWinInfo,
                                Value = capturedSibWinTitle,
                                Description = $"Switch window to '{capturedSibWinTitle}'"
                            });
                            UpdateStatusAfterAction($"Switch Window [Window] {capturedSibWinTitle}");
                        };

                        // Add "Buttons ▶" sub-flyout for the child window's buttons.
                        try
                        {
                            var cf2 = _automation.ConditionFactory;
                            var sibButtons = sibWin.FindAllDescendants(cf2.ByControlType(ControlType.Button));
                            if (sibButtons.Length > 0)
                            {
                                var sibBtnFlyout = new ToolStripMenuItem("Buttons ▶");
                                foreach (var btn in sibButtons.Take(MaxChildrenToDisplay))
                                {
                                    var btnInfo = BuildElementInfo(btn);
                                    var btnMenuLabel =
                                        $"[Button]  {btnInfo.Name ?? btnInfo.AutomationId ?? "(unnamed)"}";
                                    sibBtnFlyout.DropDownItems.Add(
                                        CreateWindowButtonMenuItem(btnMenuLabel, capturedSibWin, btnInfo, btn));
                                }

                                if (sibButtons.Length > MaxChildrenToDisplay)
                                    sibBtnFlyout.DropDownItems.Add(
                                        new ToolStripMenuItem(
                                            $"… and {sibButtons.Length - MaxChildrenToDisplay} more")
                                        { Enabled = false });

                                sibWinItem.DropDownItems.Add(sibBtnFlyout);
                            }
                        }
                        catch { /* best effort */ }

                        popupWindowsFlyout.DropDownItems.Add(sibWinItem);
                    }

                    if (siblingWindows.Length > MaxChildrenToDisplay)
                        popupWindowsFlyout.DropDownItems.Add(
                            new ToolStripMenuItem(
                                $"… and {siblingWindows.Length - MaxChildrenToDisplay} more")
                            { Enabled = false });

                    menu.Items.Add(popupWindowsFlyout);
                }
            }
            catch { /* best effort */ }
        }

        AddCloseItem(menu);
        // Show using screen coordinates so the menu appears at the correct position
        // regardless of the overlay window's location, and ensure the overlay is
        // topmost so the menu renders above the target application's windows.
        EnsureOverlayVisible();
        menu.Show(pt);
    }

    // ── Assistive mode: Ctrl+Right Click window context menu ─────────────────

    /// <summary>
    /// Guard wrapper for <see cref="ShowWindowContextMenuCore"/>. Mirrors the pattern
    /// used by <see cref="ShowAssistiveContextMenu"/>.
    /// </summary>
    private void ShowWindowContextMenu(System.Drawing.Point pt)
    {
        if (_menuOpen)
        {
            _statusLabel.Text = "Window menu blocked: _menuOpen=true. Resetting.";
            _logger.LogWarning("ShowWindowContextMenu called while _menuOpen=true; resetting flag to allow new menu");
            _menuOpen = false;
        }
        if (_automation == null)
        {
            _statusLabel.Text = "Window menu blocked: _automation=null";
            return;
        }

        _menuOpen = true;

        // Safety fallback: reset _menuOpen after 10 s in case the menu's Closed event
        // never fires (e.g. edge-case focus or disposal race).
        StartMenuSafetyTimer();

        try
        {
            _statusLabel.Text = $"Opening window/popup menu at {pt.X},{pt.Y}";
            ShowWindowContextMenuCore(pt);
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "Window/popup menu failed: " + ex.Message;
            _logger.LogWarning(ex, "Failed to show window context menu at {Pt}", pt);
            _menuOpen = false;
        }
    }

    /// <summary>
    /// Builds and shows a context menu that targets the <em>window</em> under the cursor
    /// rather than the individual UI element. Triggered by Ctrl+Right Click in Assistive mode.
    ///
    /// <para>The menu always contains:</para>
    /// <list type="bullet">
    ///   <item>A read-only header showing the window title.</item>
    ///   <item><b>Switch Window</b> — records <see cref="ActionType.SwitchWindow"/>.</item>
    ///   <item><b>Maximize / Minimize / Close Window</b> — when the window supports the WindowPattern.</item>
    ///   <item><b>Window Buttons ▶</b> — flyout listing all Button descendants of the window,
    ///         e.g. OK, Cancel, Yes, No. Each records <see cref="ActionType.Click"/>.</item>
    /// </list>
    /// </summary>
    private void ShowWindowContextMenuCore(System.Drawing.Point pt)
    {
        if (_automation == null) return;

        bool isPopupMode = false;
        AutomationElement? windowElement = null;

        // ── Priority 0: use recently detected popup cache ────────────────────
        // This avoids waiting for mouse movement / FromPoint refresh.
        if (_lastDetectedPopupWindow != null &&
            _lastDetectedPopupHwnd != IntPtr.Zero &&
            DateTime.UtcNow - _lastPopupDetectionUtc < TimeSpan.FromSeconds(5))
        {
            try
            {
                if (!_lastDetectedPopupWindow.IsOffscreen)
                {
                    windowElement = _lastDetectedPopupWindow;
                    isPopupMode = true;
                }
            }
            catch
            {
                _lastDetectedPopupWindow = null;
                _lastDetectedPopupHwnd = IntPtr.Zero;
            }
        }

        // ── Priority 1: use GetForegroundWindow to detect popup/dialog ──────
        // This is more reliable than FromPoint(pt) which can return TitleBar,
        // MenuBar, or MenuItem "System" when a modal popup has grabbed focus.
        var foregroundHwnd = GetForegroundWindow();
        var mainHwnd = _service.GetApplicationMainWindowHandle();

        _logger.LogInformation(
            "Assistive window menu: foreground={ForegroundHwnd}, main={MainHwnd}",
            foregroundHwnd,
            mainHwnd);

        if (windowElement == null &&
            foregroundHwnd != IntPtr.Zero &&
            mainHwnd != IntPtr.Zero &&
            foregroundHwnd != mainHwnd)
        {
            // The foreground window differs from the application's main window:
            // treat it as a popup / dialog.
            try
            {
                var foregroundElement = _automation.FromHandle(foregroundHwnd);
                if (foregroundElement != null)
                {
                    // Walk up to the Window ancestor if FromHandle returned a child.
                    var cur = foregroundElement;
                    for (int d = 0; d < MaxWindowSearchDepth && cur != null; d++)
                    {
                        if (cur.ControlType == ControlType.Window)
                        {
                            windowElement = cur;
                            break;
                        }
                        cur = cur.Parent;
                    }
                    windowElement ??= foregroundElement;
                    isPopupMode = true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not resolve foreground window HWND 0x{Hwnd:X}", foregroundHwnd);
            }
        }

        // ── Priority 2: fall back to FromPoint if no popup was detected ─────
        if (windowElement == null)
        {
            AutomationElement? element = null;
            try { element = _automation.FromPoint(pt); }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not get element at point {Pt} for window context menu", pt);
            }

            // Walk up the UIA tree to find the nearest Window ancestor (or self).
            try
            {
                var current = element;
                for (int _ = 0; _ < MaxWindowSearchDepth && current != null; _++)
                {
                    if (current.ControlType == ControlType.Window)
                    {
                        windowElement = current;
                        break;
                    }
                    current = current.Parent;
                }
            }
            catch { /* best effort */ }

            // Fall back to the element itself if no Window parent was found.
            windowElement ??= element;

            // Priority 2a: if FromPoint returned an element inside the window chrome
            // (TitleBar / MenuBar / MenuItem / System menu), the cursor may be near a
            // popup window border.  Try to locate a child popup window as a fallback.
            if (element != null && AssistivePopupResolver.IsInsideChromeOrSystemMenu(element) && windowElement != null)
            {
                var childPopup = AssistivePopupResolver.DetectChildPopupWindow(windowElement, _automation);
                if (childPopup != null)
                {
                    windowElement = childPopup;
                    isPopupMode = true;
                }
            }

            // Priority 2b: prefer the innermost child Window that contains the cursor.
            // This handles embedded/child dialogs (e.g. a "Working Environment" popup that
            // appears as a Window child inside the main application window after clicking OK).
            // FromPoint may return an element that belongs to the outer frame when the cursor
            // is near the parent window's border rather than strictly inside the child window.
            if (!isPopupMode && windowElement != null)
            {
                try
                {
                    var cf = _automation.ConditionFactory;
                    var childWindows = windowElement.FindAllChildren(cf.ByControlType(ControlType.Window));
                    var innerWindow = childWindows.FirstOrDefault(w =>
                    {
                        try
                        {
                            var bounds = w.BoundingRectangle;
                            return !bounds.IsEmpty && bounds.Contains(pt.X, pt.Y);
                        }
                        catch { return false; }
                    });
                    if (innerWindow != null)
                        windowElement = innerWindow;
                }
                catch { /* best effort */ }
            }
        }

        var windowInfo = windowElement != null ? BuildElementInfo(windowElement) : null;
        var windowTitle = windowInfo?.Name ?? string.Empty;

        // Capture the HWND before the menu appears (mirrors the existing Right Click handler).
        IntPtr capturedHwnd = IntPtr.Zero;
        if (windowElement != null)
        {
            try
            {
                var hwnd = windowElement.Properties.NativeWindowHandle.Value;
                if (hwnd != IntPtr.Zero)
                {
                    var root = GetAncestor(hwnd, GA_ROOT);
                    capturedHwnd = root != IntPtr.Zero ? root : hwnd;
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not capture HWND for window at {Pt}", pt);
            }
        }

        // ── Build the context menu ──────────────────────────────────────────
        var menu = new NoActivateContextMenuStrip { ShowImageMargin = false };
        menu.Font = new Font("Segoe UI", 10f);

        menu.Closed += (_, _) =>
        {
            _menuOpen = false;
            ReapplyClickThroughStyle();
        };

        // ── Popup-mode: focused header + popup-specific action items ─────────
        if (isPopupMode && windowElement != null)
        {
            var popupHandle = AssistivePopupResolver.SafeWindowHandle(windowElement);
            var handleHex = popupHandle != IntPtr.Zero
                ? $"0x{popupHandle.ToInt64():X}"
                : "(unknown handle)";

            // Header (not selectable)
            var popupHeader = new ToolStripMenuItem($"[Popup]  {windowTitle}")
            {
                Enabled = false,
                ForeColor = Color.DimGray
            };
            menu.Items.Add(popupHeader);

            var handleHeader = new ToolStripMenuItem($"Handle: {handleHex}")
            {
                Enabled = false,
                ForeColor = Color.DimGray
            };
            menu.Items.Add(handleHeader);
            menu.Items.Add(new ToolStripSeparator());

            // Click OK — find the OK button and invoke it.
            var capturedPopupWindow = windowElement;
            var capturedPopupInfo = windowInfo;
            var clickOkItem = new ToolStripMenuItem("Click OK");
            clickOkItem.Click += (_, _) =>
            {
                if (_automation == null) return;

                if (capturedHwnd != IntPtr.Zero)
                    SetForegroundWindow(capturedHwnd);

                try
                {
                    var cf = _automation.ConditionFactory;
                    var okBtn = AssistivePopupResolver.FindOkButton(capturedPopupWindow.AsWindow()!, cf);
                    if (okBtn != null)
                    {
                        var okInfo = BuildElementInfo(okBtn);
                        AssistivePopupResolver.InvokeOrClick(okBtn);
                        _service.AddAction(new RecordedAction
                        {
                            ActionType = ActionType.Click,
                            Mode = RecordingMode.Assistive,
                            Element = okInfo,
                            Description = $"Click OK on '{windowTitle}'"
                        });
                        UpdateStatusAfterAction($"Click OK on [Popup] {windowTitle}");
                        StartPopupProbeAfterAction();
                    }
                    else
                    {
                        System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                        _service.AddAction(new RecordedAction
                        {
                            ActionType = ActionType.Click,
                            Mode = RecordingMode.Assistive,
                            Element = capturedPopupInfo,
                            Description = $"Press Enter on '{windowTitle}' (OK button not found)"
                        });
                        UpdateStatusAfterAction($"Press Enter on [Popup] {windowTitle}");
                        StartPopupProbeAfterAction();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Click OK failed for popup '{Title}'", windowTitle);
                }
            };
            menu.Items.Add(clickOkItem);

            // Press Enter
            var pressEnterItem = new ToolStripMenuItem("Press ↵ Enter");
            pressEnterItem.Click += (_, _) =>
            {
                if (capturedHwnd != IntPtr.Zero)
                    SetForegroundWindow(capturedHwnd);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.Type,
                    Mode = RecordingMode.Assistive,
                    Element = capturedPopupInfo,
                    Value = "{ENTER}",
                    Description = $"Press Enter on '{windowTitle}'"
                });
                UpdateStatusAfterAction($"Press Enter on [Popup] {windowTitle}");
            };
            menu.Items.Add(pressEnterItem);

            // Press Esc
            var pressEscItem = new ToolStripMenuItem("Press Esc");
            pressEscItem.Click += (_, _) =>
            {
                if (capturedHwnd != IntPtr.Zero)
                    SetForegroundWindow(capturedHwnd);
                System.Windows.Forms.SendKeys.SendWait("{ESC}");
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.Type,
                    Mode = RecordingMode.Assistive,
                    Element = capturedPopupInfo,
                    Value = "{ESC}",
                    Description = $"Press Esc on '{windowTitle}'"
                });
                UpdateStatusAfterAction($"Press Esc on [Popup] {windowTitle}");
            };
            menu.Items.Add(pressEscItem);

            menu.Items.Add(new ToolStripSeparator());

            // Close Popup
            var closePopupItem = new ToolStripMenuItem("Close Popup");
            closePopupItem.Click += (_, _) =>
            {
                try
                {
                    if (capturedPopupWindow.Patterns.Window.IsSupported)
                        capturedPopupWindow.Patterns.Window.Pattern.Close();
                    else
                    {
                        if (capturedHwnd != IntPtr.Zero)
                            SetForegroundWindow(capturedHwnd);
                        System.Windows.Forms.SendKeys.SendWait("%{F4}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Close Popup failed for '{Title}'", windowTitle);
                }

                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.CloseWindow,
                    Mode = RecordingMode.Assistive,
                    Element = capturedPopupInfo,
                    Description = $"Close popup '{windowTitle}'"
                });
                UpdateStatusAfterAction($"Close Popup [Window] {windowTitle}");
            };
            menu.Items.Add(closePopupItem);

            // Inspect Popup Controls — flyout of all popup elements
            if (_automation != null)
            {
                try
                {
                    var cf = _automation.ConditionFactory;
                    var allElements = capturedPopupWindow
                        .FindAllDescendants()
                        .Where(e =>
                        {
                            try { return !AssistivePopupResolver.IsInsideChromeOrSystemMenu(e); }
                            catch { return false; }
                        })
                        .ToList();

                    if (allElements.Count > 0)
                    {
                        var inspectFlyout = new ToolStripMenuItem("Inspect Popup Controls ▶");

                        foreach (var e in allElements.Take(MaxChildrenToDisplay))
                        {
                            var eInfo = BuildElementInfo(e);
                            var eLabel = $"[{eInfo.ControlType}]  {eInfo.Name ?? eInfo.AutomationId ?? "(unnamed)"}";
                            var eItem = new ToolStripMenuItem(eLabel) { Enabled = false };
                            inspectFlyout.DropDownItems.Add(eItem);
                        }

                        if (allElements.Count > MaxChildrenToDisplay)
                            inspectFlyout.DropDownItems.Add(
                                new ToolStripMenuItem(
                                    $"… and {allElements.Count - MaxChildrenToDisplay} more")
                                { Enabled = false });

                        menu.Items.Add(inspectFlyout);
                    }
                }
                catch { /* best effort */ }
            }

            menu.Items.Add(new ToolStripSeparator());

            // Make This Current Window — records a SwitchWindow action so the automation
            // script will target this popup window on subsequent operations.
            var makeCurrentItem = new ToolStripMenuItem("Make This Current Window");
            makeCurrentItem.Click += (_, _) =>
            {
                var label = ElementInfo.GetLabel(capturedPopupInfo);
                _service.AddAction(new RecordedAction
                {
                    ActionType = ActionType.SwitchWindow,
                    Mode = RecordingMode.Assistive,
                    Element = capturedPopupInfo,
                    Value = windowTitle,
                    Description = $"Switch window to '{windowTitle}'"
                });
                UpdateStatusAfterAction($"Make Current Window [Window] {label}");
            };
            menu.Items.Add(makeCurrentItem);

            AddCloseItem(menu);
            EnsureOverlayVisible();
            menu.Show(pt);
            return;
        }

        // ── Normal-mode: existing window context menu ─────────────────────────

        // Header (not selectable)
        var headerName = windowInfo?.Name ?? windowInfo?.AutomationId ?? "(no name)";
        var header = new ToolStripMenuItem($"[Window]  {headerName}")
        {
            Enabled = false,
            ForeColor = Color.DimGray
        };
        menu.Items.Add(header);
        menu.Items.Add(new ToolStripSeparator());

        // Switch Window — always available
        var switchItem = new ToolStripMenuItem("Switch Window");
        switchItem.Click += (_, _) =>
        {
            var label = ElementInfo.GetLabel(windowInfo);
            _service.AddAction(new RecordedAction
            {
                ActionType = ActionType.SwitchWindow,
                Mode = RecordingMode.Assistive,
                Element = windowInfo,
                Value = windowTitle,
                Description = $"Switch window to '{windowTitle}'"
            });
            UpdateStatusAfterAction($"Switch Window [{windowInfo?.ControlType}] {label}");
        };
        menu.Items.Add(switchItem);

        // Window pattern operations (Maximize / Minimize / Close Window)
        if (windowElement != null && windowElement.Patterns.Window.IsSupported)
        {
            AddActionItem(menu, "Maximize", windowElement, windowInfo, ActionType.Maximize,
                () => windowElement.Patterns.Window.Pattern.SetWindowVisualState(
                    FlaUI.Core.Definitions.WindowVisualState.Maximized));
            AddActionItem(menu, "Minimize", windowElement, windowInfo, ActionType.Minimize,
                () => windowElement.Patterns.Window.Pattern.SetWindowVisualState(
                    FlaUI.Core.Definitions.WindowVisualState.Minimized));
            AddActionItem(menu, "Close Window", windowElement, windowInfo, ActionType.CloseWindow,
                () => windowElement.Patterns.Window.Pattern.Close());
        }

        // Window Buttons flyout — Button descendants of the window
        if (windowElement != null && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                var buttons = windowElement.FindAllDescendants(cf.ByControlType(ControlType.Button));

                if (buttons.Length > 0)
                {
                    menu.Items.Add(new ToolStripSeparator());
                    var buttonsFlyout = new ToolStripMenuItem("Window Buttons ▶");

                    foreach (var btn in buttons.Take(MaxChildrenToDisplay))
                    {
                        var btnInfo = BuildElementInfo(btn);
                        var btnMenuLabel = $"[Button]  {btnInfo.Name ?? btnInfo.AutomationId ?? "(unnamed)"}";
                        var btnItem = CreateWindowButtonMenuItem(
                            btnMenuLabel, windowElement, btnInfo, btn, capturedHwnd);
                        buttonsFlyout.DropDownItems.Add(btnItem);
                    }

                    if (buttons.Length > MaxChildrenToDisplay)
                        buttonsFlyout.DropDownItems.Add(
                            new ToolStripMenuItem($"… and {buttons.Length - MaxChildrenToDisplay} more") { Enabled = false });

                    menu.Items.Add(buttonsFlyout);
                }
            }
            catch { /* best effort */ }
        }

        // Child Windows flyout — Window children of the located window (e.g. embedded
        // dialogs such as "Working Environment" that appear as child Window UIA elements
        // inside the parent window after a user action like clicking OK).
        if (windowElement != null && _automation != null)
        {
            try
            {
                var cf = _automation.ConditionFactory;
                var childWindows = windowElement.FindAllChildren(cf.ByControlType(ControlType.Window));

                if (childWindows.Length > 0)
                {
                    menu.Items.Add(new ToolStripSeparator());
                    var childWindowsFlyout = new ToolStripMenuItem("Child Windows ▶");

                    foreach (var childWin in childWindows.Take(MaxChildrenToDisplay))
                    {
                        var childWinInfo = BuildElementInfo(childWin);
                        var childWinTitle = childWinInfo.Name ?? childWinInfo.AutomationId ?? "(unnamed)";
                        var childWinItem = new ToolStripMenuItem($"[Window]  {childWinTitle}");

                        // Clicking the child window entry records a SwitchWindow action.
                        var capturedChildWinInfo = childWinInfo;
                        var capturedChildWinTitle = childWinTitle;
                        var capturedChildWin = childWin;
                        childWinItem.Click += (_, _) =>
                        {
                            _service.AddAction(new RecordedAction
                            {
                                ActionType = ActionType.SwitchWindow,
                                Mode = RecordingMode.Assistive,
                                Element = capturedChildWinInfo,
                                Value = capturedChildWinTitle,
                                Description = $"Switch window to '{capturedChildWinTitle}'"
                            });
                            UpdateStatusAfterAction($"Switch Window [Window] {capturedChildWinTitle}");
                        };

                        // Build a "Buttons ▶" sub-flyout for this child window's Button descendants.
                        try
                        {
                            var cf2 = _automation.ConditionFactory;
                            var childButtons = childWin.FindAllDescendants(cf2.ByControlType(ControlType.Button));
                            if (childButtons.Length > 0)
                            {
                                var btnSubFlyout = new ToolStripMenuItem("Buttons ▶");
                                foreach (var btn in childButtons.Take(MaxChildrenToDisplay))
                                {
                                    var btnInfo = BuildElementInfo(btn);
                                    var btnMenuLabel = $"[Button]  {btnInfo.Name ?? btnInfo.AutomationId ?? "(unnamed)"}";
                                    btnSubFlyout.DropDownItems.Add(
                                        CreateWindowButtonMenuItem(btnMenuLabel, capturedChildWin, btnInfo, btn));
                                }

                                if (childButtons.Length > MaxChildrenToDisplay)
                                    btnSubFlyout.DropDownItems.Add(
                                        new ToolStripMenuItem(
                                            $"… and {childButtons.Length - MaxChildrenToDisplay} more")
                                        { Enabled = false });

                                childWinItem.DropDownItems.Add(btnSubFlyout);
                            }
                        }
                        catch { /* best effort */ }

                        childWindowsFlyout.DropDownItems.Add(childWinItem);
                    }

                    if (childWindows.Length > MaxChildrenToDisplay)
                        childWindowsFlyout.DropDownItems.Add(
                            new ToolStripMenuItem(
                                $"… and {childWindows.Length - MaxChildrenToDisplay} more")
                            { Enabled = false });

                    menu.Items.Add(childWindowsFlyout);
                }
            }
            catch { /* best effort */ }
        }

        AddCloseItem(menu);
        EnsureOverlayVisible();
        menu.Show(pt);
    }

    /// <summary>
    /// Appends a separator and a "✕  Close" item that lets the user dismiss the context
    /// menu without recording any action.
    /// </summary>
    private static void AddCloseItem(ContextMenuStrip menu)
    {
        menu.Items.Add(new ToolStripSeparator());
        var closeItem = new ToolStripMenuItem("✕  Close")
        {
            ForeColor = Color.Gray
        };
        closeItem.Click += (_, _) => menu.Close(ToolStripDropDownCloseReason.ItemClicked);
        menu.Items.Add(closeItem);
    }

    /// <summary>
    /// Returns <c>true</c> when an active popup or child dialog window is detected,
    /// updating the popup cache fields on success.
    /// Uses <see cref="DetectActivePopupWindow"/> for stronger detection that covers
    /// child windows of the main window (not only foreground HWND comparison).
    /// </summary>
    private bool IsPopupCurrentlyActive()
    {
        try
        {
            var popup = DetectActivePopupWindow();

            if (popup != null)
            {
                _lastDetectedPopupWindow = popup;
                _lastDetectedPopupHwnd = AssistivePopupResolver.SafeWindowHandle(popup);
                _lastPopupDetectionUtc = DateTime.UtcNow;
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "IsPopupCurrentlyActive check failed; assuming no popup");
            return false;
        }
    }

    /// <summary>
    /// Attempts to locate an active popup or child dialog window using three strategies:
    /// <list type="number">
    ///   <item>Foreground window differs from the main application window.</item>
    ///   <item>A visible, named child Window exists under the main application window
    ///         (covers popups that appear as child windows of the main HWND).</item>
    ///   <item>The element under the current cursor position belongs to a window other
    ///         than the main application window.</item>
    /// </list>
    /// Returns the first matched <see cref="AutomationElement"/>, or <c>null</c> if no
    /// popup is found.
    /// </summary>
    private AutomationElement? DetectActivePopupWindow()
    {
        if (_automation == null)
            return null;

        var mainHwnd = _service.GetApplicationMainWindowHandle();
        var foregroundHwnd = GetForegroundWindow();

        // 1. Strongest signal: foreground window is different from main window.
        if (foregroundHwnd != IntPtr.Zero &&
            mainHwnd != IntPtr.Zero &&
            foregroundHwnd != mainHwnd)
        {
            try
            {
                var foregroundElement = _automation.FromHandle(foregroundHwnd);

                if (foregroundElement != null)
                {
                    var win = FindWindowAncestorOrSelf(foregroundElement);

                    if (win != null)
                    {
                        _logger.LogInformation(
                            "Popup detected: name={Name}, hwnd=0x{Hwnd:X}",
                            win.Name,
                            AssistivePopupResolver.SafeWindowHandle(win).ToInt64());
                        return win;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not resolve foreground popup hwnd 0x{Hwnd:X}", foregroundHwnd);
            }
        }

        // 2. Important for this app:
        // Popup appears as a child Window under LoginForm, while foreground can still be main window.
        try
        {
            AutomationElement? mainWindow = null;

            if (mainHwnd != IntPtr.Zero)
                mainWindow = _automation.FromHandle(mainHwnd);

            if (mainWindow != null)
            {
                var childPopup = AssistivePopupResolver.DetectChildPopupWindow(mainWindow, _automation);

                if (childPopup != null)
                {
                    _logger.LogInformation(
                        "Popup detected: name={Name}, hwnd=0x{Hwnd:X}",
                        childPopup.Name,
                        AssistivePopupResolver.SafeWindowHandle(childPopup).ToInt64());
                    return childPopup;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Could not detect child popup under main window");
        }

        // 3. Last fallback:
        // If current mouse point is inside a child Window, treat that child as popup.
        try
        {
            var pt = Cursor.Position;
            var element = _automation.FromPoint(pt);

            if (element != null)
            {
                var win = FindWindowAncestorOrSelf(element);

                if (win != null)
                {
                    var winHwnd = AssistivePopupResolver.SafeWindowHandle(win);

                    if (winHwnd != IntPtr.Zero &&
                        mainHwnd != IntPtr.Zero &&
                        winHwnd != mainHwnd)
                    {
                        _logger.LogInformation(
                            "Popup detected: name={Name}, hwnd=0x{Hwnd:X}",
                            win.Name,
                            winHwnd.ToInt64());
                        return win;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Could not detect popup from cursor point");
        }

        return null;
    }

    /// <summary>
    /// Walks up the UIA ancestor chain looking for a <see cref="ControlType.Window"/>
    /// element, returning the first one found or <c>null</c> if none is found within
    /// <see cref="MaxWindowSearchDepth"/> steps.
    /// </summary>
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

            return element.ControlType == ControlType.Window ? element : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Starts a short-lived probe timer that polls for a new popup window for up to
    /// 2 seconds after an assistive click action that may open a popup.
    /// Updates the overlay status label and popup cache when a popup is found.
    /// </summary>
    private void StartPopupProbeAfterAction()
    {
        if (_automation == null)
            return;

        var attempts = 0;
        var probeTimer = new System.Windows.Forms.Timer { Interval = 100 };

        probeTimer.Tick += (_, _) =>
        {
            attempts++;

            try
            {
                var popup = DetectActivePopupWindow();

                if (popup != null)
                {
                    _lastDetectedPopupWindow = popup;
                    _lastDetectedPopupHwnd = AssistivePopupResolver.SafeWindowHandle(popup);
                    _lastPopupDetectionUtc = DateTime.UtcNow;

                    var popupName = popup.Name ?? popup.AutomationId ?? "popup";
                    _statusLabel.Text = $"● [Popup] {popupName}  Right-click = Popup Actions";

                    probeTimer.Stop();
                    probeTimer.Dispose();
                    return;
                }
            }
            catch
            {
                // best effort
            }

            if (attempts >= 20) // 20 * 100 ms = 2 seconds
            {
                probeTimer.Stop();
                probeTimer.Dispose();
            }
        };

        probeTimer.Start();
    }


    private void EnsureOverlayVisible()
    {
        TopMost = true;
        Activate();
        BringToFront();
    }

    /// <summary>
    /// Re-applies the WS_EX_TRANSPARENT | WS_EX_LAYERED extended styles so that the
    /// overlay stays click-through after a ContextMenuStrip is dismissed. Showing a
    /// popup may cause Windows to recalculate window styles on the owner form.
    /// </summary>
    private void ReapplyClickThroughStyle()
    {
        if (!IsHandleCreated || IsDisposed) return;

        // Cancel the safety timer — the menu's Closed event fired normally so the
        // fallback is no longer needed.  Do this before touching window styles so
        // any exception in SetWindowLong does not leave the timer running.
        _menuSafetyTimer?.Stop();
        _menuSafetyTimer?.Dispose();
        _menuSafetyTimer = null;

        try
        {
            var style = GetWindowLong(Handle, GWL_EXSTYLE);
            SetWindowLong(Handle, GWL_EXSTYLE, style | WS_EX_TRANSPARENT | WS_EX_LAYERED);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to reapply click-through styles on recording overlay");
        }
    }

    /// <summary>
    /// Starts a 10-second one-shot timer that resets <see cref="_menuOpen"/> to
    /// <c>false</c> as a safety net in case a context menu's <c>Closed</c> event
    /// never fires (e.g. an edge-case focus or disposal race condition).
    /// The timer reference is stored in <see cref="_menuSafetyTimer"/> so that
    /// <see cref="ReapplyClickThroughStyle"/> can stop and dispose it when the
    /// menu closes normally, preventing the timer from outliving the menu.
    /// </summary>
    private void StartMenuSafetyTimer()
    {
        // Stop any previous safety timer in case StartMenuSafetyTimer is called again
        // before the previous menu fully closed.
        _menuSafetyTimer?.Stop();
        _menuSafetyTimer?.Dispose();

        const int MenuSafetyTimeoutMs = 10000;
        _menuSafetyTimer = new System.Windows.Forms.Timer { Interval = MenuSafetyTimeoutMs };
        _menuSafetyTimer.Tick += (_, _) =>
        {
            _menuSafetyTimer?.Stop();
            _menuSafetyTimer?.Dispose();
            _menuSafetyTimer = null;
            _menuOpen = false;
        };
        _menuSafetyTimer.Start();
    }


    private void AddActionItem(
        ContextMenuStrip menu,
        string label,
        AutomationElement? element,
        ElementInfo? info,
        ActionType actionType,
        Action perform)
    {
        var item = new ToolStripMenuItem(label);
        item.Click += (_, _) =>
        {
            try { perform(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Action {Action} failed", actionType); }

            _service.AddAction(new RecordedAction
            {
                ActionType = actionType,
                Mode = RecordingMode.Assistive,
                Element = info,
                Description = BuildDescription(label, info)
            });
            UpdateStatusAfterAction($"{label} on [{info?.ControlType}] {info?.Name ?? "(element)"}");

            // After a Click action, start a short probe to detect any popup that may
            // have opened as a result (e.g. clicking the login OK button).
            if (actionType == ActionType.Click)
                StartPopupProbeAfterAction();
        };
        menu.Items.Add(item);
    }

    private void AddQueryItem(
        ContextMenuStrip menu,
        string label,
        AutomationElement? element,
        ElementInfo? info,
        ActionType actionType,
        Func<bool> evaluate)
    {
        var item = new ToolStripMenuItem(label + "?");
        item.Click += (_, _) =>
        {
            bool result;
            try { result = evaluate(); }
            catch { result = false; }

            _service.AddAction(new RecordedAction
            {
                ActionType = actionType,
                Mode = RecordingMode.Assistive,
                Element = info,
                QueryResult = result,
                Description = $"{label} check on {ElementInfo.GetLabel(info)}: {result}"
            });
            UpdateStatusAfterAction($"{label}: {result}  │  [{info?.ControlType}] {info?.Name ?? "(element)"}");
        };
        menu.Items.Add(item);
    }

    /// <summary>
    /// Creates a <see cref="ToolStripMenuItem"/> whose click handler brings
    /// <paramref name="searchRoot"/> to the foreground, re-locates a fresh reference
    /// to the button by AutomationId / Name, invokes it, and records a
    /// <see cref="ActionType.Click"/> action.
    ///
    /// Used by the <em>Window Buttons ▶</em>, <em>Child Windows ▶</em>, and
    /// <em>Popup Windows ▶</em> menus which share identical button-invocation logic.
    /// </summary>
    /// <param name="menuLabel">Label shown in the menu.</param>
    /// <param name="searchRoot">Window element used to search for a fresh button reference.</param>
    /// <param name="buttonInfo">Captured element info for the button (AutomationId / Name used for re-location).</param>
    /// <param name="originalButton">Stale-fallback button reference when re-location fails.</param>
    /// <param name="preferredHwnd">
    ///   If non-zero, <c>SetForegroundWindow</c> is called with this HWND before the click
    ///   (faster than walking the UIA tree to the root window).
    ///   Pass <c>IntPtr.Zero</c> (default) to use <see cref="BringElementWindowToForeground"/>.
    /// </param>
    private ToolStripMenuItem CreateWindowButtonMenuItem(
        string menuLabel,
        AutomationElement searchRoot,
        ElementInfo buttonInfo,
        AutomationElement originalButton,
        IntPtr preferredHwnd = default)
    {
        var item = new ToolStripMenuItem(menuLabel);
        item.Click += (_, _) =>
        {
            // Bring the window to the foreground before firing the click so it lands
            // on the correct window rather than whatever currently has focus.
            if (preferredHwnd != IntPtr.Zero)
                SetForegroundWindow(preferredHwnd);
            else
                BringElementWindowToForeground(searchRoot);
            Thread.Sleep(WindowActivationDelayMs);

            try
            {
                // Re-locate a fresh button reference so that the click is not fired on a
                // stale element (the menu interaction may have shifted focus away from the
                // window, potentially invalidating the captured reference).
                AutomationElement? freshBtn = null;
                if (_automation != null)
                {
                    var cf = _automation.ConditionFactory;
                    if (!string.IsNullOrEmpty(buttonInfo.AutomationId))
                        freshBtn = searchRoot.FindFirstDescendant(cf.ByAutomationId(buttonInfo.AutomationId));
                    if (freshBtn == null && !string.IsNullOrEmpty(buttonInfo.Name))
                        freshBtn = searchRoot.FindFirstDescendant(
                            cf.ByControlType(ControlType.Button).And(cf.ByName(buttonInfo.Name)));
                }

                if (freshBtn != null)
                {
                    if (freshBtn.Patterns.Invoke.IsSupported)
                        freshBtn.Patterns.Invoke.Pattern.Invoke();
                    else
                        freshBtn.Click();
                }
                else
                {
                    // Fall back to the (possibly stale) original reference.
                    if (originalButton.Patterns.Invoke.IsSupported)
                        originalButton.Patterns.Invoke.Pattern.Invoke();
                    else
                        originalButton.Click();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Window button click failed for '{Label}'", buttonInfo.Name);
            }

            var buttonLabel = ElementInfo.GetLabel(buttonInfo);
            _service.AddAction(new RecordedAction
            {
                ActionType = ActionType.Click,
                Mode = RecordingMode.Assistive,
                Element = buttonInfo,
                Description = $"Click on {buttonLabel}"
            });
            UpdateStatusAfterAction($"Click [Button] {buttonInfo.Name ?? "(button)"}");
            StartPopupProbeAfterAction();
        };
        return item;
    }

    private void UpdateStatusAfterAction(string detail)
    {
        _statusLabel.Text = $"  ✓ Recorded: {detail}  │  Ctrl+S = Stop  ";
        // Revert to mode label after 2 s
        var timer = new System.Windows.Forms.Timer { Interval = 2000 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            if (_service.CurrentMode == RecordingMode.Assistive)
                _statusLabel.Text = "  Assistive ACTIVE  │  Right-click element  │  Ctrl+Right-click for window actions  │  Ctrl+P = Passive  │  Ctrl+S = Stop  ";
            else if (_service.CurrentMode == RecordingMode.Passive)
                _statusLabel.Text = "  Passive ACTIVE  │  Recording clicks & keys  │  Ctrl+A = Assistive  │  Ctrl+S = Stop  ";
        };
        timer.Start();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds a description for an interactive action, e.g. "Click on Login Button".
    /// </summary>
    private static string BuildDescription(string actionLabel, ElementInfo? info) =>
        $"{actionLabel} on {ElementInfo.GetLabel(info)}";

    internal static ElementInfo BuildElementInfo(AutomationElement element)
    {
        string ct;
        try { ct = element.ControlType.ToString(); }
        catch (Exception) { ct = string.Empty; }

        string? name = null;
        try { name = element.Name; } catch (Exception) { }

        string? automationId = null;
        try { automationId = element.AutomationId; } catch (Exception) { }

        string? className = null;
        try { className = element.ClassName; } catch (Exception) { }

        string? boundingRect = null;
        try { boundingRect = element.BoundingRectangle.ToString(); } catch (Exception) { }

        var info = new ElementInfo
        {
            Name = name,
            AutomationId = automationId,
            ClassName = className,
            ControlType = ct,
            BoundingRectangle = boundingRect
        };

        // Build the best possible XPath hint
        if (!string.IsNullOrEmpty(automationId))
            info.SuggestedXPath = $"//{ct}[@AutomationId='{automationId}']";
        else if (!string.IsNullOrEmpty(name))
            info.SuggestedXPath = $"//{ct}[@Name='{name}']";
        else if (!string.IsNullOrEmpty(className))
            info.SuggestedXPath = $"//{ct}[@ClassName='{className}']";

        return info;
    }

    private static bool IsSelectable(AutomationElement element) =>
        element.Patterns.SelectionItem.IsSupported ||
        element.ControlType == ControlType.ListItem ||
        element.ControlType == ControlType.MenuItem ||
        element.ControlType == ControlType.TreeItem ||
        element.ControlType == ControlType.TabItem ||
        element.ControlType == ControlType.RadioButton;

    private static bool IsContainer(ControlType ct) =>
        ct == ControlType.List ||
        ct == ControlType.Menu ||
        ct == ControlType.MenuBar ||
        ct == ControlType.Tree ||
        ct == ControlType.Tab ||
        ct == ControlType.ToolBar;

    /// <summary>
    /// When <paramref name="element"/> is a container (e.g. MenuBar, Menu, ToolBar) and
    /// <c>automation.FromPoint</c> returned the parent instead of the child that is
    /// visually under the cursor, drill down one level by matching child bounding rectangles.
    /// Returns the most specific child found, or the original element when none matches.
    /// </summary>
    internal static AutomationElement DrillDownToElementAtPoint(
        AutomationElement element, System.Drawing.Point pt)
    {
        if (!IsContainer(element.ControlType))
            return element;
        try
        {
            var children = element.FindAllChildren();
            var child = children.FirstOrDefault(c =>
            {
                try
                {
                    var r = c.BoundingRectangle;
                    return !r.IsEmpty && r.Contains(pt.X, pt.Y);
                }
                catch { return false; }
            });
            if (child != null)
                return child;
        }
        catch { /* best effort */ }
        return element;
    }

    /// <summary>
    /// Brings the top-level window that contains <paramref name="element"/> to the foreground
    /// using Win32 <c>SetForegroundWindow</c>.  This is more reliable than UIA
    /// <c>element.Focus()</c> when the recording context menu dismissal allows another
    /// application (e.g. IntelliJ) to reclaim the foreground before a physical click fires.
    ///
    /// The driver process is allowed to call <c>SetForegroundWindow</c> while it still
    /// holds the foreground lock (immediately after the context-menu closes).
    ///
    /// For virtual-element frameworks (e.g. Java Swing via Java Access Bridge) where
    /// <c>NativeWindowHandle</c> is zero, falls back to the process's main window handle.
    /// </summary>
    private static void BringElementWindowToForeground(AutomationElement? element)
    {
        if (element == null) return;

        // Primary path: resolve the HWND from the element.
        bool activated = false;
        try
        {
            var hwnd = element.Properties.NativeWindowHandle.Value;
            if (hwnd != IntPtr.Zero)
            {
                // Walk up to the root top-level window so we never activate a child HWND
                // (child HWNDs do not become foreground windows; only top-level ones do).
                var root = GetAncestor(hwnd, GA_ROOT);
                if (root != IntPtr.Zero)
                    activated = SetForegroundWindow(root);
            }
        }
        catch { /* best effort */ }

        if (activated) return;

        // Fallback for virtual-element frameworks (e.g. Java Swing via Java Access Bridge)
        // where UIA elements do not have their own HWND, or when SetForegroundWindow failed.
        // Resolve the owning process and bring its main window to the foreground.
        try
        {
            var pid = element.Properties.ProcessId.Value;
            var proc = System.Diagnostics.Process.GetProcessById(pid);
            var mainHwnd = proc.MainWindowHandle;
            if (mainHwnd != IntPtr.Zero)
                SetForegroundWindow(mainHwnd);
        }
        catch { /* best effort — never fail a click because of a focus hint */ }
    }

    /// <summary>
    /// Escapes characters that have special meaning in <see cref="System.Windows.Forms.SendKeys"/>
    /// (+, ^, %, ~, (, ), {, }, [, ]) by wrapping each one in curly braces.
    /// </summary>
    private static string EscapeForSendKeys(string text)
    {
        // Characters that SendKeys interprets as special commands
        const string SpecialChars = "+^%~(){}[]";
        var sb = new System.Text.StringBuilder(text.Length * 2);
        foreach (var ch in text)
        {
            if (SpecialChars.Contains(ch, StringComparison.Ordinal))
            {
                sb.Append('{');
                sb.Append(ch);
                sb.Append('}');
            }
            else
            {
                sb.Append(ch);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Shows a small modal input dialog prompting the user to enter text to type
    /// into the focused element. Returns the entered text, or null if cancelled.
    /// </summary>
    private static string? ShowTypePrompt(string elementLabel)
    {
        string? result = null;

        using var form = new Form
        {
            Text = "Type into element",
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            Width = 380,
            Height = 145,
            MaximizeBox = false,
            MinimizeBox = false,
            TopMost = true
        };

        var label = new WinLabel
        {
            Text = $"Text to type into \"{elementLabel}\":",
            Left = 12,
            Top = 12,
            Width = 350,
            AutoSize = true
        };

        var textBox = new System.Windows.Forms.TextBox
        {
            Left = 12,
            Top = 35,
            Width = 340
        };

        var okBtn = new System.Windows.Forms.Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Left = 192,
            Top = 68,
            Width = 80
        };

        var cancelBtn = new System.Windows.Forms.Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = 278,
            Top = 68,
            Width = 80
        };

        form.AcceptButton = okBtn;
        form.CancelButton = cancelBtn;
        form.Controls.AddRange([label, textBox, okBtn, cancelBtn]);

        if (form.ShowDialog() == DialogResult.OK)
            result = textBox.Text;

        return result;
    }

    /// <summary>
    /// Shows a small modal input dialog prompting the user to enter a numeric value to
    /// set on the focused element (e.g. a Slider or Spinner).
    /// Returns the entered text, or <c>null</c> if the user cancelled.
    /// </summary>
    private static string? ShowValuePrompt(string elementLabel, double minimum, double maximum)
    {
        string? result = null;

        using var form = new Form
        {
            Text = "Set value",
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            Width = 380,
            Height = 145,
            MaximizeBox = false,
            MinimizeBox = false,
            TopMost = true
        };

        var label = new WinLabel
        {
            Text = $"Numeric value for \"{elementLabel}\" (min {minimum}, max {maximum}):",
            Left = 12,
            Top = 12,
            Width = 350,
            AutoSize = true
        };

        var textBox = new System.Windows.Forms.TextBox
        {
            Left = 12,
            Top = 35,
            Width = 340
        };

        var okBtn = new System.Windows.Forms.Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Left = 192,
            Top = 68,
            Width = 80
        };

        var cancelBtn = new System.Windows.Forms.Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = 278,
            Top = 68,
            Width = 80
        };

        form.AcceptButton = okBtn;
        form.CancelButton = cancelBtn;
        form.Controls.AddRange([label, textBox, okBtn, cancelBtn]);

        if (form.ShowDialog() == DialogResult.OK)
            result = textBox.Text;

        return result;
    }

    /// <summary>
    /// A <see cref="ContextMenuStrip"/> that carries the <c>WS_EX_NOACTIVATE</c>
    /// extended window style so that showing it does not steal input focus from the
    /// currently active window.
    ///
    /// <para>
    /// Without this, displaying the assistive context menu over a modal popup window
    /// (UIA <c>IsModal = true</c>) sends <c>WM_ACTIVATE(WA_INACTIVE)</c> to the popup,
    /// causing light-dismiss popups to close before the user can select an action.
    /// With <c>WS_EX_NOACTIVATE</c> the previously active window retains focus, so
    /// the popup stays open and the simulated right-click (or click) reaches it.
    /// </para>
    ///
    /// <para>
    /// Mouse clicks on menu items are still delivered based on hit-testing and are
    /// unaffected by the lack of activation; keyboard navigation of the menu is not
    /// supported, but the assistive overlay is a mouse-driven tool.
    /// </para>
    /// </summary>
    private sealed class NoActivateContextMenuStrip : ContextMenuStrip
    {
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE;
                return cp;
            }
        }
    }
}
