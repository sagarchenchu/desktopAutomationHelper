using System.Runtime.InteropServices;
using DesktopAutomationDriver.Models.Recording;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
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
    private const int WM_LBUTTONDBLCLK = 0x0203;
    private const int WM_RBUTTONDOWN = 0x0204;

    private const byte VK_CONTROL = 0x11;
    private const byte VK_P = 0x50;
    private const byte VK_A = 0x41;
    private const byte VK_S = 0x53;

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

    /// <summary>Maximum number of child elements shown in the "Children ▶" submenu.</summary>
    private const int MaxChildrenToDisplay = 30;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

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

        var pt = Cursor.Position;
        var info = _service.GetElementAtPoint(pt);
        var name = info?.Name ?? info?.AutomationId ?? info?.ControlType ?? "unknown";
        var ct = info?.ControlType ?? string.Empty;

        // Truncate the element name so the label always fits in the 420 px corner widget
        // (≈ 9.5 pt Segoe UI → ~18 px/char → 22 chars keeps the label under ~400 px).
        const int MaxNameLen = 22;
        var displayName = name.Length > MaxNameLen ? name[..MaxNameLen] + "…" : name;

        var text = $"● [{ct}] {displayName}  Right-click";
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
            }
        }

        return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private void ActivateMode(RecordingMode mode)
    {
        _service.SetMode(mode);
        var label = mode == RecordingMode.Passive
            ? "● PASSIVE  Recording clicks & keys  S=Stop"
            : "● ASSISTIVE  Right-click element  S=Stop";
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
                BeginInvoke(new Action(() => RecordPassiveClick(pt, ActionType.Click)));
            }
            else if (wParam == (IntPtr)WM_LBUTTONDBLCLK)
            {
                var pt = ms.pt;
                BeginInvoke(new Action(() => RecordPassiveClick(pt, ActionType.DoubleClick)));
            }
        }
        else if (_service.CurrentMode == RecordingMode.Assistive)
        {
            if (wParam == (IntPtr)WM_RBUTTONDOWN)
            {
                var pt = ms.pt;
                BeginInvoke(new Action(() => ShowAssistiveContextMenu(pt)));
                return (IntPtr)1; // suppress the native right-click
            }
        }

        return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    // ── Passive recording helpers ────────────────────────────────────────────
    private void RecordPassiveClick(System.Drawing.Point pt, ActionType actionType)
    {
        var info = _service.GetElementAtPoint(pt);
        var actionLabel = actionType == ActionType.DoubleClick ? "Double Click" : "Click";
        _service.AddAction(new RecordedAction
        {
            ActionType = actionType,
            Mode = RecordingMode.Passive,
            Element = info,
            Description = BuildDescription(actionLabel, info)
        });
    }

    // ── Assistive mode: context menu ─────────────────────────────────────────
    private void ShowAssistiveContextMenu(System.Drawing.Point pt)
    {
        if (_automation == null) return;

        AutomationElement? element = null;
        try { element = _automation.FromPoint(pt); }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not get element at point {Pt} for assistive menu", pt);
        }

        // For container controls (e.g. WinForms MenuStrip/MenuBar), FromPoint may return
        // the parent instead of the specific child under the cursor. Drill down one level.
        if (element != null)
            element = DrillDownToElementAtPoint(element, pt);

        var elementInfo = element != null ? BuildElementInfo(element) : null;

        var menu = new ContextMenuStrip { ShowImageMargin = false };
        menu.Font = new Font("Segoe UI", 10f);

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
            () => element?.Click());
        AddActionItem(menu, "Double Click", element, elementInfo, ActionType.DoubleClick,
            () => element?.DoubleClick());
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

        // ── Children submenu (for container controls) ────────────────────────
        if (element != null && IsContainer(element.ControlType))
        {
            try
            {
                var children = element.FindAllChildren();
                if (children.Length > 0)
                {
                    menu.Items.Add(new ToolStripSeparator());
                    var childrenMenu = new ToolStripMenuItem("Children ▶");
                    foreach (var child in children.Take(MaxChildrenToDisplay))
                    {
                        var childInfo = BuildElementInfo(child);
                        var label = $"[{childInfo.ControlType}]  {childInfo.Name ?? childInfo.AutomationId ?? "(unnamed)"}";
                        var childItem = new ToolStripMenuItem(label);
                        var capturedChild = child;
                        var capturedChildInfo = childInfo;

                        childItem.Click += (_, _) =>
                        {
                            _service.AddAction(new RecordedAction
                            {
                                ActionType = ActionType.Select,
                                Mode = RecordingMode.Assistive,
                                Element = capturedChildInfo,
                                Description = BuildDescription("Select", capturedChildInfo)
                            });
                            try
                            {
                                if (capturedChild.Patterns.SelectionItem.IsSupported)
                                    capturedChild.Patterns.SelectionItem.Pattern.Select();
                                else
                                    capturedChild.Click();
                            }
                            catch { /* best effort */ }
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

        menu.Show(pt);
        menu.Disposed += (_, _) => menu.Items.Clear();
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
                _statusLabel.Text = "  Assistive ACTIVE  │  Right-click any element for available actions  │  Ctrl+P = Passive  │  Ctrl+S = Stop  ";
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
        ct == ControlType.ComboBox ||
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
}
