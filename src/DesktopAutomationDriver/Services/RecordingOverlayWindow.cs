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
    private const int WM_RBUTTONUP   = 0x0205;

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

    /// <summary>
    /// Delay in milliseconds between intercepting a right-click and showing the assistive
    /// context menu.  A one-tick timer is used instead of BeginInvoke so that any
    /// window-activation or native-menu-dismissal messages triggered by the right-click
    /// are fully processed before the menu is built, ensuring the correct UIA element is
    /// identified on the first right-click over MenuBar / MenuItem elements.
    /// </summary>
    private const int ContextMenuDelayMs = 1;

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

    // Guard against showing multiple context menus simultaneously (e.g. rapid right-clicks).
    private bool _menuOpen;

    // Set to true after we suppress WM_RBUTTONDOWN so we can also suppress the
    // matching WM_RBUTTONUP (preventing the target app from receiving WM_CONTEXTMENU
    // which would open the app's own context menu and immediately close ours).
    private bool _suppressNextRButtonUp;

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
                // Capture element info eagerly in the hook callback, before the hook
                // returns and the application processes the click (which may close a
                // dialog before BeginInvoke executes, e.g. a Save As "Save"/"Cancel" button).
                ElementInfo? eagerInfo = null;
                try { eagerInfo = _service.GetElementAtPoint(pt); } catch { /* best effort */ }
                BeginInvoke(new Action(() => RecordPassiveClick(pt, eagerInfo, ActionType.Click)));
            }
            else if (wParam == (IntPtr)WM_LBUTTONDBLCLK)
            {
                var pt = ms.pt;
                ElementInfo? eagerInfo = null;
                try { eagerInfo = _service.GetElementAtPoint(pt); } catch { /* best effort */ }
                BeginInvoke(new Action(() => RecordPassiveClick(pt, eagerInfo, ActionType.DoubleClick)));
            }
        }
        else if (_service.CurrentMode == RecordingMode.Assistive)
        {
            if (wParam == (IntPtr)WM_RBUTTONDOWN)
            {
                var pt = ms.pt;
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

        var menu = new ContextMenuStrip { ShowImageMargin = false };
        menu.Font = new Font("Segoe UI", 10f);

        // Re-apply click-through styles when the menu closes so rapid right-clicks
        // cannot leave the overlay in a non-transparent state.
        menu.Closed += (_, _) =>
        {
            _menuOpen = false;
            ReapplyClickThroughStyle();
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
                if (element?.Patterns.Invoke.IsSupported == true)
                    element.Patterns.Invoke.Pattern.Invoke();
                else
                    element?.Click();
            });
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

        // Is Editable — only for Edit controls (text boxes)
        if (element != null && element.ControlType == ControlType.Edit)
        {
            AddQueryItem(menu, "Is Editable", element, elementInfo, ActionType.IsEditable,
                () => element.Patterns.Value.IsSupported && element.Patterns.Value.Pattern?.IsReadOnly == false);
        }

        // Type — only for Edit controls (text boxes)
        if (element != null && element.ControlType == ControlType.Edit)
        {
            menu.Items.Add(new ToolStripSeparator());
            var typeItem = new ToolStripMenuItem("Type…");
            typeItem.Click += (_, _) =>
            {
                var text = ShowTypePrompt(elementInfo?.Name ?? elementInfo?.AutomationId ?? "element");
                if (text == null) return; // user cancelled

                var elementLabel = ElementInfo.GetLabel(elementInfo);
                try
                {
                    if (element.Patterns.Value.IsSupported)
                        element.Patterns.Value.Pattern.SetValue(text);
                    else
                    {
                        element.Focus();
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
                    Element = elementInfo,
                    Value = text,
                    Description = $"Type '{displayText}' into {elementLabel}"
                });
                UpdateStatusAfterAction($"Type into [{elementInfo?.ControlType}] {elementLabel}");
            };
            menu.Items.Add(typeItem);
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
                                        // Use Click() as the primary action: it reliably triggers
                                        // ComboBox selection events and updates the displayed value.
                                        // SelectionItem.Select() may not fire SelectedIndexChanged on
                                        // some WinForms ComboBox implementations, leaving the
                                        // displayed text unchanged.
                                        freshItem.Click();

                                        Thread.Sleep(100);
                                        element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();
                                        selected = true;
                                    }
                                }
                                catch { /* best effort */ }
                            }

                            if (!selected)
                            {
                                // Last resort: use the (possibly stale) original reference.
                                try
                                {
                                    child.Click();

                                    Thread.Sleep(100);
                                    element.Patterns.ExpandCollapse.PatternOrDefault?.Collapse();
                                }
                                catch { /* best effort */ }
                            }

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

        AddCloseItem(menu);
        // Show the context menu relative to this overlay form so it inherits the form's
        // topmost z-order and appears above the target application's window.
        menu.Show(this, PointToClient(pt));
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
    /// Re-applies the WS_EX_TRANSPARENT | WS_EX_LAYERED extended styles so that the
    /// overlay stays click-through after a ContextMenuStrip is dismissed. Showing a
    /// popup may cause Windows to recalculate window styles on the owner form.
    /// </summary>
    private void ReapplyClickThroughStyle()
    {
        if (!IsHandleCreated || IsDisposed) return;
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
}
