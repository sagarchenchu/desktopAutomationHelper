using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Helpers for the Assistive-mode Ctrl+Right-Click handler that detect whether a
/// modal or owned popup window is currently active, and provide common popup actions
/// (Click OK, Close, Inspect) that the recording overlay can expose to the user.
/// </summary>
public static class AssistivePopupResolver
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    /// <summary>
    /// Returns the foreground window as a FlaUI <see cref="Window"/>, or
    /// <c>null</c> if the foreground window cannot be resolved.
    /// </summary>
    public static Window? GetForegroundWindowElement(UIA3Automation automation)
    {
        var hwnd = GetForegroundWindow();

        if (hwnd == IntPtr.Zero)
            return null;

        AutomationElement? element;
        try
        {
            element = automation.FromHandle(hwnd);
        }
        catch
        {
            return null;
        }

        if (element == null)
            return null;

        if (element.ControlType != ControlType.Window)
        {
            var windowAncestor = FindAncestorWindow(element);
            return windowAncestor?.AsWindow();
        }

        return element.AsWindow();
    }

    /// <summary>
    /// Returns <c>true</c> when <paramref name="candidate"/> has a different native
    /// window handle from <paramref name="mainWindow"/>, indicating it is a separate
    /// popup or dialog window.
    /// </summary>
    public static bool IsDifferentWindow(Window? candidate, AutomationElement? mainWindow)
    {
        if (candidate == null || mainWindow == null)
            return false;

        var candidateHandle = SafeWindowHandle(candidate);
        var mainHandle = SafeWindowHandle(mainWindow);

        if (candidateHandle == IntPtr.Zero || mainHandle == IntPtr.Zero)
            return false;

        return candidateHandle != mainHandle;
    }

    /// <summary>
    /// Returns <c>true</c> when <paramref name="element"/> (or any of its ancestors)
    /// is a TitleBar, MenuBar, MenuItem, or Menu – i.e. part of the window chrome or
    /// the system ("Application") menu.
    /// </summary>
    public static bool IsInsideChromeOrSystemMenu(AutomationElement? element)
    {
        if (element == null)
            return false;

        try
        {
            var current = element;

            while (current != null)
            {
                if (current.ControlType == ControlType.TitleBar ||
                    current.ControlType == ControlType.MenuBar ||
                    current.ControlType == ControlType.MenuItem ||
                    current.ControlType == ControlType.Menu)
                {
                    return true;
                }

                var name = current.Name ?? string.Empty;
                var automationId = current.AutomationId ?? string.Empty;

                if (name.Contains("System Menu", StringComparison.OrdinalIgnoreCase) ||
                    automationId.Contains("SystemMenu", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("System", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                current = current.Parent;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    /// <summary>
    /// Searches the direct child window elements of <paramref name="mainWindow"/> for
    /// a visible, named window that is different from <paramref name="mainWindow"/> itself.
    /// Returns the first match as a <see cref="Window"/>, or <c>null</c> if none is found.
    /// </summary>
    public static Window? DetectChildPopupWindow(AutomationElement mainWindow, AutomationBase automation)
    {
        if (mainWindow == null)
            return null;

        var cf = automation.ConditionFactory;
        var mainHandle = SafeWindowHandle(mainWindow);

        AutomationElement[] childWindows;
        try
        {
            childWindows = mainWindow.FindAllDescendants(cf.ByControlType(ControlType.Window));
        }
        catch
        {
            return null;
        }

        foreach (var child in childWindows)
        {
            try
            {
                var childHandle = SafeWindowHandle(child);

                if (childHandle == IntPtr.Zero)
                    continue;

                if (childHandle == mainHandle)
                    continue;

                if (child.IsOffscreen)
                    continue;

                if (string.IsNullOrWhiteSpace(child.Name))
                    continue;

                return child.AsWindow();
            }
            catch
            {
                // Ignore unstable UIA nodes.
            }
        }

        return null;
    }

    /// <summary>
    /// Finds the best "OK" button inside <paramref name="popupWindow"/> using
    /// <paramref name="cf"/> to enumerate buttons.  Skips buttons that are
    /// offscreen, disabled, or inside the window chrome / system menu.
    /// Preference order: <paramref name="preferredAutomationIds"/> first, then
    /// <paramref name="names"/> (case-insensitive, ignoring leading ampersands).
    /// Returns <c>null</c> if no matching button is found.
    /// </summary>
    public static AutomationElement? FindPopupButton(
        Window popupWindow,
        string[] names,
        string[] preferredAutomationIds,
        ConditionFactory cf)
    {
        AutomationElement[] buttons;
        try
        {
            buttons = popupWindow
                .FindAllDescendants(cf.ByControlType(ControlType.Button));
        }
        catch
        {
            return null;
        }

        var candidates = buttons.Where(b =>
        {
            try
            {
                if (b.IsOffscreen)
                    return false;

                if (!b.IsEnabled)
                    return false;

                if (IsInsideChromeOrSystemMenu(b))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }).ToList();

        foreach (var id in preferredAutomationIds)
        {
            var match = candidates.FirstOrDefault(b =>
                string.Equals(b.AutomationId ?? string.Empty, id, StringComparison.OrdinalIgnoreCase));

            if (match != null)
                return match;
        }

        foreach (var name in names)
        {
            var match = candidates.FirstOrDefault(b =>
                string.Equals(NormalizeName(b.Name), NormalizeName(name), StringComparison.OrdinalIgnoreCase));

            if (match != null)
                return match;
        }

        return null;
    }

    /// <summary>
    /// Finds the OK/Yes/Continue button inside <paramref name="popupWindow"/> and
    /// invokes it, or falls back to pressing Enter if no button is found.
    /// </summary>
    public static AutomationElement? FindOkButton(Window popupWindow, ConditionFactory cf)
        => FindPopupButton(
            popupWindow,
            new[] { "OK", "&OK", "Yes", "&Yes", "Continue", "Save" },
            new[] { "2", "OK", "Yes", "Continue", "Save" },
            cf);

    /// <summary>
    /// Finds the Cancel/No button inside <paramref name="popupWindow"/> and invokes it,
    /// or falls back to pressing Escape if no button is found.
    /// </summary>
    public static AutomationElement? FindCancelButton(Window popupWindow, ConditionFactory cf)
        => FindPopupButton(
            popupWindow,
            new[] { "Cancel", "&Cancel", "No", "&No" },
            new[] { "Cancel", "No" },
            cf);

    /// <summary>
    /// Invokes the element's Invoke pattern if supported; otherwise uses a mouse click.
    /// </summary>
    public static void InvokeOrClick(AutomationElement element)
    {
        if (element.Patterns.Invoke.IsSupported)
            element.Patterns.Invoke.Pattern.Invoke();
        else
            element.Click();
    }

    /// <summary>
    /// Returns an inspection snapshot of all descendants of <paramref name="popupWindow"/>,
    /// including their name, automation ID, control type, class name, HWND, visibility,
    /// and whether they sit inside the window chrome or system menu.
    /// </summary>
    public static object Inspect(Window popupWindow)
    {
        var elements = popupWindow
            .FindAllDescendants()
            .Select(e =>
            {
                try
                {
                    return new
                    {
                        name = e.Name,
                        automationId = e.AutomationId,
                        controlType = e.ControlType.ToString(),
                        className = e.ClassName,
                        nativeWindowHandle = SafeWindowHandle(e).ToInt64(),
                        isInsideChrome = IsInsideChromeOrSystemMenu(e),
                        enabled = SafeIsEnabled(e),
                        visible = SafeIsVisible(e)
                    };
                }
                catch
                {
                    return null;
                }
            })
            .Where(x => x != null)
            .ToList();

        return new
        {
            popup = new
            {
                name = popupWindow.Name,
                automationId = popupWindow.AutomationId,
                controlType = popupWindow.ControlType.ToString(),
                nativeWindowHandle = SafeWindowHandle(popupWindow).ToInt64()
            },
            elements
        };
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private static AutomationElement? FindAncestorWindow(AutomationElement element)
    {
        var current = element;

        while (current != null)
        {
            if (current.ControlType == ControlType.Window)
                return current;

            try { current = current.Parent; }
            catch { return null; }
        }

        return null;
    }

    private static string NormalizeName(string? value)
        => (value ?? string.Empty).Replace("&", string.Empty).Trim();

    private static bool SafeIsEnabled(AutomationElement element)
    {
        try { return element.IsEnabled; }
        catch { return false; }
    }

    private static bool SafeIsVisible(AutomationElement element)
    {
        try { return !element.IsOffscreen; }
        catch { return false; }
    }

    /// <summary>
    /// Returns the native window handle of <paramref name="element"/>, or
    /// <see cref="IntPtr.Zero"/> if the property is unavailable.
    /// </summary>
    internal static IntPtr SafeWindowHandle(AutomationElement element)
    {
        try
        {
            return new IntPtr(element.Properties.NativeWindowHandle.ValueOrDefault);
        }
        catch
        {
            return IntPtr.Zero;
        }
    }
}
