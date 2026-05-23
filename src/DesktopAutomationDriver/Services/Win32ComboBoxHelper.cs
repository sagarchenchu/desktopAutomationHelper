using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Helpers for selecting items in native Win32 ComboBox controls using raw window messages.
/// This mirrors the pywinauto Win32 backend approach: CB_FINDSTRINGEXACT / CB_SETCURSEL
/// for silent, non-visual selection without opening the dropdown.
/// </summary>
internal static class Win32ComboBoxHelper
{
    // -------------------------------------------------------------------------
    // ComboBox window messages
    // -------------------------------------------------------------------------

    private const int CB_GETCOUNT        = 0x0146;
    private const int CB_GETCURSEL       = 0x0147;
    private const int CB_GETLBTEXT       = 0x0148;
    private const int CB_GETLBTEXTLEN    = 0x0149;
    private const int CB_SELECTSTRING    = 0x014D;
    private const int CB_SETCURSEL       = 0x014E;
    private const int CB_SHOWDROPDOWN    = 0x014F;
    private const int CB_FINDSTRINGEXACT = 0x0158;

    // -------------------------------------------------------------------------
    // WM_COMMAND / notification codes
    // -------------------------------------------------------------------------

    private const int WM_COMMAND  = 0x0111;
    private const int CBN_SELCHANGE = 1;
    private const int CBN_SELENDOK  = 9;
    private const int CBN_CLOSEUP   = 8;
    private const int GWL_ID        = -12;

    // -------------------------------------------------------------------------
    // P/Invokes
    // -------------------------------------------------------------------------

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, StringBuilder lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    // -------------------------------------------------------------------------
    // Detection
    // -------------------------------------------------------------------------

    /// <summary>
    /// Returns true when <paramref name="hwnd"/> is likely a native Win32 ComboBox.
    /// Detection is based on class name equality or a successful CB_GETCOUNT probe.
    /// </summary>
    public static bool IsLikelyNativeWin32ComboBox(IntPtr hwnd, string? className)
    {
        if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
            return false;

        // Authoritative check: Win32 class name.
        if (!string.IsNullOrWhiteSpace(className) &&
            className.Equals("ComboBox", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // Fallback probe: a non-ComboBox control returns CB_ERR (-1).
        try
        {
            return SendMessage(hwnd, CB_GETCOUNT, IntPtr.Zero, IntPtr.Zero).ToInt32() >= 0;
        }
        catch
        {
            return false;
        }
    }

    // -------------------------------------------------------------------------
    // Read helpers
    // -------------------------------------------------------------------------

    /// <summary>Returns the total number of items in the ComboBox list.</summary>
    public static int GetCount(IntPtr hwnd)
        => SendMessage(hwnd, CB_GETCOUNT, IntPtr.Zero, IntPtr.Zero).ToInt32();

    /// <summary>Returns the zero-based index of the currently selected item, or -1.</summary>
    public static int GetCurrentIndex(IntPtr hwnd)
        => SendMessage(hwnd, CB_GETCURSEL, IntPtr.Zero, IntPtr.Zero).ToInt32();

    /// <summary>Returns the text of the item at <paramref name="index"/>, or null on error.</summary>
    public static string? GetItemText(IntPtr hwnd, int index)
    {
        var len = SendMessage(hwnd, CB_GETLBTEXTLEN, new IntPtr(index), IntPtr.Zero).ToInt32();
        if (len < 0)
            return null;

        var sb = new StringBuilder(len + 1);
        SendMessage(hwnd, CB_GETLBTEXT, new IntPtr(index), sb);
        return sb.ToString();
    }

    /// <summary>Returns the text of the currently selected item, or null.</summary>
    public static string? GetSelectedText(IntPtr hwnd)
    {
        var index = GetCurrentIndex(hwnd);
        return index < 0 ? null : GetItemText(hwnd, index);
    }

    /// <summary>
    /// Uses CB_FINDSTRINGEXACT to locate <paramref name="value"/> (case-insensitive exact match).
    /// Returns the zero-based index, or -1 when not found.
    /// </summary>
    public static int FindExactIndex(IntPtr hwnd, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return -1;

        return SendMessage(hwnd, CB_FINDSTRINGEXACT, new IntPtr(-1), value).ToInt32();
    }

    // -------------------------------------------------------------------------
    // Selection
    // -------------------------------------------------------------------------

    /// <summary>
    /// Selects the item at <paramref name="index"/> using CB_SETCURSEL and posts the
    /// standard WM_COMMAND selection notifications to the parent window.
    /// </summary>
    public static bool SelectByIndex(IntPtr hwnd, int index, ILogger logger, string source)
    {
        if (hwnd == IntPtr.Zero || index < 0)
            return false;

        var count = GetCount(hwnd);
        if (count < 0 || index >= count)
            return false;

        logger.LogInformation(
            "Win32 ComboBox selecting by index. source={Source}, hwnd=0x{Hwnd:X}, index={Index}, count={Count}",
            source, hwnd.ToInt64(), index, count);

        var result = SendMessage(hwnd, CB_SETCURSEL, new IntPtr(index), IntPtr.Zero).ToInt32();

        NotifySelectionChanged(hwnd);

        return result >= 0;
    }

    /// <summary>
    /// Selects the item whose text exactly matches <paramref name="value"/> (CB_FINDSTRINGEXACT),
    /// then falls back to CB_SELECTSTRING (prefix match) when no exact match is found.
    /// Posts WM_COMMAND notifications after selection.
    /// </summary>
    public static bool SelectByText(IntPtr hwnd, string value, ILogger logger, string source)
    {
        if (hwnd == IntPtr.Zero || string.IsNullOrWhiteSpace(value))
            return false;

        // 1. Exact match.
        var exactIndex = FindExactIndex(hwnd, value);
        if (exactIndex >= 0)
            return SelectByIndex(hwnd, exactIndex, logger, source + "-exact");

        // 2. Prefix/starts-with fallback (CB_SELECTSTRING).
        var result = SendMessage(hwnd, CB_SELECTSTRING, new IntPtr(-1), value).ToInt32();
        if (result >= 0)
        {
            logger.LogInformation(
                "Win32 ComboBox selected by CB_SELECTSTRING fallback. source={Source}, hwnd=0x{Hwnd:X}, value={Value}, index={Index}",
                source, hwnd.ToInt64(), value, result);

            NotifySelectionChanged(hwnd);
            return true;
        }

        logger.LogInformation(
            "Win32 ComboBox did not find value. source={Source}, hwnd=0x{Hwnd:X}, value={Value}",
            source, hwnd.ToInt64(), value);

        return false;
    }

    /// <summary>Closes the ComboBox dropdown if it is open.</summary>
    public static void HideDropdown(IntPtr hwnd)
    {
        try
        {
            SendMessage(hwnd, CB_SHOWDROPDOWN, IntPtr.Zero, IntPtr.Zero);
        }
        catch
        {
            // Ignore; some hosts do not need explicit close.
        }
    }

    // -------------------------------------------------------------------------
    // Notification helper
    // -------------------------------------------------------------------------

    /// <summary>
    /// Posts CBN_SELCHANGE, CBN_SELENDOK, and CBN_CLOSEUP notifications to the
    /// parent window so the application processes the selection change.
    /// </summary>
    private static void NotifySelectionChanged(IntPtr hwnd)
    {
        try
        {
            var parent = GetParent(hwnd);
            if (parent == IntPtr.Zero)
                return;

            var controlId = GetWindowLong(hwnd, GWL_ID);

            // WM_COMMAND wParam layout (Win32 standard):
            //   HIWORD = notification code  (CBN_*)
            //   LOWORD = control ID
            PostMessage(parent, WM_COMMAND,
                new IntPtr((CBN_SELCHANGE << 16) | (controlId & 0xFFFF)), hwnd);

            PostMessage(parent, WM_COMMAND,
                new IntPtr((CBN_SELENDOK << 16) | (controlId & 0xFFFF)), hwnd);

            PostMessage(parent, WM_COMMAND,
                new IntPtr((CBN_CLOSEUP << 16) | (controlId & 0xFFFF)), hwnd);
        }
        catch
        {
            // Some apps do not need WM_COMMAND notifications; swallow silently.
        }
    }
}
