using System.Diagnostics;
using System.Runtime.InteropServices;
using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Live Win32 window identity probe for vision authorization revalidation.</summary>
public sealed class Win32LiveWindowProbe : ILiveWindowProbe
{
    public DebugWindowIdentity? ProbeByHandle(IntPtr windowHandle)
    {
        if (windowHandle == IntPtr.Zero || !IsWindow(windowHandle))
            return null;

        try
        {
            GetWindowThreadProcessId(windowHandle, out var pid);
            if (pid == 0)
                return null;

            if (!GetWindowRect(windowHandle, out var rect))
                return null;

            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            if (width <= 0 || height <= 0)
                return null;

            string? processName = null;
            string? title = null;
            try
            {
                processName = Process.GetProcessById(unchecked((int)pid)).ProcessName;
            }
            catch
            {
                // optional
            }

            try
            {
                var capacity = GetWindowTextLength(windowHandle) + 1;
                if (capacity > 1)
                {
                    var buffer = new char[capacity];
                    var len = GetWindowText(windowHandle, buffer, capacity);
                    if (len > 0)
                        title = new string(buffer, 0, len);
                }
            }
            catch
            {
                // optional
            }

            var dpi = GetDpiForWindow(windowHandle);
            var scale = dpi == 0 ? 1.0 : dpi / 96.0;

            return new DebugWindowIdentity
            {
                WindowHandle = windowHandle,
                ProcessId = unchecked((int)pid),
                ProcessName = processName,
                WindowTitle = title,
                Bounds = new Rectangle(rect.Left, rect.Top, width, height),
                DpiScaleX = scale,
                DpiScaleY = scale
            };
        }
        catch
        {
            return null;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, char[] lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);
}
