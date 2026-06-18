using System.Drawing;
using System.Runtime.InteropServices;

namespace DesktopAutomationDriver.NativeUia;

internal static class NativeUiaInput
{
    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    public static bool ClickPoint(Point point)
    {
        SetCursorPos(point.X, point.Y);
        Thread.Sleep(30);

        var inputs = new[]
        {
            MouseInput(MouseEventFlags.MOUSEEVENTF_LEFTDOWN),
            MouseInput(MouseEventFlags.MOUSEEVENTF_LEFTUP)
        };

        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>()) == inputs.Length;
    }

    public static void WheelDown(int clicks = 3)
    {
        for (var i = 0; i < clicks; i++)
        {
            var input = new INPUT
            {
                type = 0,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        mouseData = unchecked((uint)-120),
                        dwFlags = MouseEventFlags.MOUSEEVENTF_WHEEL
                    }
                }
            };
            SendInput(1, [input], Marshal.SizeOf<INPUT>());
            Thread.Sleep(40);
        }
    }

    public static void SendKey(ushort key, bool release = false)
    {
        var input = new INPUT
        {
            type = 1,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = key,
                    dwFlags = release ? KeyboardEventFlags.KEYEVENTF_KEYUP : 0
                }
            }
        };
        SendInput(1, [input], Marshal.SizeOf<INPUT>());
    }

    public static void SendChord(params ushort[] keys)
    {
        foreach (var key in keys)
            SendKey(key);
        Thread.Sleep(30);
        for (var i = keys.Length - 1; i >= 0; i--)
            SendKey(keys[i], release: true);
    }

    public static void TypeText(string text)
    {
        foreach (var ch in text)
        {
            var down = new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = (ushort)ch,
                        dwFlags = KeyboardEventFlags.KEYEVENTF_UNICODE
                    }
                }
            };
            var up = down;
            up.U.ki.dwFlags = KeyboardEventFlags.KEYEVENTF_UNICODE | KeyboardEventFlags.KEYEVENTF_KEYUP;
            SendInput(1, [down], Marshal.SizeOf<INPUT>());
            SendInput(1, [up], Marshal.SizeOf<INPUT>());
            Thread.Sleep(5);
        }
    }

    public static IntPtr ForegroundWindowHandle() => GetForegroundWindow();

    private static INPUT MouseInput(MouseEventFlags flags) => new()
    {
        type = 0,
        U = new InputUnion { mi = new MOUSEINPUT { dwFlags = flags } }
    };

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public MouseEventFlags dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public KeyboardEventFlags dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [Flags]
    private enum MouseEventFlags : uint
    {
        MOUSEEVENTF_LEFTDOWN = 0x0002,
        MOUSEEVENTF_LEFTUP = 0x0004,
        MOUSEEVENTF_WHEEL = 0x0800
    }

    [Flags]
    private enum KeyboardEventFlags : uint
    {
        KEYEVENTF_KEYUP = 0x0002,
        KEYEVENTF_UNICODE = 0x0004
    }

    internal static class VirtualKeys
    {
        public const ushort Return = 0x0D;
        public const ushort Space = 0x20;
        public const ushort Down = 0x28;
        public const ushort F4 = 0x73;
        public const ushort Menu = 0x12;
        public const ushort Next = 0x22;
        public const ushort Escape = 0x1B;
        public const ushort Control = 0x11;
        public const ushort A = 0x41;
    }
}
