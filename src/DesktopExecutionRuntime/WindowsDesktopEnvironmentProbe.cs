using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace DesktopExecutionRuntime;

/// <summary>
/// Read-only, per-call observations. Never switches desktops, sends input, opens RDP,
/// captures the screen, unlocks a session, or modifies a display.
/// </summary>
public sealed class WindowsDesktopEnvironmentProbe : IDesktopEnvironmentProbe
{
    public DesktopSnapshot Capture()
    {
        var snapshot = new DesktopSnapshot
        {
            IsWindows = OperatingSystem.IsWindows(),
            OperatingSystem = RuntimeInformation.OSDescription,
            ProcessId = Environment.ProcessId
        };
        return snapshot.IsWindows && OperatingSystem.IsWindows() ? CaptureWindows(snapshot) : snapshot;
    }

    [SupportedOSPlatform("windows")]
    private static DesktopSnapshot CaptureWindows(DesktopSnapshot snapshot)
    {
        var failures = new List<ProbeFailure>();
        snapshot = snapshot with { ProbeFailures = failures };
        try
        {
            if (Native.ProcessIdToSessionId((uint)snapshot.ProcessId, out var sessionId))
            {
                snapshot = snapshot with
                {
                    SessionId = checked((int)sessionId),
                    ConnectionState = ReadConnectionState(sessionId, failures)
                };
            }
            else AddFailure(failures, "ProcessIdToSessionId");

            // These two handles are borrowed: do not close them.
            var station = Native.GetProcessWindowStation();
            var desktop = Native.GetThreadDesktop(Native.GetCurrentThreadId());
            snapshot = snapshot with
            {
                WindowStation = ReadName(station, "WindowStation", failures),
                WindowStationVisible = ReadStationVisible(station, failures),
                ThreadDesktop = ReadName(desktop, "ThreadDesktop", failures),
                ThreadDesktopReceivesInput = ReadReceivesInput(desktop, failures)
            };

            // DESKTOP_READOBJECTS only; requesting switch access is unnecessary.
            var input = Native.OpenInputDesktop(0, false, 0x0001);
            if (input == IntPtr.Zero) AddFailure(failures, "OpenInputDesktop");
            else
            {
                try { snapshot = snapshot with { InputDesktop = ReadName(input, "InputDesktop", failures) }; }
                finally
                {
                    if (!Native.CloseDesktop(input)) AddFailure(failures, "CloseDesktop");
                }
            }

            snapshot = snapshot with
            {
                MonitorCount = Native.GetSystemMetrics(80), // SM_CMONITORS
                VirtualScreenWidth = Native.GetSystemMetrics(78), // SM_CXVIRTUALSCREEN
                VirtualScreenHeight = Native.GetSystemMetrics(79) // SM_CYVIRTUALSCREEN
            };
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException or OverflowException)
        {
            failures.Add(new ProbeFailure("WindowsDesktopProbe", null, ex.GetType().Name));
        }
        return snapshot;
    }

    private static string? ReadConnectionState(uint sessionId, List<ProbeFailure> failures)
    {
        var success = Native.WTSQuerySessionInformationW(IntPtr.Zero, sessionId, 8, out var buffer, out var length);
        var error = Marshal.GetLastWin32Error();
        try
        {
            if (!success || buffer == IntPtr.Zero || length < sizeof(int))
            {
                failures.Add(new ProbeFailure("WTSQuerySessionInformationW", success ? null : error,
                    "Session connection state could not be read."));
                return null;
            }
            var state = Marshal.ReadInt32(buffer);
            string[] states = { "Active", "Connected", "ConnectQuery", "Shadow", "Disconnected",
                "Idle", "Listen", "Reset", "Down", "Init" };
            return state >= 0 && state < states.Length ? states[state] : $"Unknown({state})";
        }
        finally { if (buffer != IntPtr.Zero) Native.WTSFreeMemory(buffer); }
    }

    private static string? ReadName(IntPtr handle, string kind, List<ProbeFailure> failures)
    {
        if (handle == IntPtr.Zero)
        {
            failures.Add(new ProbeFailure(kind, null, "Object handle unavailable."));
            return null;
        }
        Native.GetUserObjectInformationW(handle, 2, IntPtr.Zero, 0, out var needed);
        if (needed == 0 || needed > 65536)
        {
            AddFailure(failures, $"GetUserObjectInformationW({kind}.Name)");
            return null;
        }
        var buffer = Marshal.AllocHGlobal(checked((int)needed));
        try
        {
            if (Native.GetUserObjectInformationW(handle, 2, buffer, needed, out _))
                return Marshal.PtrToStringUni(buffer);
            AddFailure(failures, $"GetUserObjectInformationW({kind}.Name)");
            return null;
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    private static bool? ReadStationVisible(IntPtr station, List<ProbeFailure> failures)
    {
        // USEROBJECTFLAGS: BOOL fInherit; BOOL fReserved; DWORD dwFlags.
        var buffer = Marshal.AllocHGlobal(12);
        try
        {
            if (Native.GetUserObjectInformationW(station, 1, buffer, 12, out _))
                return (Marshal.ReadInt32(buffer, 8) & 1) != 0; // WSF_VISIBLE
            AddFailure(failures, "GetUserObjectInformationW(WindowStation.Flags)");
            return null;
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    private static bool? ReadReceivesInput(IntPtr desktop, List<ProbeFailure> failures)
    {
        var buffer = Marshal.AllocHGlobal(sizeof(int));
        try
        {
            if (Native.GetUserObjectInformationW(desktop, 6, buffer, sizeof(int), out _)) // UOI_IO
                return Marshal.ReadInt32(buffer) != 0;
            AddFailure(failures, "GetUserObjectInformationW(ThreadDesktop.IO)");
            return null;
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    private static void AddFailure(List<ProbeFailure> failures, string api) =>
        failures.Add(new ProbeFailure(api, Marshal.GetLastWin32Error(), "Native observation failed."));

    private static class Native
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ProcessIdToSessionId(uint processId, out uint sessionId);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetProcessWindowStation();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetThreadDesktop(uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr OpenInputDesktop(uint flags, [MarshalAs(UnmanagedType.Bool)] bool inherit, uint access);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseDesktop(IntPtr desktop);

        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetUserObjectInformationW(IntPtr handle, int index, IntPtr buffer, uint length, out uint needed);

        [DllImport("user32.dll")]
        internal static extern int GetSystemMetrics(int index);

        [DllImport("wtsapi32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool WTSQuerySessionInformationW(IntPtr server, uint sessionId, int infoClass, out IntPtr buffer, out uint bytes);

        [DllImport("wtsapi32.dll")]
        internal static extern void WTSFreeMemory(IntPtr memory);
    }
}
