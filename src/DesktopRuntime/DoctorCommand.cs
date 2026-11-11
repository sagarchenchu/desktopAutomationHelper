using System.Text.Json;
using DesktopExecutionRuntime;

namespace DesktopRuntime;

public static class DoctorCommand
{
    public static int Run(string[] args, IDesktopEnvironmentProbe probe, TextWriter output, TextWriter error)
    {
        const string usage = "Usage: DesktopRuntime doctor [--json]";
        if (args is ["--help"] or ["-h"])
        {
            output.WriteLine(usage);
            return 0;
        }
        if (args is not ["doctor"] and not ["doctor", "--json"])
        {
            error.WriteLine(usage);
            return 64;
        }

        var diagnostics = DesktopDiagnostics.Evaluate(probe.Capture());
        if (args.Length == 2)
        {
            output.WriteLine(JsonSerializer.Serialize(diagnostics,
                new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true }));
        }
        else
        {
            output.WriteLine(diagnostics.PreflightPassed ? "Desktop preflight passed." : "Desktop preflight blocked.");
            output.WriteLine($"Process: {diagnostics.Snapshot.ProcessId}; session: {diagnostics.Snapshot.SessionId}; state: {diagnostics.Snapshot.ConnectionState}");
            foreach (var reason in diagnostics.BlockingReasons) output.WriteLine($"  {reason}");
            output.WriteLine("Input, capture, target compatibility, and execution without a viewer remain unverified.");
            output.WriteLine("Use --json for observations and native errors. This command does not create a display or session.");
        }
        return diagnostics.PreflightPassed ? 0 : 2;
    }
}
