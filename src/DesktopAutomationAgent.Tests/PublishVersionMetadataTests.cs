using System.Diagnostics;
using System.Reflection;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Tests;

/// <summary>
/// Lightweight checks that Version/InformationalVersion MSBuild properties flow into
/// assembly metadata. CI publish steps also assert ProductVersion on the published EXE.
/// </summary>
public class PublishVersionMetadataTests
{
    [Fact]
    public void AgentAssembly_InformationalVersion_IsPresent()
    {
        var informational = typeof(JiraKeyContract).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

        Assert.False(string.IsNullOrWhiteSpace(informational));
        // Local/CI builds use 0.0.0-local unless overridden; release injects 1.0.<n>.
        Assert.Matches(@"^\d+\.\d+\.\d+", informational!.Split('+')[0]);
    }

    [Fact]
    public void DotnetBuild_HonorsInjectedInformationalVersion()
    {
        var repoRoot = FindRepoRoot();
        var output = Path.Combine(Path.GetTempPath(), "da-agent-ver-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);

        try
        {
            const string injected = "9.8.7";
            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments =
                    $"build \"{Path.Combine(repoRoot, "src/DesktopAutomationAgent/DesktopAutomationAgent.csproj")}\" " +
                    $"--configuration Release --nologo -v q " +
                    $"-p:Version={injected} " +
                    $"-p:InformationalVersion={injected} " +
                    $"-p:IncludeSourceRevisionInInformationalVersion=false " +
                    $"-p:OutputPath={output}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start dotnet build.");
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            Assert.True(process.WaitForExit(120_000), "dotnet build timed out.");
            Assert.True(process.ExitCode == 0, $"dotnet build failed.\n{stdout}\n{stderr}");

            var dll = Path.Combine(output, "DesktopAutomationAgent.dll");
            Assert.True(File.Exists(dll), $"Expected build output at {dll}");

            var asm = Assembly.LoadFrom(dll);
            var informational = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            Assert.Equal(injected, informational);
        }
        finally
        {
            try { Directory.Delete(output, recursive: true); } catch { /* ignore */ }
        }
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.slnx")))
                return dir.FullName;
            dir = dir.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root from test base directory.");
    }
}
