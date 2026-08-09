using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Readiness;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class DoctorJsonCliTests
{
    [Fact]
    public async Task DoctorJson_StdoutContainsExactlyOneValidJsonDocument()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var handler = new FakeHttpMessageHandler(req =>
        {
            if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    status = 0,
                    value = new
                    {
                        ready = true,
                        message = "ok",
                        build = new { version = "1.0.105" }
                    }
                });
            }

            return FakeHttpMessageHandler.Json(new
            {
                success = true,
                value = CatalogFixtures.ValidCatalog()
            });
        });

        var factory = TestSupport.CreateFactory(handler);

        IHost HostBuilder(string[] _, bool jsonMode)
        {
            var builder = Host.CreateApplicationBuilder();
            builder.Logging.ClearProviders();
            builder.Logging.AddSimpleConsole();
            builder.Services.Configure<Microsoft.Extensions.Logging.Console.ConsoleLoggerOptions>(o =>
            {
                o.LogToStandardErrorThreshold = LogLevel.Trace;
            });
            if (jsonMode)
                builder.Logging.AddFilter("DesktopAutomationAgent", LogLevel.Warning);

            builder.Services.AddSingleton(Options.Create(options));
            builder.Services.AddSingleton<IWorkspaceManager>(workspace);
            builder.Services.AddSingleton<ISuiteManifestReader>(_ =>
                new SuiteManifestReader(Options.Create(options), workspace));
            builder.Services.AddSingleton<IDriverConnectionResolver>(_ =>
                new DriverConnectionResolver(Options.Create(options), factory, NullLogger<DriverConnectionResolver>.Instance));
            builder.Services.AddSingleton<IDriverCatalogClient>(_ =>
                new DriverCatalogClient(Options.Create(options), factory, NullLogger<DriverCatalogClient>.Instance));
            builder.Services.AddSingleton<IAgentReadinessService, AgentReadinessService>();
            return builder.Build();
        }

        var originalOut = Console.Out;
        var originalErr = Console.Error;
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        Console.SetOut(stdout);
        Console.SetError(stderr);

        int exitCode;
        try
        {
            // Emit an information log through the host path indirectly by running doctor.
            exitCode = await AgentCli.RunAsync(["doctor", "--json"], HostBuilder);
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }

        Assert.Equal(ExitCodes.Success, exitCode);

        var stdoutText = stdout.ToString().Trim();
        Assert.False(string.IsNullOrWhiteSpace(stdoutText));
        AssertExactlyOneJsonDocument(stdoutText);

        using var document = JsonDocument.Parse(stdoutText);
        Assert.Equal(JsonValueKind.Object, document.RootElement.ValueKind);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.DoesNotContain("secret-token", stdoutText, StringComparison.Ordinal);

        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    private static void AssertExactlyOneJsonDocument(string stdoutText)
    {
        var bytes = Encoding.UTF8.GetBytes(stdoutText);
        var reader = new Utf8JsonReader(bytes);
        Assert.True(JsonDocument.TryParseValue(ref reader, out var first));
        first!.Dispose();
        Assert.Equal(bytes.Length, reader.BytesConsumed);
    }
}
