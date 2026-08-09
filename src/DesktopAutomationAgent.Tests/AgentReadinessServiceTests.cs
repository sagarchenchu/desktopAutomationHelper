using System.Net;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Readiness;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class AgentReadinessServiceTests
{
    [Fact]
    public async Task Doctor_SucceedsWithHumanAndJsonShapes()
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
                        message = "Desktop Automation Driver is running",
                        build = new { version = "1.0.105" }
                    }
                });
            }

            if (req.RequestUri!.AbsolutePath.EndsWith("/ui/operations", StringComparison.OrdinalIgnoreCase))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = CatalogFixtures.ValidCatalog()
                });
            }

            return new HttpResponseMessage(HttpStatusCode.NotFound);
        });

        var factory = TestSupport.CreateFactory(handler);
        var readiness = new AgentReadinessService(
            TestSupport.Wrap(options),
            workspace,
            new DriverConnectionResolver(TestSupport.Wrap(options), factory, NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(TestSupport.Wrap(options), factory, NullLogger<DriverCatalogClient>.Instance),
            NullLogger<AgentReadinessService>.Instance);

        var report = await readiness.RunDoctorAsync();

        Assert.True(report.Success);
        Assert.Equal(ExitCodes.Success, report.ExitCode);
        Assert.Equal("1.0.105", report.DriverVersion);
        Assert.Equal(2, report.CatalogSchemaVersion);
        Assert.Equal(2, report.OperationCount);
        Assert.Equal("http://127.0.0.1:33201", report.DriverBaseUrl);
        Assert.DoesNotContain("secret-token", JsonSerializer.Serialize(report), StringComparison.Ordinal);

        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public async Task Doctor_ReturnsDriverUnavailableForVerifyFailure()
    {
        var options = TestSupport.CreateOptions();
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var handler = new FakeHttpMessageHandler(_ =>
            throw new HttpRequestException("connection refused"));

        var readiness = new AgentReadinessService(
            TestSupport.Wrap(options),
            workspace,
            new DriverConnectionResolver(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverCatalogClient>.Instance),
            NullLogger<AgentReadinessService>.Instance);

        var report = await readiness.RunDoctorAsync();

        Assert.False(report.Success);
        Assert.Equal(ExitCodes.DriverUnavailable, report.ExitCode);
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public async Task Doctor_ReturnsAuthOrCatalogForSchemaMismatch()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var badCatalog = CatalogFixtures.ValidCatalog();
        badCatalog.SchemaVersion = 9;

        var handler = new FakeHttpMessageHandler(req =>
        {
            if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    status = 0,
                    value = new { ready = true, message = "ok", build = new { version = "1.0.105" } }
                });
            }

            return FakeHttpMessageHandler.Json(new { success = true, value = badCatalog });
        });

        var factory = TestSupport.CreateFactory(handler);
        var readiness = new AgentReadinessService(
            TestSupport.Wrap(options),
            workspace,
            new DriverConnectionResolver(TestSupport.Wrap(options), factory, NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(TestSupport.Wrap(options), factory, NullLogger<DriverCatalogClient>.Instance),
            NullLogger<AgentReadinessService>.Instance);

        var report = await readiness.RunDoctorAsync();

        Assert.False(report.Success);
        Assert.Equal(ExitCodes.AuthOrCatalog, report.ExitCode);
        Directory.Delete(options.Workspace.Root, recursive: true);
    }

    [Fact]
    public async Task Doctor_ReturnsConfigExitCodeForPartialExplicitSettings()
    {
        var options = TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: null);
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();

        var readiness = new AgentReadinessService(
            TestSupport.Wrap(options),
            workspace,
            new DriverConnectionResolver(
                TestSupport.Wrap(options),
                TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK))),
                NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(
                TestSupport.Wrap(options),
                TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK))),
                NullLogger<DriverCatalogClient>.Instance),
            NullLogger<AgentReadinessService>.Instance);

        var report = await readiness.RunDoctorAsync();

        Assert.False(report.Success);
        Assert.Equal(ExitCodes.UsageOrConfiguration, report.ExitCode);
        Directory.Delete(options.Workspace.Root, recursive: true);
    }
}
