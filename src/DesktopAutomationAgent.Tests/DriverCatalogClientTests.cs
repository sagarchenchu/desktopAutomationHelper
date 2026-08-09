using System.Net;
using System.Net.Http.Headers;
using DesktopAutomationAgent.Driver;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class DriverCatalogClientTests
{
    [Fact]
    public async Task GetStatus_ParsesReadyPayloadAndSendsBearer()
    {
        var handler = new FakeHttpMessageHandler(req =>
        {
            Assert.Equal(HttpMethod.Get, req.Method);
            Assert.EndsWith("/status", req.RequestUri!.AbsolutePath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("secret", req.Headers.Authorization?.Parameter);
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
        });

        var client = new DriverCatalogClient(
            TestSupport.Wrap(TestSupport.CreateOptions()),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverCatalogClient>.Instance);

        var status = await client.GetStatusAsync(new DriverConnection
        {
            BaseUri = new Uri("http://127.0.0.1:33201/"),
            BearerToken = "secret",
            DiscoveryMethod = "explicit"
        });

        Assert.True(status.Ready);
        Assert.Equal("1.0.105", status.Build!.Version);
    }

    [Fact]
    public async Task GetOperations_ParsesSchemaVersion2()
    {
        var catalog = CatalogFixtures.ValidCatalog();
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            success = true,
            value = catalog
        }));

        var client = new DriverCatalogClient(
            TestSupport.Wrap(TestSupport.CreateOptions()),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverCatalogClient>.Instance);

        var result = await client.GetOperationsAsync(new DriverConnection
        {
            BaseUri = new Uri("http://127.0.0.1:33201/"),
            BearerToken = "secret",
            DiscoveryMethod = "explicit"
        });

        Assert.Equal(2, result.SchemaVersion);
        Assert.Equal(2, result.Operations.Count);
    }

    [Fact]
    public void CatalogCompatibility_RejectsUnsupportedSchema()
    {
        var catalog = CatalogFixtures.ValidCatalog();
        catalog.SchemaVersion = 1;
        var ex = Assert.Throws<DriverCatalogException>(() => CatalogCompatibility.Validate(catalog, 2));
        Assert.Contains("Unsupported catalog schemaVersion", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CatalogCompatibility_RejectsEmptyCatalog()
    {
        var catalog = CatalogFixtures.ValidCatalog();
        catalog.Operations.Clear();
        Assert.Throws<DriverCatalogException>(() => CatalogCompatibility.Validate(catalog, 2));
    }

    [Fact]
    public void CatalogCompatibility_RejectsDuplicateCanonicalNames()
    {
        var catalog = CatalogFixtures.ValidCatalog();
        catalog.Operations.Add(new Driver.Models.OperationDescriptorDto { Name = "CLICK" });
        var ex = Assert.Throws<DriverCatalogException>(() => CatalogCompatibility.Validate(catalog, 2));
        Assert.Contains("Duplicate canonical", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CatalogCompatibility_RejectsAliasCollisionWithCanonicalName()
    {
        var catalog = CatalogFixtures.ValidCatalog();
        catalog.Operations[0].Aliases.Add("launch");
        var ex = Assert.Throws<DriverCatalogException>(() => CatalogCompatibility.Validate(catalog, 2));
        Assert.Contains("collides with canonical", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task UnauthorizedStatus_ThrowsCatalogException()
    {
        var handler = new FakeHttpMessageHandler(_ =>
            new HttpResponseMessage(HttpStatusCode.Unauthorized));

        var client = new DriverCatalogClient(
            TestSupport.Wrap(TestSupport.CreateOptions()),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverCatalogClient>.Instance);

        var ex = await Assert.ThrowsAsync<DriverCatalogException>(() => client.GetStatusAsync(new DriverConnection
        {
            BaseUri = new Uri("http://127.0.0.1:33201/"),
            BearerToken = "bad",
            DiscoveryMethod = "explicit"
        }));

        Assert.Contains("401", ex.Message, StringComparison.Ordinal);
    }
}
