using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Readiness;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace DesktopAutomationAgent.Tests;

internal static class TestSupport
{
    public static AgentOptions CreateOptions(
        string? workspaceRoot = null,
        string? baseUrl = null,
        string? bearerToken = null,
        string? verifyUrl = "http://localhost:9102/verify",
        bool allowRemote = false)
    {
        return new AgentOptions
        {
            Driver = new DriverOptions
            {
                BaseUrl = baseUrl,
                BearerToken = bearerToken,
                VerifyUrl = verifyUrl ?? "http://localhost:9102/verify",
                AllowRemoteDriver = allowRemote,
                RequestTimeoutSeconds = 5,
                ExpectedCatalogSchemaVersion = 2
            },
            Workspace = new WorkspaceOptions
            {
                Root = workspaceRoot ?? Path.Combine(Path.GetTempPath(), "da-agent-tests", Guid.NewGuid().ToString("N"))
            },
            Suites = new SuiteOptions()
        };
    }

    public static IOptions<AgentOptions> Wrap(AgentOptions options) => Options.Create(options);

    public static WorkspaceManager CreateWorkspace(AgentOptions options) =>
        new(Wrap(options), NullLogger<WorkspaceManager>.Instance);

    public static SuiteManifestReader CreateSuiteReader(AgentOptions options, IWorkspaceManager? workspace = null) =>
        new(Wrap(options), workspace ?? CreateWorkspace(options));

    public static IHttpClientFactory CreateFactory(FakeHttpMessageHandler handler)
    {
        var mock = new Mock<IHttpClientFactory>();
        mock.Setup(f => f.CreateClient(It.IsAny<string>()))
            .Returns(() => new HttpClient(handler, disposeHandler: false));
        return mock.Object;
    }
}

internal sealed class FakeHttpMessageHandler : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _handler;

    public List<HttpRequestMessage> Requests { get; } = [];

    public FakeHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> handler)
        : this((req, _) => Task.FromResult(handler(req)))
    {
    }

    public FakeHttpMessageHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> handler)
    {
        _handler = handler;
    }

    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        Requests.Add(request);
        return await _handler(request, cancellationToken).ConfigureAwait(false);
    }

    public static HttpResponseMessage Json(object payload, HttpStatusCode code = HttpStatusCode.OK)
    {
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
        return new HttpResponseMessage(code)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
    }
}

internal static class CatalogFixtures
{
    public static OperationsCatalogDto ValidCatalog() => new()
    {
        SchemaVersion = 2,
        DriverVersion = "1.0.105",
        Operations =
        [
            new OperationDescriptorDto
            {
                Name = "click",
                Aliases = [],
                Category = "element-action",
                OperationType = "action"
            },
            new OperationDescriptorDto
            {
                Name = "launch",
                Aliases = ["start"],
                Category = "session-window",
                OperationType = "session"
            }
        ]
    };
}
