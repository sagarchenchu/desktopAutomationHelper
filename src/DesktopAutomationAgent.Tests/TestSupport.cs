using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Plans;
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
        bool allowRemote = false,
        int maxPlanBytes = 1_048_576)
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
            Suites = new SuiteOptions(),
            Runner = new RunnerOptions
            {
                MaxPlanBytes = maxPlanBytes,
                MaxResponseBytes = 1_048_576,
                StepTransportTimeoutSeconds = 5,
                CleanupTimeoutSeconds = 2,
                RegexTimeoutMilliseconds = 50
            },
            ObjectRepository = new ObjectRepositoryOptions
            {
                MaxFileBytes = 5_242_880,
                MaxPages = 500,
                MaxElementsPerPage = 5000,
                MaxTotalElements = 50_000,
                DiagnosticTimeoutMilliseconds = 15_000
            }
        };
    }

    public static IOptions<AgentOptions> Wrap(AgentOptions options) => Options.Create(options);

    public static WorkspaceManager CreateWorkspace(AgentOptions options) =>
        new(Wrap(options), NullLogger<WorkspaceManager>.Instance);

    public static SuiteManifestReader CreateSuiteReader(AgentOptions options, IWorkspaceManager? workspace = null) =>
        new(Wrap(options), workspace ?? CreateWorkspace(options));

    public static PlanManifestReader CreatePlanReader(AgentOptions options, IWorkspaceManager? workspace = null) =>
        new(Wrap(options), workspace ?? CreateWorkspace(options));

    public static IHttpClientFactory CreateFactory(FakeHttpMessageHandler handler)
    {
        var mock = new Mock<IHttpClientFactory>();
        mock.Setup(f => f.CreateClient(It.IsAny<string>()))
            .Returns(() => new HttpClient(handler, disposeHandler: false));
        return mock.Object;
    }

    public static string WritePlan(AgentOptions options, string fileName, string json)
    {
        var workspace = CreateWorkspace(options);
        workspace.Initialize();
        var path = Path.Combine(workspace.RootPath, "plans", fileName);
        File.WriteAllText(path, json);
        return path;
    }

    public static string MinimalPlanJson(
        string planId = "SAMPLE-1",
        string operation = "listwindows",
        string? extraTopLevel = null,
        string? stepsOverride = null,
        string? onFailure = null)
    {
        var steps = stepsOverride ??
            $$"""
              [
                {
                  "id": "step-1",
                  "operation": "{{operation}}",
                  "arguments": {}
                }
              ]
              """;

        var onFailureBlock = onFailure is null
            ? string.Empty
            : $",\n  \"onFailureSteps\": {onFailure}";

        var extra = extraTopLevel is null ? string.Empty : ",\n  " + extraTopLevel.Trim().TrimStart(',');

        return $$"""
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "{{planId}}",
              "name": "Sample plan"{{extra}},
              "steps": {{steps}}{{onFailureBlock}}
            }
            """;
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
        // Clone request content so tests can re-read bodies later.
        if (request.Content is not null)
        {
            var bytes = await request.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
            var clone = new HttpRequestMessage(request.Method, request.RequestUri);
            foreach (var header in request.Headers)
                clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
            clone.Content = new ByteArrayContent(bytes);
            foreach (var header in request.Content.Headers)
                clone.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
            Requests.Add(clone);
        }
        else
        {
            Requests.Add(request);
        }

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

    public static HttpResponseMessage Text(string body, HttpStatusCode code = HttpStatusCode.OK) =>
        new(code) { Content = new StringContent(body, Encoding.UTF8, "application/json") };
}

internal static class CatalogFixtures
{
    /// <summary>Minimal Phase 1-compatible catalog used by existing readiness tests.</summary>
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

    public static OperationsCatalogDto Phase2Catalog() => new()
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
                OperationType = "action",
                RequiresSession = true,
                RequiredInputs = ["locator"]
            },
            new OperationDescriptorDto
            {
                Name = "launch",
                Aliases = ["start"],
                DeprecatedAliases = ["runapp"],
                Category = "session-window",
                OperationType = "session",
                RequiresSession = false,
                RequiredInputs = ["value"]
            },
            new OperationDescriptorDto
            {
                Name = "quit",
                Aliases = [],
                Category = "session-window",
                OperationType = "session",
                RequiresSession = true
            },
            new OperationDescriptorDto
            {
                Name = "close",
                Aliases = [],
                Category = "session-window",
                OperationType = "session",
                RequiresSession = true
            },
            new OperationDescriptorDto
            {
                Name = "closewindow",
                Aliases = [],
                Category = "session-window",
                OperationType = "action",
                RequiresSession = true
            },
            new OperationDescriptorDto
            {
                Name = "listwindows",
                Aliases = [],
                Category = "query",
                OperationType = "query",
                RequiresSession = false
            },
            new OperationDescriptorDto
            {
                Name = "gettext",
                Aliases = [],
                Category = "element-query",
                OperationType = "query",
                RequiresSession = true,
                RequiredInputs = ["locator"]
            },
            new OperationDescriptorDto
            {
                Name = "setvalue",
                Aliases = [],
                Category = "element-action",
                OperationType = "action",
                RequiresSession = true,
                RequiredInputs = [],
                RequiredInputAlternatives =
                [
                    ["locator", "value"],
                    ["automationId", "value"]
                ]
            },
            new OperationDescriptorDto
            {
                Name = "legacyop",
                Aliases = [],
                Category = "misc",
                OperationType = "action",
                RequiresSession = false,
                Deprecated = true
            },
            new OperationDescriptorDto
            {
                Name = "dumpuia",
                Aliases = [],
                Category = "native-uia",
                OperationType = "diagnostic",
                RequiresSession = false
            },
            new OperationDescriptorDto
            {
                Name = "finduia",
                Aliases = [],
                Category = "native-uia",
                OperationType = "diagnostic",
                RequiresSession = false
            }
        ]
    };
}
