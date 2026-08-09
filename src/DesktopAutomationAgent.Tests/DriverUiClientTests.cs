using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class DriverUiClientTests
{
    [Fact]
    public async Task SendsPostUiWithBearerAndFlattenedBody()
    {
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            success = true,
            value = "ok"
        }));
        var client = CreateClient(handler);
        var step = new PlanStep
        {
            Id = "click-1",
            Operation = "click",
            Arguments = new Dictionary<string, JsonElement>
            {
                ["locator"] = JsonSerializer.SerializeToElement(new { automationId = "Submit" }),
                ["timeoutMs"] = JsonSerializer.SerializeToElement(20000)
            }
        };

        var response = await client.ExecuteStepAsync(Connection(), step);
        Assert.True(response.Success);

        var request = Assert.Single(handler.Requests);
        Assert.Equal(HttpMethod.Post, request.Method);
        Assert.Equal("/ui", request.RequestUri!.AbsolutePath);
        Assert.Equal("Bearer", request.Headers.Authorization!.Scheme);
        Assert.Equal("test-token", request.Headers.Authorization.Parameter);

        var body = await request.Content!.ReadAsStringAsync();
        Assert.Contains("\"operation\":\"click\"", body);
        Assert.Contains("\"locator\"", body);
        Assert.DoesNotContain("\"arguments\"", body);
        Assert.DoesNotContain("test-token", body);
    }

    [Fact]
    public async Task HandlesSuccessFalseAsOperationFailure()
    {
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            success = false,
            error = "not found",
            reason = "ElementNotFound"
        }));
        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.OperationFailure, ex.Classification);
    }

    [Theory]
    [InlineData(HttpStatusCode.Unauthorized, UiFailureClassification.Authentication)]
    [InlineData(HttpStatusCode.Forbidden, UiFailureClassification.Authentication)]
    public async Task HandlesAuthFailures(HttpStatusCode code, UiFailureClassification expected)
    {
        var handler = new FakeHttpMessageHandler(_ => new HttpResponseMessage(code)
        {
            Content = new StringContent("{}", Encoding.UTF8, "application/json")
        });
        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(expected, ex.Classification);
    }

    [Theory]
    [InlineData(HttpStatusCode.BadRequest)]
    [InlineData(HttpStatusCode.NotFound)]
    [InlineData(HttpStatusCode.Conflict)]
    [InlineData(HttpStatusCode.InternalServerError)]
    public async Task HandlesOperationHttpFailures(HttpStatusCode code)
    {
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            success = false,
            error = "boom",
            reason = "diagnostic"
        }, code));
        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.OperationFailure, ex.Classification);
        Assert.Equal((int)code, ex.Response!.HttpStatusCode);
    }

    [Fact]
    public async Task HandlesMalformedJsonAndEmptyBody()
    {
        var bad = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Text("{not-json"));
        var ex1 = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(bad).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.OperationFailure, ex1.Classification);

        var empty = new FakeHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(string.Empty)
        });
        var ex2 = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(empty).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.OperationFailure, ex2.Classification);
    }

    [Fact]
    public async Task HandlesTransportFailure()
    {
        var handler = new FakeHttpMessageHandler((_, _) => throw new HttpRequestException("refused"));
        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.DriverUnavailable, ex.Classification);
    }

    [Fact]
    public async Task HandlesTimeout()
    {
        var handler = new FakeHttpMessageHandler(async (_, ct) =>
        {
            await Task.Delay(TimeSpan.FromSeconds(30), ct);
            return FakeHttpMessageHandler.Json(new { success = true });
        });
        var options = TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: "test-token");
        options.Runner.StepTransportTimeoutSeconds = 1;
        var client = new DriverUiClient(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverUiClient>.Instance);

        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            client.ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(UiFailureClassification.ExecutionTimeout, ex.Classification);
    }

    [Fact]
    public async Task RejectsOversizedResponse()
    {
        var options = TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: "test-token");
        options.Runner.MaxResponseBytes = 32;
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Text(new string('x', 128)));
        var client = new DriverUiClient(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverUiClient>.Instance);

        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            client.ExecuteStepAsync(Connection(), Step()));
        Assert.Contains("maximum size", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task HonorsCancellation()
    {
        var handler = new FakeHttpMessageHandler(async (_, ct) =>
        {
            await Task.Delay(TimeSpan.FromSeconds(30), ct);
            return FakeHttpMessageHandler.Json(new { success = true });
        });
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var ex = await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step(), cts.Token));
        Assert.Equal(UiFailureClassification.Cancelled, ex.Classification);
    }

    [Fact]
    public async Task DoesNotRetry()
    {
        var calls = 0;
        var handler = new FakeHttpMessageHandler(_ =>
        {
            calls++;
            return FakeHttpMessageHandler.Json(new { success = false, error = "fail" });
        });
        await Assert.ThrowsAsync<UiExecutionException>(() =>
            CreateClient(handler).ExecuteStepAsync(Connection(), Step()));
        Assert.Equal(1, calls);
        Assert.Single(handler.Requests);
    }

    private static DriverUiClient CreateClient(FakeHttpMessageHandler handler)
    {
        var options = TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33201", bearerToken: "test-token");
        return new DriverUiClient(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverUiClient>.Instance);
    }

    private static DriverConnection Connection() =>
        new()
        {
            BaseUri = new Uri("http://127.0.0.1:33201/"),
            BearerToken = "test-token",
            DiscoveryMethod = "explicit"
        };

    private static PlanStep Step() =>
        new()
        {
            Id = "s1",
            Operation = "listwindows",
            Arguments = new Dictionary<string, JsonElement>()
        };
}
