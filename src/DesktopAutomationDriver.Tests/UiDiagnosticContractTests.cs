using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Resolver;
using DesktopAutomationDriver.Models.Response;
using ResolvedElement = DesktopAutomationDriver.Models.Resolver.ResolvedElement;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class UiDiagnosticContractTests
{
    private readonly Mock<IUiService> _ui = new();
    private readonly UiController _controller;
    private readonly DefaultHttpContext _http;

    public UiDiagnosticContractTests()
    {
        var config = new Mock<IConfiguration>();
        config.Setup(c => c["FailureScreenshotDirectory"]).Returns((string?)null);

        _http = new DefaultHttpContext();
        _controller = new UiController(
            _ui.Object,
            new UiOperationCatalog(),
            NullLogger<UiController>.Instance,
            config.Object)
        {
            ControllerContext = new ControllerContext { HttpContext = _http }
        };
    }

    [Theory]
    [InlineData("dumpuia")]
    [InlineData("finduia")]
    [InlineData("findlocator")]
    [InlineData("findall")]
    [InlineData("dumptree")]
    public void Diagnostic_RequestIsPassedThroughUnmodified(string operation)
    {
        var request = new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu" },
            View = "raw",
            MaxDepth = 4
        };

        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Returns(new { operation, success = true, matchCount = 1 });

        var result = _controller.Execute(request);
        Assert.IsType<OkObjectResult>(result);

        _ui.Verify(s => s.Execute(
            It.Is<UiRequest>(r =>
                ReferenceEquals(r, request)
                && r.Operation == operation
                && r.View == "raw"
                && r.MaxDepth == 4
                && r.Locator!.Name == "DQA"),
            _http.RequestAborted), Times.Once);
    }

    [Fact]
    public void Diagnostic_SuccessPayload_RemainsInUiResponseValue()
    {
        var payload = new { operation = "dumpuia", success = true, nodes = new[] { "a" } };
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns(payload);

        var result = _controller.Execute(new UiRequest { Operation = "dumpuia" });
        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Same(payload, response.Value);
    }

    [Fact]
    public void Diagnostic_PayloadSuccessFalse_EnvelopeSuccessFalse_PreservesReason()
    {
        var payload = new
        {
            operation = "finduia",
            success = false,
            reason = "element-not-found",
            message = "not found",
            candidates = new[] { new { name = "Near" } },
            suggestions = new[] { "try raw view" }
        };

        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>())).Returns(payload);

        var result = _controller.Execute(new UiRequest { Operation = "finduia", Locator = new UiLocator { Name = "x" } });
        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.False(response.Success);
        Assert.Equal("element-not-found", response.Reason);
        Assert.Same(payload, response.Value);
    }

    [Fact]
    public void Diagnostic_UiResolutionException_Returns404WithDetailsAndScreenshot()
    {
        var locator = new UiLocator { AutomationId = "x" };
        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new UiResolutionException(
                "no-match",
                "not found",
                locator,
                Array.Empty<ResolvedElement>(),
                new[] { "widen search" }));
        _ui.Setup(s => s.TakeFailureScreenshot(It.IsAny<string>())).Returns("C:/shots/fail.png");

        var result = _controller.Execute(new UiRequest { Operation = "findlocator", Locator = locator });
        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Equal("no-match", response.Reason);
        Assert.Equal("C:/shots/fail.png", response.ScreenshotPath);
        Assert.NotNull(response.Candidates);
        Assert.NotNull(response.Suggestions);
        Assert.Same(locator, response.Locator);
    }

    [Fact]
    public void Diagnostic_CancellationToken_IsRequestAborted()
    {
        using var cts = new CancellationTokenSource();
        _http.RequestAborted = cts.Token;

        _ui.Setup(s => s.Execute(It.IsAny<UiRequest>(), cts.Token))
            .Returns(new { operation = "dumptree", success = true });

        _controller.Execute(new UiRequest { Operation = "dumptree" });
        _ui.Verify(s => s.Execute(It.IsAny<UiRequest>(), cts.Token), Times.Once);
    }
}
