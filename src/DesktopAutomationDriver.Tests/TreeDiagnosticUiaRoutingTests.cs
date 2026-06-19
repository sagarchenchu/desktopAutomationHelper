using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class TreeDiagnosticUiaRoutingTests
{
    [Fact]
    public void DumpUia_WithoutSession_FailsFastWithoutCallingService()
    {
        var treeMock = new Mock<INativeUiaTreeDiagnosticService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, treeMock: treeMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "dumpuia",
            Root = "activeWindow",
            View = "control",
            MaxDepth = 8,
            MaxChildren = 300,
            NameContains = "DQA",
            IncludePath = true,
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("no-active-window", json);
        treeMock.VerifyNoOtherCalls();
    }

    [Fact]
    public void FindUia_WithoutSession_WithRequestProcessId_CallsService()
    {
        var treeMock = new Mock<INativeUiaTreeDiagnosticService>();
        treeMock
            .Setup(s => s.FindElement(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new Dictionary<string, object?> { ["operation"] = "finduia", ["found"] = false });

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, treeMock: treeMock);

        service.Execute(new UiRequest
        {
            Operation = "finduia",
            NameContains = "DQA",
            View = "raw",
            Root = "processWindows",
            ProcessId = Environment.ProcessId,
            TimeoutMs = 5000
        });

        treeMock.Verify(
            s => s.FindElement(
                It.Is<UiRequest>(r => r.Operation == "finduia" && r.NameContains == "DQA"),
                null,
                Environment.ProcessId,
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void DumpUia_WithSession_CallsTreeDiagnosticService()
    {
        var treeMock = new Mock<INativeUiaTreeDiagnosticService>();
        treeMock
            .Setup(s => s.DumpTree(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new Dictionary<string, object?>
            {
                ["operation"] = "dumpuia",
                ["success"] = true,
                ["matchCount"] = 0
            });

        using var automation = new UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "tree-diag-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, treeMock: treeMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "dumpuia",
            View = "control",
            Root = "activeWindow",
            MaxDepth = 8,
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        treeMock.Verify(
            s => s.DumpTree(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void FindUia_WithoutCriteria_ReturnsInvalidRequestFromService()
    {
        var service = new NativeUiaTreeDiagnosticService(NullLogger<NativeUiaTreeDiagnosticService>.Instance);

        var result = service.FindElement(
            new UiRequest { Operation = "finduia", TimeoutMs = 1000 },
            IntPtr.Zero,
            Environment.ProcessId);

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("invalid-request", json);
    }
}
