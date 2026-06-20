using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="UiController"/> covering dropdown/select operation paths.
/// All UiService interactions are mocked so no live Windows application is required.
/// </summary>
public class UiControllerTests
{
    private readonly Mock<IUiService> _uiServiceMock;
    private readonly UiController _controller;

    public UiControllerTests()
    {
        _uiServiceMock = new Mock<IUiService>();

        var configMock = new Mock<IConfiguration>();
        configMock.Setup(c => c["FailureScreenshotDirectory"]).Returns((string?)null);

        _controller = new UiController(
            _uiServiceMock.Object,
            NullLogger<UiController>.Instance,
            configMock.Object);
        _controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
    }

    // ------------------------------------------------------------------
    // Null / missing request
    // ------------------------------------------------------------------

    [Fact]
    public void Execute_WhenRequestIsNull_Returns400()
    {
        var result = _controller.Execute(null);

        var bad = Assert.IsType<BadRequestObjectResult>(result);
        var response = Assert.IsType<UiResponse>(bad.Value);
        Assert.False(response.Success);
    }

    // ------------------------------------------------------------------
    // Dropdown / ComboBox – success paths
    // ------------------------------------------------------------------

    [Fact]
    public void Execute_WhenNativeUiaPayloadReportsFailure_EnvelopeSuccessFalse()
    {
        var payload = new
        {
            operation = "clickmenuuia",
            success = false,
            reason = "element-not-found",
            message = "Native UIA resolver could not find an element for the locator."
        };

        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "clickmenuuia"), It.IsAny<CancellationToken>()))
            .Returns(payload);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "clickmenuuia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu" }
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.False(response.Success);
        Assert.Equal(payload, response.Value);
        Assert.Equal("element-not-found", response.Reason);
    }

    [Fact]
    public void Execute_SelectOperation_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "select"), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "MyComboBox" },
            Value = "Option A"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    [Fact]
    public void Execute_SelectAidOperation_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "selectaid"), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectaid",
            Locator = new UiLocator { AutomationId = "cbxStatus" },
            Value = "StatusActive"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    [Fact]
    public void Execute_SelectComboBoxItemOperation_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "selectcomboboxitem"), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { AutomationId = "cbxCountry" },
            Value = "Canada"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    [Fact]
    public void Execute_TypeAndSelectOperation_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "typeandselect"), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "typeandselect",
            Locator = new UiLocator { Name = "SearchBox" },
            Value = "Can"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    [Fact]
    public void Execute_GetSelectedOperation_Returns200WithSelectedValue()
    {
        var expectedResult = new { selected = "Option B" };
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "getselected"), It.IsAny<CancellationToken>()))
            .Returns(expectedResult);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "getselected",
            Locator = new UiLocator { Name = "MyComboBox" }
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expectedResult, response.Value);
    }

    [Fact]
    public void Execute_SelectByIndexOperation_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r => r.Operation == "select" && r.Index == 2), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "MyComboBox" },
            Index = 2
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    [Fact]
    public void Execute_SelectWithKeyboardFallback_Returns200WithSuccess()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r =>
                r.Operation == "select" && r.AllowKeyboardFallback == true), It.IsAny<CancellationToken>()))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "HugeComboBox" },
            Value = "Item 999",
            AllowKeyboardFallback = true
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
    }

    // ------------------------------------------------------------------
    // Error propagation paths
    // ------------------------------------------------------------------

    [Fact]
    public void Execute_WhenInvalidOperationExceptionThrown_Returns404()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new InvalidOperationException("ComboBox item not found."));
        _uiServiceMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns((string?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "MyComboBox" },
            Value = "NonExistentItem"
        });

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Contains("ComboBox item not found", response.Error);
    }

    [Fact]
    public void Execute_WhenArgumentExceptionThrown_Returns400()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new ArgumentException("Locator is required for 'select'."));
        _uiServiceMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns((string?)null);

        var result = _controller.Execute(new UiRequest { Operation = "select" });

        var bad = Assert.IsType<BadRequestObjectResult>(result);
        var response = Assert.IsType<UiResponse>(bad.Value);
        Assert.False(response.Success);
        Assert.Contains("Locator is required", response.Error);
    }

    [Fact]
    public void Execute_WhenUnexpectedExceptionThrown_Returns500()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new Exception("Unexpected failure"));
        _uiServiceMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns((string?)null);

        var result = _controller.Execute(new UiRequest { Operation = "select" });

        var obj = Assert.IsType<ObjectResult>(result);
        Assert.Equal(500, obj.StatusCode);
        var response = Assert.IsType<UiResponse>(obj.Value);
        Assert.False(response.Success);
    }

    [Fact]
    public void Execute_WhenExceptionThrownAndScreenshotAvailable_ResponseIncludesScreenshotPath()
    {
        _uiServiceMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>(), It.IsAny<CancellationToken>()))
            .Throws(new InvalidOperationException("Not found."));
        _uiServiceMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns(@"C:\screenshots\failure.png");

        var result = _controller.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "MyComboBox" },
            Value = "Missing"
        });

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Equal(@"C:\screenshots\failure.png", response.ScreenshotPath);
    }

    // ------------------------------------------------------------------
    // UiRequest model – dropdown-related property defaults
    // ------------------------------------------------------------------

    [Fact]
    public void UiRequest_AllowKeyboardFallback_DefaultsToNull()
    {
        var req = new UiRequest();
        Assert.Null(req.AllowKeyboardFallback);
    }

    [Fact]
    public void UiRequest_Index_DefaultsToNull()
    {
        var req = new UiRequest();
        Assert.Null(req.Index);
    }

    [Fact]
    public void UiRequest_Value_DefaultsToNull()
    {
        var req = new UiRequest();
        Assert.Null(req.Value);
    }

    [Fact]
    public void UiRequest_ClickRegion_DefaultsToNull()
    {
        var req = new UiRequest();
        Assert.Null(req.ClickRegion);
    }

    [Fact]
    public void UiRequest_Limit_DefaultsToNull()
    {
        var req = new UiRequest();
        Assert.Null(req.Limit);
    }
}
