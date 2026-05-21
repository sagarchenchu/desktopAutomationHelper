using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for ComboBox / dropdown selection covering both huge-list (≥ 100 items)
/// and small-list (fewer than 100 items) paths.
///
/// All UiService interactions are mocked so no live Windows application is required.
/// The tests verify that:
/// <list type="bullet">
///   <item>Each strategy result is correctly forwarded through the controller.</item>
///   <item>The fallback chain for small lists (new paged/anchor-window visual strategies)
///         propagates success and failure the same way as the huge-list path.</item>
///   <item>The AllowKeyboardFallback flag is forwarded to the service unchanged.</item>
///   <item>Validation failures (missing value / locator) produce 400 responses.</item>
///   <item>Item-not-found failures produce 404 responses.</item>
/// </list>
/// </summary>
public class ComboBoxDropdownSelectionTests
{
    private readonly Mock<IUiService> _uiMock;
    private readonly UiController _controller;

    public ComboBoxDropdownSelectionTests()
    {
        _uiMock = new Mock<IUiService>(MockBehavior.Strict);

        var configMock = new Mock<IConfiguration>();
        configMock.Setup(c => c["FailureScreenshotDirectory"]).Returns((string?)null);

        _controller = new UiController(
            _uiMock.Object,
            NullLogger<UiController>.Instance,
            configMock.Object);
    }

    // -----------------------------------------------------------------------
    // Huge-list success paths
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_HugeList_PagedVisibleSearch_Returns200WithStrategy()
    {
        var expected = new
        {
            selected = "Item 150",
            actual = "Item 150",
            comboBox = "CountryDropdown",
            verified = true,
            strategy = "huge-list-paged-visible-search"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "CountryDropdown", "Item 150");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    [Fact]
    public void SelectComboBoxItem_HugeList_AnchorWindowSearch_Returns200WithStrategy()
    {
        var expected = new
        {
            selected = "Item 200",
            actual = "Item 200",
            comboBox = "CountryDropdown",
            verified = true,
            strategy = "huge-list-visible-anchor-window-search"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "CountryDropdown", "Item 200");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    [Fact]
    public void SelectComboBoxItem_HugeList_WithKeyboardFallback_TypeAheadSucceeds_Returns200()
    {
        var expected = new
        {
            selected = "Item 999",
            actual = "Item 999",
            comboBox = "HugeCombo",
            verified = true,
            strategy = "huge-list-explicit-typeahead-fallback"
        };
        _uiMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r =>
                r.Operation == "selectcomboboxitem" &&
                r.AllowKeyboardFallback == true &&
                r.Value == "Item 999")))
            .Returns(expected);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { Name = "HugeCombo" },
            Value = "Item 999",
            AllowKeyboardFallback = true
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    [Fact]
    public void SelectComboBoxItem_HugeList_DirectUia_Returns200WithStrategy()
    {
        var expected = new
        {
            selected = "Germany",
            actual = "Germany",
            comboBox = "CountryCombo",
            verified = true,
            strategy = "direct-uia-selectionitem"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "CountryCombo", "Germany");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    // -----------------------------------------------------------------------
    // Small-list success paths
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_SmallList_ExactVisibleCommit_Returns200WithStrategy()
    {
        var expected = new
        {
            selected = "Option A",
            actual = "Option A",
            comboBox = "SmallDropdown",
            verified = true,
            strategy = "small-combobox-exact-visible-commit"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "SmallDropdown", "Option A");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    [Fact]
    public void SelectComboBoxItem_SmallList_PagedVisibleSearchFallback_Returns200WithStrategy()
    {
        // Covers the new fallback added for small lists: when user-like commit fails
        // the paged visible search strategy takes over.
        var expected = new
        {
            selected = "Yes",
            actual = "Yes",
            comboBox = "TinyDropdown",
            verified = true,
            strategy = "small-combobox-paged-visible-search-fallback"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "TinyDropdown", "Yes");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    [Fact]
    public void SelectComboBoxItem_SmallList_AnchorWindowSearchFallback_Returns200WithStrategy()
    {
        // Covers the second new fallback for small lists: anchor-window visual search.
        var expected = new
        {
            selected = "No",
            actual = "No",
            comboBox = "TinyDropdown",
            verified = true,
            strategy = "small-combobox-anchor-window-search-fallback"
        };
        SetupExecute(expected);

        var result = Execute("selectcomboboxitem", "TinyDropdown", "No");

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);
        Assert.Equal(expected, response.Value);
    }

    // -----------------------------------------------------------------------
    // Error paths – item not found
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_HugeList_ItemNotFound_Returns404WithMessage()
    {
        const string msg =
            "Huge ComboBox item 'Ghost Item' was not found/verified by paged visible search or visible anchor-window search. " +
            "dropdownListDetected=True, expandedState=Expanded, currentValue='Item 1', visibleBatch='Item 1, Item 2'";
        SetupExecuteThrows(new InvalidOperationException(msg));

        var result = Execute("selectcomboboxitem", "HugeCombo", "Ghost Item");

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Contains("Ghost Item", response.Error);
    }

    [Fact]
    public void SelectComboBoxItem_SmallList_ItemNotFound_Returns404WithMessage()
    {
        const string msg = "ComboBox item 'MissingOption' was visible/found but did not commit. Actual='Option A'.";
        SetupExecuteThrows(new InvalidOperationException(msg));

        var result = Execute("selectcomboboxitem", "SmallDropdown", "MissingOption");

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Contains("MissingOption", response.Error);
    }

    [Fact]
    public void SelectComboBoxItem_SelectionAborted_Returns404()
    {
        SetupExecuteThrows(
            new InvalidOperationException("ComboBox selection aborted for 'Item X'."));

        var result = Execute("selectcomboboxitem", "SomeCombo", "Item X");

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Contains("aborted", response.Error);
    }

    // -----------------------------------------------------------------------
    // Error paths – validation
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_MissingValue_Returns400WithMessage()
    {
        SetupExecuteThrows(
            new ArgumentException("value is required for selectcomboboxitem."));

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { Name = "SomeCombo" }
        });

        var bad = Assert.IsType<BadRequestObjectResult>(result);
        var response = Assert.IsType<UiResponse>(bad.Value);
        Assert.False(response.Success);
        Assert.Contains("value is required", response.Error);
    }

    [Fact]
    public void SelectComboBoxItem_EmptyValue_Returns400WithMessage()
    {
        SetupExecuteThrows(
            new ArgumentException("selectcomboboxitem requires a non-empty value."));

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { Name = "SomeCombo" },
            Value = "   "
        });

        var bad = Assert.IsType<BadRequestObjectResult>(result);
        var response = Assert.IsType<UiResponse>(bad.Value);
        Assert.False(response.Success);
        Assert.Contains("non-empty value", response.Error);
    }

    // -----------------------------------------------------------------------
    // AllowKeyboardFallback flag
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_AllowKeyboardFallbackTrue_PropagatedToService()
    {
        _uiMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r =>
                r.Operation == "selectcomboboxitem" &&
                r.AllowKeyboardFallback == true)))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { AutomationId = "cbxCountry" },
            Value = "Australia",
            AllowKeyboardFallback = true
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);

        _uiMock.Verify(
            s => s.Execute(It.Is<UiRequest>(r => r.AllowKeyboardFallback == true)),
            Times.Once);
    }

    [Fact]
    public void SelectComboBoxItem_AllowKeyboardFallbackFalse_PropagatedToService()
    {
        _uiMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r =>
                r.Operation == "selectcomboboxitem" &&
                r.AllowKeyboardFallback == false)))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { AutomationId = "cbxCountry" },
            Value = "Australia",
            AllowKeyboardFallback = false
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);

        _uiMock.Verify(
            s => s.Execute(It.Is<UiRequest>(r => r.AllowKeyboardFallback == false)),
            Times.Once);
    }

    [Fact]
    public void SelectComboBoxItem_AllowKeyboardFallbackOmitted_DefaultsToNull()
    {
        _uiMock
            .Setup(s => s.Execute(It.Is<UiRequest>(r =>
                r.Operation == "selectcomboboxitem" &&
                r.AllowKeyboardFallback == null)))
            .Returns((object?)null);

        var result = _controller.Execute(new UiRequest
        {
            Operation = "selectcomboboxitem",
            Locator = new UiLocator { AutomationId = "cbxCountry" },
            Value = "Australia"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        var response = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(response.Success);

        _uiMock.Verify(
            s => s.Execute(It.Is<UiRequest>(r => r.AllowKeyboardFallback == null)),
            Times.Once);
    }

    // -----------------------------------------------------------------------
    // Screenshot on failure
    // -----------------------------------------------------------------------

    [Fact]
    public void SelectComboBoxItem_HugeList_ItemNotFound_IncludesScreenshotPath()
    {
        SetupExecuteThrows(new InvalidOperationException("Item not found in huge list."));
        _uiMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns(@"C:\screenshots\huge_fail.png");

        var result = Execute("selectcomboboxitem", "HugeCombo", "Ghost");

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Equal(@"C:\screenshots\huge_fail.png", response.ScreenshotPath);
    }

    [Fact]
    public void SelectComboBoxItem_SmallList_ItemNotFound_IncludesScreenshotPath()
    {
        SetupExecuteThrows(
            new InvalidOperationException("ComboBox item 'Missing' was visible/found but did not commit. Actual='X'."));
        _uiMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns(@"C:\screenshots\small_fail.png");

        var result = Execute("selectcomboboxitem", "SmallCombo", "Missing");

        var notFound = Assert.IsType<NotFoundObjectResult>(result);
        var response = Assert.IsType<UiResponse>(notFound.Value);
        Assert.False(response.Success);
        Assert.Equal(@"C:\screenshots\small_fail.png", response.ScreenshotPath);
    }

    // -----------------------------------------------------------------------
    // helpers
    // -----------------------------------------------------------------------

    private IActionResult Execute(string operation, string comboBoxName, string value)
        => _controller.Execute(new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { Name = comboBoxName },
            Value = value
        });

    private void SetupExecute(object? returnValue)
        => _uiMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>()))
            .Returns(returnValue);

    private void SetupExecuteThrows(Exception ex)
    {
        _uiMock
            .Setup(s => s.Execute(It.IsAny<UiRequest>()))
            .Throws(ex);
        _uiMock
            .Setup(s => s.TakeFailureScreenshot(It.IsAny<string>()))
            .Returns((string?)null);
    }
}
