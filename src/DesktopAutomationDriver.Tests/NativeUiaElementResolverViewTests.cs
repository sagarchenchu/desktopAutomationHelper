using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaElementResolverViewTests
{
    [Fact]
    public void InferDefaultView_ClickMenuUia_DefaultsToRaw()
    {
        var view = NativeUiaElementResolver.InferDefaultView(new UiRequest
        {
            Operation = "clickmenuuia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Button" }
        });

        Assert.Equal("raw", view);
    }

    [Fact]
    public void InferDefaultView_MenuControlType_DefaultsToRaw()
    {
        var view = NativeUiaElementResolver.InferDefaultView(new UiRequest
        {
            Operation = "clickuia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu" }
        });

        Assert.Equal("raw", view);
    }

    [Fact]
    public void InferDefaultView_ExplicitViewOverridesDefault()
    {
        var view = new UiRequest
        {
            Operation = "clickmenuuia",
            View = "control",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu" }
        };

        Assert.Equal("control", view.View ?? view.TreeView ?? NativeUiaElementResolver.InferDefaultView(view));
    }

    [Fact]
    public void InferDefaultView_ComboBoxOperations_DefaultToRaw()
    {
        Assert.Equal(
            "raw",
            NativeUiaElementResolver.InferDefaultView(new UiRequest
            {
                Operation = "findcomboboxuia",
                Locator = new UiLocator { AutomationId = "cmbTest", ControlType = "ComboBox" }
            }));

        Assert.Equal(
            "raw",
            NativeUiaElementResolver.InferDefaultView(new UiRequest
            {
                Operation = "selectcomboboxuia",
                Locator = new UiLocator { AutomationId = "cmbTest", ControlType = "ComboBox" }
            }));
    }

    [Fact]
    public void InferDefaultView_GenericButton_DefaultsToControl()
    {
        var view = NativeUiaElementResolver.InferDefaultView(new UiRequest
        {
            Operation = "clickuia",
            Locator = new UiLocator { AutomationId = "btnOk", ControlType = "Button" }
        });

        Assert.Equal("control", view);
    }
}
