using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaLocatorTests
{
    [Fact]
    public void AsComboBoxLocator_ForcesComboBoxControlType()
    {
        var locator = new NativeUiaLocator
        {
            AutomationId = "cmbinbound",
            ControlType = "Edit"
        };

        var comboLocator = locator.AsComboBoxLocator();

        Assert.Equal("cmbinbound", comboLocator.AutomationId);
        Assert.Equal("ComboBox", comboLocator.ControlType);
    }

    [Fact]
    public void WithoutControlType_PreservesAutomationIdForRelaxedSearch()
    {
        var locator = new NativeUiaLocator
        {
            AutomationId = "cmbinbound",
            ControlType = "ComboBox",
            ProcessId = 1234
        };

        var relaxed = locator.WithoutControlType();

        Assert.Equal("cmbinbound", relaxed.AutomationId);
        Assert.Null(relaxed.ControlType);
        Assert.Equal(1234, relaxed.ProcessId);
    }
}
