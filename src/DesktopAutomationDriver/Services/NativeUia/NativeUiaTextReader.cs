using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

internal sealed class NativeUiaTextReader
{
    private readonly NativeUiaAutomation _uia;

    public NativeUiaTextReader(NativeUiaAutomation uia) => _uia = uia;

    public string ReadElementText(IUIAutomationElement element) => _uia.GetElementText(element);

    public List<string> ReadCandidateTexts(IUIAutomationElement element) =>
        new List<string>
        {
            _uia.GetStringProperty(element, UIA_PropertyIds.UIA_NamePropertyId),
            _uia.GetValuePatternText(element),
            _uia.GetTextPatternText(element),
            _uia.GetLegacyAccessibleName(element),
            _uia.GetLegacyAccessibleValue(element),
            _uia.GetStringProperty(element, UIA_PropertyIds.UIA_AutomationIdPropertyId)
        }.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    public string ReadComboBoxValue(IUIAutomationElement combo)
    {
        var valuePattern = _uia.GetValuePatternText(combo);
        if (!string.IsNullOrWhiteSpace(valuePattern))
            return valuePattern;

        var legacyValue = _uia.GetLegacyAccessibleValue(combo);
        if (!string.IsNullOrWhiteSpace(legacyValue))
            return legacyValue;

        var name = _uia.GetStringProperty(combo, UIA_PropertyIds.UIA_NamePropertyId);
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        var selectedItem = FindSelectedItem(combo);
        if (selectedItem != null)
            return ReadElementText(selectedItem);

        var innerEdit = _uia.FindFirstDescendant(
            combo,
            _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50004));
        if (innerEdit != null)
        {
            var editValue = _uia.GetValuePatternText(innerEdit);
            if (!string.IsNullOrWhiteSpace(editValue))
                return editValue;
        }

        return ReadElementText(combo);
    }

    private IUIAutomationElement? FindSelectedItem(IUIAutomationElement combo)
    {
        foreach (var item in _uia.FindAllDescendants(
                     combo,
                     _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50007),
                     100))
        {
            if (_uia.TryGetSelectionItemPattern(item, out var selection) && selection!.CurrentIsSelected != 0)
                return item;
        }

        return null;
    }
}
