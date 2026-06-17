using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

internal sealed class NativeUiaResolveResult
{
    public IUIAutomationElement? Element { get; init; }
    public bool IsAmbiguous { get; init; }
    public List<object> Candidates { get; init; } = new();
    public string? Stage { get; init; }
    public string? LastError { get; init; }
}

internal sealed class NativeUiaElementResolver
{
    private const int ComboBoxControlTypeId = 50003;
    private const int EditControlTypeId = 50004;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaElementFinder _finder;

    public NativeUiaElementResolver(NativeUiaAutomation uia)
    {
        _uia = uia;
        _finder = new NativeUiaElementFinder(_uia);
    }

    public NativeUiaResolveResult ResolveComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var locator = ToNativeLocator(request, processId);
        var searchRoots = BuildSearchRoots(activeWindowHwnd, processId);
        var allMatches = new List<IUIAutomationElement>();

        foreach (var root in searchRoots)
        {
            allMatches.AddRange(_finder.FindElements(root, locator.AsComboBoxLocator()));

            if (allMatches.Count == 0 &&
                (!string.IsNullOrWhiteSpace(locator.AutomationId) || !string.IsNullOrWhiteSpace(locator.Name)))
            {
                foreach (var relaxed in _finder.FindElements(root, locator.WithoutControlType()))
                {
                    var promoted = PromoteToComboBox(relaxed);
                    if (promoted != null)
                        allMatches.Add(promoted);
                }
            }
        }

        allMatches = allMatches.DistinctBy(RuntimeKey).ToList();
        var candidates = allMatches
            .Select((element, index) => NativeUiaDiagnostics.CandidateDiagnostic(
                index,
                _uia.CreateSnapshot(element)))
            .Cast<object>()
            .ToList();

        if (allMatches.Count == 0)
        {
            return new NativeUiaResolveResult
            {
                Stage = "combo-not-found",
                LastError = "Native UIA resolver could not find a ComboBox for the locator.",
                Candidates = candidates
            };
        }

        var disambiguationIndex = request.FoundIndex;
        if (allMatches.Count > 1 && !disambiguationIndex.HasValue)
        {
            return new NativeUiaResolveResult
            {
                IsAmbiguous = true,
                Stage = "ambiguous-combobox",
                LastError = $"Found {allMatches.Count} ComboBox candidates. Provide index or foundIndex.",
                Candidates = candidates
            };
        }

        if (disambiguationIndex.HasValue)
        {
            if (disambiguationIndex.Value < 0 || disambiguationIndex.Value >= allMatches.Count)
            {
                return new NativeUiaResolveResult
                {
                    Stage = "index-out-of-range",
                    LastError = $"ComboBox index {disambiguationIndex.Value} is out of range (count={allMatches.Count}).",
                    Candidates = candidates
                };
            }

            return new NativeUiaResolveResult
            {
                Element = allMatches[disambiguationIndex.Value],
                Candidates = candidates
            };
        }

        return new NativeUiaResolveResult
        {
            Element = allMatches[0],
            Candidates = candidates
        };
    }

    private IUIAutomationElement? PromoteToComboBox(IUIAutomationElement element)
    {
        var controlType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
        if (controlType == ComboBoxControlTypeId)
            return element;

        if (controlType == EditControlTypeId)
        {
            return _uia.WalkAncestor(
                element,
                e => _uia.GetIntProperty(e, UIA_PropertyIds.UIA_ControlTypePropertyId) == ComboBoxControlTypeId);
        }

        return null;
    }

    private List<IUIAutomationElement> BuildSearchRoots(IntPtr? activeWindowHwnd, int? processId)
    {
        var roots = new List<IUIAutomationElement>();

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        var foreground = NativeUiaInput.ForegroundWindowHandle();
        if (foreground != IntPtr.Zero)
        {
            var fgRoot = _uia.FromHandle(foreground);
            if (fgRoot != null && !roots.Any(r => SameElement(r, fgRoot)))
                roots.Add(fgRoot);
        }

        if (processId.HasValue)
        {
            var sameProcess = _uia.FindAllDescendants(
                _uia.Root,
                _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                200);
            roots.AddRange(sameProcess.Where(e =>
                _uia.GetIntProperty(e, UIA_PropertyIds.UIA_ControlTypePropertyId) == 50032));
        }

        roots.Add(_uia.Root);
        return roots.DistinctBy(RuntimeKey).ToList();
    }

    private static NativeUiaLocator ToNativeLocator(UiRequest request, int? processId) =>
        new()
        {
            Name = request.Locator?.Name,
            AutomationId = request.Locator?.AutomationId,
            ClassName = request.Locator?.ClassName ?? request.ClassName,
            ControlType = request.Locator?.ControlType,
            Hwnd = request.Hwnd ?? request.Locator?.Hwnd,
            ProcessId = request.ProcessId ?? request.Locator?.ProcessId ?? processId,
            FoundIndex = request.FoundIndex,
            MatchMode = string.IsNullOrWhiteSpace(request.MatchMode) ? "exact" : request.MatchMode!
        };

    private static bool SameElement(IUIAutomationElement left, IUIAutomationElement right) =>
        string.Equals(RuntimeKey(left), RuntimeKey(right), StringComparison.Ordinal);

    private static string RuntimeKey(IUIAutomationElement element)
    {
        try
        {
            var runtimeId = element.GetRuntimeId();
            return runtimeId == null ? element.GetHashCode().ToString() : string.Join(".", runtimeId);
        }
        catch
        {
            return Guid.NewGuid().ToString("N");
        }
    }
}
