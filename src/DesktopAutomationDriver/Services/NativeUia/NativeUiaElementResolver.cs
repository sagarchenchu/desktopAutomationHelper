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
    private const int MaxSearchLimit = 200;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaElementFinder _finder;

    public NativeUiaElementResolver(NativeUiaAutomation uia)
    {
        _uia = uia;
        _finder = new NativeUiaElementFinder(uia);
    }

    public NativeUiaResolveResult ResolveComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var locator = ToNativeLocator(request, processId);
        if (locator.Hwnd is > 0)
        {
            var fromHandle = _uia.FromHandle(new IntPtr(locator.Hwnd.Value));
            if (fromHandle != null)
            {
                var promoted = PromoteToComboBox(fromHandle) ?? fromHandle;
                if (ElementMatchesComboLocator(promoted, locator))
                    return SingleResult(promoted);
            }
        }

        var allMatches = new List<IUIAutomationElement>();
        foreach (var root in BuildSearchRoots(activeWindowHwnd, processId, includeDesktop: false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            allMatches.AddRange(FindComboMatches(root, locator));
        }

        if (allMatches.Count == 0)
        {
            foreach (var root in BuildSearchRoots(activeWindowHwnd, processId, includeDesktop: true))
            {
                cancellationToken.ThrowIfCancellationRequested();
                allMatches.AddRange(FindComboMatches(root, locator));
            }
        }

        allMatches = allMatches.DistinctBy(RuntimeKey).ToList();
        var candidates = allMatches
            .Select((element, index) => NativeUiaDiagnostics.CandidateDiagnostic(index, _uia.CreateSnapshot(element)))
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

        if (request.FoundIndex is int foundIndex)
        {
            if (foundIndex < 0 || foundIndex >= allMatches.Count)
            {
                return new NativeUiaResolveResult
                {
                    Stage = "index-out-of-range",
                    LastError = $"ComboBox foundIndex {foundIndex} is out of range (count={allMatches.Count}).",
                    Candidates = candidates
                };
            }

            return new NativeUiaResolveResult { Element = allMatches[foundIndex], Candidates = candidates };
        }

        if (allMatches.Count > 1)
        {
            return new NativeUiaResolveResult
            {
                IsAmbiguous = true,
                Stage = "ambiguous-combobox",
                LastError = $"Found {allMatches.Count} ComboBox candidates. Provide foundIndex to disambiguate.",
                Candidates = candidates
            };
        }

        return new NativeUiaResolveResult { Element = allMatches[0], Candidates = candidates };
    }

    private NativeUiaResolveResult SingleResult(IUIAutomationElement element) =>
        new() { Element = element, Candidates = [NativeUiaDiagnostics.CandidateDiagnostic(0, _uia.CreateSnapshot(element))] };

    private List<IUIAutomationElement> FindComboMatches(IUIAutomationElement root, NativeUiaLocator locator)
    {
        var matches = _finder.FindElements(root, locator.AsComboBoxLocator(), limit: MaxSearchLimit);
        if (matches.Count > 0)
            return matches;

        if (string.IsNullOrWhiteSpace(locator.AutomationId) && string.IsNullOrWhiteSpace(locator.Name))
            return matches;

        var promoted = new List<IUIAutomationElement>();
        foreach (var candidate in _finder.FindElements(root, locator.WithoutControlType(), limit: MaxSearchLimit))
        {
            var combo = PromoteToComboBox(candidate);
            if (combo != null && ElementMatchesComboLocator(combo, locator))
                promoted.Add(combo);
        }

        return promoted;
    }

    private bool ElementMatchesComboLocator(IUIAutomationElement element, NativeUiaLocator locator)
    {
        var snapshot = _uia.CreateSnapshot(element);
        if (!string.IsNullOrWhiteSpace(locator.Name)
            && !NativeUiaText.Matches(snapshot.Name, locator.Name, locator.MatchMode))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.AutomationId)
            && !NativeUiaText.Matches(snapshot.AutomationId, locator.AutomationId, locator.MatchMode))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ClassName)
            && !NativeUiaText.Matches(snapshot.ClassName, locator.ClassName, locator.MatchMode))
            return false;

        return _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId) == ComboBoxControlTypeId;
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

    private List<IUIAutomationElement> BuildSearchRoots(IntPtr? activeWindowHwnd, int? processId, bool includeDesktop)
    {
        var roots = new List<IUIAutomationElement>();

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        if (processId.HasValue && activeWindowHwnd is not > 0)
        {
            var processWindows = _uia.FindAllDescendants(
                _uia.Root,
                _uia.And(
                    _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                    _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50032)),
                20);
            roots.AddRange(processWindows);
        }

        if (includeDesktop)
            roots.Add(_uia.Root);

        return roots.DistinctBy(RuntimeKey).ToList();
    }

    private static NativeUiaLocator ToNativeLocator(UiRequest request, int? processId) => new()
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
