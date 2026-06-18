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
    private const int WindowControlTypeId = 50032;
    private const int MaxComboBoxCandidates = 80;
    private const int MaxSearchDepth = 8;
    private const int MaxChildrenPerNode = 200;
    private const int MaxProcessWindows = 20;

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
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        bool allowDesktopFallback = false)
    {
        ThrowIfExpired(deadlineUtc, "ResolveComboBox");
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

        if (!HasSearchContext(request, activeWindowHwnd, processId))
        {
            return new NativeUiaResolveResult
            {
                Stage = "no-active-window",
                LastError = "No active window hwnd or processId. Launch/attach and switchwindow first."
            };
        }

        var searchRoots = new List<IUIAutomationElement>();
        if (request.ParentLocator != null)
        {
            var parentLocator = ToNativeLocator(request.ParentLocator, request, processId);
            foreach (var root in BuildSearchRoots(activeWindowHwnd, processId, allowDesktopFallback, deadlineUtc, cancellationToken))
            {
                ThrowIfExpired(deadlineUtc, "ResolveComboBox parent-root");
                cancellationToken.ThrowIfCancellationRequested();

                var parent = _finder.FindFirst(root, parentLocator, descendants: true, limit: 50);
                if (parent != null)
                {
                    searchRoots.Add(parent);
                    break;
                }
            }

            if (searchRoots.Count == 0)
            {
                return new NativeUiaResolveResult
                {
                    Stage = "parent-not-found",
                    LastError = "parentLocator did not resolve to a container element."
                };
            }
        }
        else
        {
            searchRoots.AddRange(BuildSearchRoots(activeWindowHwnd, processId, allowDesktopFallback, deadlineUtc, cancellationToken));
        }

        if (searchRoots.Count == 0)
        {
            return new NativeUiaResolveResult
            {
                Stage = "no-active-window",
                LastError = "No active window hwnd or processId. Launch/attach and switchwindow first."
            };
        }

        var allMatches = new List<IUIAutomationElement>();
        foreach (var root in searchRoots)
        {
            ThrowIfExpired(deadlineUtc, "ResolveComboBox search");
            cancellationToken.ThrowIfCancellationRequested();
            allMatches.AddRange(FindComboMatches(root, locator, deadlineUtc, cancellationToken));
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

    private List<IUIAutomationElement> FindComboMatches(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(locator.AutomationId) || !string.IsNullOrWhiteSpace(locator.Name))
        {
            ThrowIfExpired(deadlineUtc, "FindComboMatches targeted");
            cancellationToken.ThrowIfCancellationRequested();

            var targeted = _finder.FindFirst(root, locator.AsComboBoxLocator(), descendants: true, limit: 1);
            if (targeted != null && ElementMatchesComboLocator(targeted, locator))
                return [targeted];

            foreach (var candidate in _finder.FindElements(root, locator.WithoutControlType(), descendants: true, limit: 20))
            {
                ThrowIfExpired(deadlineUtc, "FindComboMatches promote");
                cancellationToken.ThrowIfCancellationRequested();

                var combo = PromoteToComboBox(candidate);
                if (combo != null && ElementMatchesComboLocator(combo, locator))
                    return [combo];
            }
        }

        var bounded = new List<IUIAutomationElement>();
        FindComboBoxesBounded(
            root,
            bounded,
            depth: 0,
            maxDepth: MaxSearchDepth,
            maxCandidates: MaxComboBoxCandidates,
            deadlineUtc,
            cancellationToken);

        return bounded
            .Where(e => ElementMatchesComboLocator(e, locator))
            .ToList();
    }

    private void FindComboBoxesBounded(
        IUIAutomationElement root,
        List<IUIAutomationElement> results,
        int depth,
        int maxDepth,
        int maxCandidates,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "FindComboBoxesBounded");
        cancellationToken.ThrowIfCancellationRequested();

        if (depth > maxDepth || results.Count >= maxCandidates)
            return;

        IUIAutomationElement[] children;
        try
        {
            children = _uia.GetChildren(root, MaxChildrenPerNode);
        }
        catch
        {
            return;
        }

        foreach (var child in children)
        {
            ThrowIfExpired(deadlineUtc, "FindComboBoxesBounded loop");
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var controlType = _uia.GetIntProperty(child, UIA_PropertyIds.UIA_ControlTypePropertyId);
                if (controlType == ComboBoxControlTypeId)
                {
                    results.Add(child);
                    if (results.Count >= maxCandidates)
                        return;
                }
            }
            catch
            {
                // Ignore stale/unavailable element.
            }

            FindComboBoxesBounded(
                child,
                results,
                depth + 1,
                maxDepth,
                maxCandidates,
                deadlineUtc,
                cancellationToken);

            if (results.Count >= maxCandidates)
                return;
        }
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

    private List<IUIAutomationElement> BuildSearchRoots(
        IntPtr? activeWindowHwnd,
        int? processId,
        bool allowDesktopFallback,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var roots = new List<IUIAutomationElement>();

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }
        else if (processId.HasValue)
        {
            ThrowIfExpired(deadlineUtc, "BuildSearchRoots process-windows");
            cancellationToken.ThrowIfCancellationRequested();

            var processWindows = _uia.FindAllChildren(
                _uia.Root,
                _uia.And(
                    _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                    _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, WindowControlTypeId)),
                MaxProcessWindows);
            roots.AddRange(processWindows);
        }

        if (allowDesktopFallback && roots.Count == 0)
            roots.Add(_uia.Root);

        return roots.DistinctBy(RuntimeKey).ToList();
    }

    private static bool HasSearchContext(UiRequest request, IntPtr? activeWindowHwnd, int? processId)
    {
        if (request.Hwnd is > 0 || request.Locator?.Hwnd is > 0 || request.Locator?.Handle is > 0)
            return true;

        if (activeWindowHwnd is > 0)
            return true;

        return processId.HasValue;
    }

    private static void ThrowIfExpired(DateTime deadlineUtc, string stage)
    {
        if (DateTime.UtcNow > deadlineUtc)
            throw new TimeoutException($"{stage} exceeded timeout.");
    }

    private static NativeUiaLocator ToNativeLocator(UiRequest request, int? processId) =>
        ToNativeLocator(request.Locator, request, processId);

    private static NativeUiaLocator ToNativeLocator(UiLocator? locator, UiRequest request, int? processId) => new()
    {
        Name = locator?.Name,
        AutomationId = locator?.AutomationId,
        ClassName = locator?.ClassName ?? request.ClassName,
        ControlType = locator?.ControlType,
        Hwnd = request.Hwnd ?? locator?.Hwnd ?? locator?.Handle,
        ProcessId = request.ProcessId ?? locator?.ProcessId ?? processId,
        FoundIndex = request.FoundIndex ?? locator?.FoundIndex ?? locator?.Index,
        MatchMode = string.IsNullOrWhiteSpace(request.MatchMode)
            ? locator?.MatchMode ?? "exact"
            : request.MatchMode!
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
