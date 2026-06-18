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

    public static NativeUiaResolveResult NotFound(string stage, string? message = null) =>
        new() { Stage = stage, LastError = message };
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

    public NativeUiaElementResolver(NativeUiaAutomation uia)
    {
        _uia = uia;
    }

    public NativeUiaResolveResult ResolveComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "ResolveComboBox start");
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

        if (activeWindowHwnd.HasValue && activeWindowHwnd.Value != IntPtr.Zero)
        {
            var root = _uia.FromHandle(activeWindowHwnd.Value);
            if (root == null)
            {
                return NativeUiaResolveResult.NotFound(
                    "not-found",
                    "Could not resolve activeWindowHwnd to UIA root.");
            }

            return FindComboBoxUnderRootBounded(
                request,
                root,
                processId,
                deadlineUtc,
                cancellationToken);
        }

        if (processId.HasValue)
        {
            var rootWindows = FindTopLevelWindowsForProcessBounded(
                processId.Value,
                deadlineUtc,
                cancellationToken);

            foreach (var window in rootWindows)
            {
                ThrowIfExpired(deadlineUtc, "ResolveComboBox process window loop");
                cancellationToken.ThrowIfCancellationRequested();

                var result = FindComboBoxUnderRootBounded(
                    request,
                    window,
                    processId,
                    deadlineUtc,
                    cancellationToken);

                if (result.Element != null)
                    return result;
            }

            return NativeUiaResolveResult.NotFound(
                "not-found",
                "No matching ComboBox found under process windows.");
        }

        return NativeUiaResolveResult.NotFound(
            "no-root",
            "No active hwnd or process id.");
    }

    private NativeUiaResolveResult FindComboBoxUnderRootBounded(
        UiRequest request,
        IUIAutomationElement root,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "FindComboBoxUnderRootBounded");
        cancellationToken.ThrowIfCancellationRequested();

        var locator = ToNativeLocator(request, processId);
        var matches = FindMatchingComboBoxesBounded(
            root,
            locator,
            deadlineUtc,
            cancellationToken);

        var candidates = matches
            .Select((element, index) => NativeUiaDiagnostics.CandidateDiagnostic(index, _uia.CreateSnapshot(element)))
            .Cast<object>()
            .ToList();

        if (matches.Count == 0)
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
            if (foundIndex < 0 || foundIndex >= matches.Count)
            {
                return new NativeUiaResolveResult
                {
                    Stage = "index-out-of-range",
                    LastError = $"ComboBox foundIndex {foundIndex} is out of range (count={matches.Count}).",
                    Candidates = candidates
                };
            }

            return new NativeUiaResolveResult { Element = matches[foundIndex], Candidates = candidates };
        }

        if (matches.Count > 1)
        {
            return new NativeUiaResolveResult
            {
                IsAmbiguous = true,
                Stage = "ambiguous-combobox",
                LastError = $"Found {matches.Count} ComboBox candidates. Provide foundIndex to disambiguate.",
                Candidates = candidates
            };
        }

        return new NativeUiaResolveResult { Element = matches[0], Candidates = candidates };
    }

    private NativeUiaResolveResult SingleResult(IUIAutomationElement element) =>
        new() { Element = element, Candidates = [NativeUiaDiagnostics.CandidateDiagnostic(0, _uia.CreateSnapshot(element))] };

    private List<IUIAutomationElement> FindMatchingComboBoxesBounded(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var results = new List<IUIAutomationElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        FindMatchingComboBoxesBoundedRecursive(
            root,
            locator,
            results,
            seen,
            depth: 0,
            maxDepth: MaxSearchDepth,
            maxCandidates: MaxComboBoxCandidates,
            deadlineUtc,
            cancellationToken);

        return results;
    }

    private void FindMatchingComboBoxesBoundedRecursive(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        List<IUIAutomationElement> results,
        HashSet<string> seen,
        int depth,
        int maxDepth,
        int maxCandidates,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "FindMatchingComboBoxesBoundedRecursive");
        cancellationToken.ThrowIfCancellationRequested();

        if (depth > maxDepth)
            return;

        if (results.Count >= maxCandidates)
            return;

        IUIAutomationElementArray children;

        try
        {
            children = root.FindAll(
                TreeScope.TreeScope_Children,
                _uia.ControlViewCondition);
        }
        catch
        {
            return;
        }

        var childCount = Math.Min(children.Length, MaxChildrenPerNode);

        for (var i = 0; i < childCount; i++)
        {
            ThrowIfExpired(deadlineUtc, "FindMatchingComboBoxesBoundedRecursive loop");
            cancellationToken.ThrowIfCancellationRequested();

            IUIAutomationElement child;

            try
            {
                child = children.GetElement(i);
            }
            catch
            {
                continue;
            }

            try
            {
                var controlType = _uia.GetIntProperty(
                    child,
                    UIA_PropertyIds.UIA_ControlTypePropertyId);

                if (controlType == ComboBoxControlTypeId)
                {
                    if (ElementMatchesComboLocator(child, locator))
                    {
                        var key = RuntimeKey(child);

                        if (seen.Add(key))
                        {
                            results.Add(child);

                            if (ShouldStopAfterFirstStrongMatch(locator))
                                return;

                            if (results.Count >= maxCandidates)
                                return;
                        }
                    }
                }
            }
            catch
            {
                // Ignore stale/bad UIA element.
            }

            FindMatchingComboBoxesBoundedRecursive(
                child,
                locator,
                results,
                seen,
                depth + 1,
                maxDepth,
                maxCandidates,
                deadlineUtc,
                cancellationToken);

            if (ShouldStopAfterFirstStrongMatch(locator) && results.Count > 0)
                return;

            if (results.Count >= maxCandidates)
                return;
        }
    }

    private static bool ShouldStopAfterFirstStrongMatch(NativeUiaLocator locator)
    {
        if (locator.FoundIndex.HasValue)
            return false;

        if (locator.Hwnd is > 0)
            return true;

        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            return true;

        return false;
    }

    private List<IUIAutomationElement> FindTopLevelWindowsForProcessBounded(
        int processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "FindTopLevelWindowsForProcessBounded");
        cancellationToken.ThrowIfCancellationRequested();

        var windows = new List<IUIAutomationElement>();

        IUIAutomationElementArray children;
        try
        {
            children = _uia.Root.FindAll(TreeScope.TreeScope_Children, _uia.ControlViewCondition);
        }
        catch
        {
            return windows;
        }

        var childCount = Math.Min(children.Length, MaxChildrenPerNode);
        for (var i = 0; i < childCount; i++)
        {
            ThrowIfExpired(deadlineUtc, "FindTopLevelWindowsForProcessBounded loop");
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var child = children.GetElement(i);
                if (_uia.GetIntProperty(child, UIA_PropertyIds.UIA_ControlTypePropertyId) == WindowControlTypeId
                    && _uia.GetIntProperty(child, UIA_PropertyIds.UIA_ProcessIdPropertyId) == processId)
                {
                    windows.Add(child);
                    if (windows.Count >= MaxProcessWindows)
                        return windows;
                }
            }
            catch
            {
                // ignore stale element
            }
        }

        return windows;
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
