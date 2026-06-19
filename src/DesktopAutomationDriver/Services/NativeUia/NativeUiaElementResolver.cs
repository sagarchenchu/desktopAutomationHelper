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
    private const int MaxElementCandidates = 80;
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

        NativeUiaResolveResult? lastComboResult = null;

        if (activeWindowHwnd.HasValue && activeWindowHwnd.Value != IntPtr.Zero)
        {
            var root = _uia.FromHandle(activeWindowHwnd.Value);
            if (root != null)
            {
                var result = FindComboBoxUnderRootBounded(
                    request,
                    root,
                    processId,
                    deadlineUtc,
                    cancellationToken,
                    stage: "active-window");

                if (result.Element != null || result.IsAmbiguous)
                    return result;

                lastComboResult = result;
            }
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
                    cancellationToken,
                    stage: "process-window");

                if (result.Element != null || result.IsAmbiguous)
                    return result;

                lastComboResult = result;
            }
        }

        if (lastComboResult != null)
            return lastComboResult;

        return NativeUiaResolveResult.NotFound(
            "no-root",
            "No active hwnd or process id.");
    }

    public NativeUiaResolveResult ResolveElement(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "ResolveElement start");
        cancellationToken.ThrowIfCancellationRequested();

        var locator = ToNativeLocator(request, processId);
        if (locator.Hwnd is > 0)
        {
            var fromHandle = _uia.FromHandle(new IntPtr(locator.Hwnd.Value));
            if (fromHandle != null && ElementMatchesLocator(fromHandle, locator))
                return SingleResult(fromHandle, "hwnd-direct");
        }

        NativeUiaResolveResult? lastResult = null;

        if (activeWindowHwnd.HasValue && activeWindowHwnd.Value != IntPtr.Zero)
        {
            var root = _uia.FromHandle(activeWindowHwnd.Value);
            if (root != null)
            {
                var result = FindElementUnderRootBounded(
                    request,
                    root,
                    processId,
                    deadlineUtc,
                    cancellationToken,
                    stage: "active-window");

                if (result.Element != null || result.IsAmbiguous)
                    return result;

                lastResult = result;
            }
        }

        if (processId.HasValue)
        {
            var rootWindows = FindTopLevelWindowsForProcessBounded(
                processId.Value,
                deadlineUtc,
                cancellationToken);

            foreach (var window in rootWindows)
            {
                ThrowIfExpired(deadlineUtc, "ResolveElement process window loop");
                cancellationToken.ThrowIfCancellationRequested();

                var result = FindElementUnderRootBounded(
                    request,
                    window,
                    processId,
                    deadlineUtc,
                    cancellationToken,
                    stage: "process-window");

                if (result.Element != null || result.IsAmbiguous)
                    return result;

                lastResult = result;
            }
        }

        if (lastResult != null)
            return lastResult;

        return NativeUiaResolveResult.NotFound(
            "no-root",
            "No active hwnd or process id.");
    }

    private NativeUiaResolveResult FindElementUnderRootBounded(
        UiRequest request,
        IUIAutomationElement root,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        string stage)
    {
        ThrowIfExpired(deadlineUtc, "FindElementUnderRootBounded");
        cancellationToken.ThrowIfCancellationRequested();

        var locator = ToNativeLocator(request, processId);
        var viewCondition = ResolveViewCondition(request);
        var viewName = ResolveViewName(request);
        var matches = FindMatchingElementsBounded(
            root,
            locator,
            viewCondition,
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
                Stage = "element-not-found",
                LastError = $"Native UIA resolver could not find an element for the locator (stage={stage}, view={viewName}).",
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
                    LastError = $"Element foundIndex {foundIndex} is out of range (count={matches.Count}).",
                    Candidates = candidates
                };
            }

            return new NativeUiaResolveResult
            {
                Element = matches[foundIndex],
                Stage = stage,
                Candidates = candidates
            };
        }

        if (matches.Count > 1)
        {
            return new NativeUiaResolveResult
            {
                IsAmbiguous = true,
                Stage = "ambiguous-element",
                LastError = $"Found {matches.Count} element candidates. Provide foundIndex to disambiguate.",
                Candidates = candidates
            };
        }

        return new NativeUiaResolveResult
        {
            Element = matches[0],
            Stage = stage,
            Candidates = candidates
        };
    }

    private NativeUiaResolveResult FindComboBoxUnderRootBounded(
        UiRequest request,
        IUIAutomationElement root,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        string stage)
    {
        ThrowIfExpired(deadlineUtc, "FindComboBoxUnderRootBounded");
        cancellationToken.ThrowIfCancellationRequested();

        var locator = ToNativeLocator(request, processId);
        var viewCondition = ResolveViewCondition(request);
        var viewName = ResolveViewName(request);
        var matches = FindMatchingComboBoxesBounded(
            root,
            locator,
            viewCondition,
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
                LastError = $"Native UIA resolver could not find a ComboBox for the locator (stage={stage}, view={viewName}).",
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

    private NativeUiaResolveResult SingleResult(IUIAutomationElement element, string? stage = null) =>
        new()
        {
            Element = element,
            Stage = stage,
            Candidates = [NativeUiaDiagnostics.CandidateDiagnostic(0, _uia.CreateSnapshot(element))]
        };

    private List<IUIAutomationElement> FindMatchingElementsBounded(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        IUIAutomationCondition viewCondition,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var results = new List<IUIAutomationElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        FindMatchingElementsBoundedRecursive(
            root,
            locator,
            viewCondition,
            results,
            seen,
            depth: 0,
            maxDepth: MaxSearchDepth,
            maxCandidates: MaxElementCandidates,
            deadlineUtc,
            cancellationToken);

        return results;
    }

    private void FindMatchingElementsBoundedRecursive(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        IUIAutomationCondition viewCondition,
        List<IUIAutomationElement> results,
        HashSet<string> seen,
        int depth,
        int maxDepth,
        int maxCandidates,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        ThrowIfExpired(deadlineUtc, "FindMatchingElementsBoundedRecursive");
        cancellationToken.ThrowIfCancellationRequested();

        if (depth > maxDepth)
            return;

        if (results.Count >= maxCandidates)
            return;

        try
        {
            if (ElementMatchesLocator(root, locator))
            {
                var key = RuntimeKey(root);
                if (seen.Add(key))
                {
                    results.Add(root);

                    if (ShouldStopAfterFirstStrongMatch(locator))
                        return;

                    if (results.Count >= maxCandidates)
                        return;
                }
            }
        }
        catch
        {
            // Ignore stale/bad UIA element.
        }

        IUIAutomationElementArray children;

        try
        {
            children = root.FindAll(
                TreeScope.TreeScope_Children,
                viewCondition);
        }
        catch
        {
            return;
        }

        var childCount = Math.Min(children.Length, MaxChildrenPerNode);

        for (var i = 0; i < childCount; i++)
        {
            ThrowIfExpired(deadlineUtc, "FindMatchingElementsBoundedRecursive loop");
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

            FindMatchingElementsBoundedRecursive(
                child,
                locator,
                viewCondition,
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

    private bool ElementMatchesLocator(IUIAutomationElement element, NativeUiaLocator locator)
    {
        if (locator.Hwnd is > 0)
        {
            var hwnd = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
            if (hwnd != locator.Hwnd.Value)
                return false;
        }

        if (locator.ProcessId.HasValue)
        {
            var pid = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ProcessIdPropertyId);
            if (pid != locator.ProcessId.Value)
                return false;
        }

        var controlTypeId = NativeUiaText.ParseControlTypeId(locator.ControlType);
        if (controlTypeId.HasValue)
        {
            var actualType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
            if (actualType != controlTypeId.Value)
                return false;
        }

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

        if (!string.IsNullOrWhiteSpace(locator.Value)
            && !NativeUiaText.Matches(snapshot.Value, locator.Value, locator.MatchMode))
            return false;

        return true;
    }

    private List<IUIAutomationElement> FindMatchingComboBoxesBounded(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        IUIAutomationCondition viewCondition,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var results = new List<IUIAutomationElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        FindMatchingComboBoxesBoundedRecursive(
            root,
            locator,
            viewCondition,
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
        IUIAutomationCondition viewCondition,
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
                viewCondition);
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
                viewCondition,
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

    private IUIAutomationCondition ResolveViewCondition(UiRequest request)
    {
        var requestedView = ResolveViewName(request);

        return requestedView.Trim().ToLowerInvariant() switch
        {
            "raw" => _uia.TrueCondition(),
            "content" => _uia.ContentViewCondition,
            _ => _uia.ControlViewCondition
        };
    }

    private static string ResolveViewName(UiRequest request) =>
        request.View ?? request.TreeView ?? InferDefaultView(request);

    internal static string InferDefaultView(UiRequest request)
    {
        var operation = request.Operation?.Trim().ToLowerInvariant();

        if (operation is "findcomboboxuia" or "selectcomboboxuia" or "inspectcombobox")
            return "control";

        if (operation == "clickmenuuia")
            return "raw";

        var controlType = request.Locator?.ControlType;

        if (string.Equals(controlType, "Menu", StringComparison.OrdinalIgnoreCase)
            || string.Equals(controlType, "MenuItem", StringComparison.OrdinalIgnoreCase)
            || string.Equals(controlType, "ControlType(50011)", StringComparison.OrdinalIgnoreCase)
            || string.Equals(controlType, "50011", StringComparison.OrdinalIgnoreCase))
        {
            return "raw";
        }

        return "control";
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
