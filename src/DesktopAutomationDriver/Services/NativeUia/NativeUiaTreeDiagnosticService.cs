using System.Diagnostics;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Read-only bounded Native UIA tree diagnostics for dumpuia/finduia.
/// </summary>
internal sealed class NativeUiaTreeDiagnosticService : INativeUiaTreeDiagnosticService
{
    private const int WindowControlTypeId = 50032;
    private const int DefaultTimeoutMs = 5000;
    private const int MaxTimeoutMs = 15000;
    private const int DefaultMaxDepth = 8;
    private const int DefaultMaxChildrenPerNode = 200;
    private const int MaxTotalNodes = 2000;
    private const int MaxProcessWindows = 20;
    private const int MaxDesktopChildren = 200;
    private const int DeadlineSafetyMs = 50;
    private const int MaxDiagnosticStringLength = 256;

    private static readonly int[] MenuLikeControlTypeIds =
    [
        50018, // MenuBar
        50021, // ToolBar
        50010, // MenuItem
        50011, // Menu
        50000, // Button
        50020, // Text
        50033, // Pane
        50026  // Custom
    ];

    private static readonly int[] CheapPatternIds =
    [
        UIA_PatternIds.UIA_InvokePatternId,
        UIA_PatternIds.UIA_ExpandCollapsePatternId,
        UIA_PatternIds.UIA_SelectionItemPatternId,
        UIA_PatternIds.UIA_TogglePatternId,
        UIA_PatternIds.UIA_ValuePatternId,
        UIA_PatternIds.UIA_SelectionPatternId,
        UIA_PatternIds.UIA_ScrollPatternId,
        UIA_PatternIds.UIA_GridPatternId,
        UIA_PatternIds.UIA_WindowPatternId
    ];

    private readonly NativeUiaAutomation _uia;
    private readonly ILogger<NativeUiaTreeDiagnosticService> _logger;

    public NativeUiaTreeDiagnosticService(ILogger<NativeUiaTreeDiagnosticService> logger)
        : this(new NativeUiaAutomation(), logger)
    {
    }

    internal NativeUiaTreeDiagnosticService(NativeUiaAutomation uia, ILogger<NativeUiaTreeDiagnosticService> logger)
    {
        _uia = uia;
        _logger = logger;
    }

    public object DumpTree(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteDiagnostic(
            "dumpuia",
            request,
            activeWindowHwnd,
            processId,
            cancellationToken,
            (options, deadlineUtc, nodes, menuLikeCandidates, timedOut, partialReason) =>
            {
                var filter = BuildDumpFilter(options);
                var collected = new List<object>();
                var totalVisited = 0;
                // IsTruncated is informational (maxDepth/maxChildren/maxNodes).
                // StopTraversal is only maxNodes, timeout, or cancellation.
                var truncated = new StrongBox<bool>(false);
                var stopTraversal = new StrongBox<bool>(false);
                var truncationWarnings = new List<string>();
                var roots = ResolveRoots(options, activeWindowHwnd, processId, deadlineUtc, cancellationToken);

                if (roots.Count == 0)
                {
                    return new Dictionary<string, object?>
                    {
                        ["operation"] = "dumpuia",
                        ["success"] = false,
                        ["reason"] = "no-root",
                        ["message"] = "Could not resolve any UIA root for the requested root mode.",
                        ["view"] = options.View,
                        ["root"] = options.RootMode
                    };
                }

                try
                {
                    var rootOrdinal = 0;
                    foreach (var (root, rootLabel) in roots)
                    {
                        CollectTreeNodes(
                            root,
                            rootLabel,
                            options,
                            filter,
                            deadlineUtc,
                            cancellationToken,
                            collected,
                            truncated,
                            stopTraversal,
                            truncationWarnings,
                            parentNodeIndex: null,
                            ordinal: rootOrdinal,
                            ref totalVisited,
                            depth: 0,
                            pathSegments: options.IncludePath ? [rootLabel] : null);

                        rootOrdinal++;

                        if (stopTraversal.Value)
                            break;

                        if (IsDeadlineNear(deadlineUtc))
                        {
                            timedOut.Value = true;
                            partialReason.Value = "timeout";
                            truncated.Value = true;
                            stopTraversal.Value = true;
                            truncationWarnings.Add("Traversal stopped early due to timeout.");
                            break;
                        }
                    }
                }
                catch (TimeoutException ex)
                {
                    timedOut.Value = true;
                    partialReason.Value = "timeout";
                    truncated.Value = true;
                    stopTraversal.Value = true;
                    truncationWarnings.Add(ex.Message);
                }

                nodes.AddRange(collected);

                return new Dictionary<string, object?>
                {
                    ["operation"] = "dumpuia",
                    ["success"] = !timedOut.Value,
                    ["reason"] = timedOut.Value ? partialReason.Value ?? "timeout" : null,
                    ["semanticDiagnosticVersion"] = NativeUiaSemanticContract.SemanticDiagnosticVersion,
                    ["view"] = options.View,
                    ["root"] = options.RootMode,
                    ["maxDepth"] = options.MaxDepth,
                    ["maxChildren"] = options.MaxChildrenPerNode,
                    ["maxNodes"] = options.MaxNodes,
                    ["includeOffscreen"] = options.IncludeOffscreen,
                    ["includeBounds"] = options.IncludeBounds,
                    ["includeHelpText"] = options.IncludeHelpText,
                    ["includeFocusState"] = options.IncludeFocusState,
                    ["includeRuntimeId"] = options.IncludeRuntimeId,
                    ["includeTreeLinks"] = options.IncludeTreeLinks,
                    ["includeWindowState"] = options.IncludeWindowState,
                    ["filter"] = filter,
                    ["totalVisited"] = totalVisited,
                    ["matchCount"] = collected.Count,
                    ["truncated"] = truncated.Value,
                    ["truncationWarnings"] = truncationWarnings.Count > 0 ? truncationWarnings : null,
                    ["nodes"] = collected,
                    ["partialResults"] = timedOut.Value || truncated.Value ? collected : null,
                    ["menuLikeCandidates"] = menuLikeCandidates.Count > 0 ? menuLikeCandidates : null
                };
            });

    public object FindElement(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteDiagnostic(
            "finduia",
            request,
            activeWindowHwnd,
            processId,
            cancellationToken,
            (options, deadlineUtc, nodes, menuLikeCandidates, timedOut, partialReason) =>
            {
                var locator = ToNativeLocator(request, processId);
                if (!HasFindCriteria(request, options, locator))
                {
                    return new Dictionary<string, object?>
                    {
                        ["operation"] = "finduia",
                        ["success"] = false,
                        ["found"] = false,
                        ["reason"] = "invalid-request",
                        ["message"] = "Provide locator fields and/or nameContains for finduia."
                    };
                }

                var matches = new List<object>();
                var totalVisited = 0;
                var roots = ResolveRoots(options, activeWindowHwnd, processId, deadlineUtc, cancellationToken);

                if (roots.Count == 0)
                {
                    return new Dictionary<string, object?>
                    {
                        ["operation"] = "finduia",
                        ["success"] = false,
                        ["found"] = false,
                        ["reason"] = "no-root",
                        ["message"] = "Could not resolve any UIA root for the requested root mode.",
                        ["view"] = options.View,
                        ["root"] = options.RootMode
                    };
                }

                try
                {
                    foreach (var (root, rootLabel) in roots)
                    {
                        FindMatchingNodes(
                            root,
                            rootLabel,
                            locator,
                            options,
                            deadlineUtc,
                            cancellationToken,
                            matches,
                            menuLikeCandidates,
                            ref totalVisited,
                            depth: 0,
                            pathSegments: options.IncludePath ? [rootLabel] : null);

                        if (IsDeadlineNear(deadlineUtc))
                        {
                            timedOut.Value = true;
                            partialReason.Value = "timeout";
                            break;
                        }
                    }
                }
                catch (TimeoutException)
                {
                    timedOut.Value = true;
                    partialReason.Value = "timeout";
                }

                if (locator.FoundIndex is int foundIndex)
                {
                    if (foundIndex < 0 || foundIndex >= matches.Count)
                        matches.Clear();
                    else
                        matches = [matches[foundIndex]];
                }

                nodes.AddRange(matches);

                if (matches.Count == 0 && !timedOut.Value)
                {
                    CollectMenuLikeCandidates(
                        options,
                        activeWindowHwnd,
                        processId,
                        deadlineUtc,
                        cancellationToken,
                        menuLikeCandidates,
                        maxDepth: Math.Min(options.MaxDepth, 4));
                }

                var found = matches.Count > 0;

                return new Dictionary<string, object?>
                {
                    ["operation"] = "finduia",
                    ["success"] = found && !timedOut.Value,
                    ["found"] = found,
                    ["reason"] = timedOut.Value
                        ? partialReason.Value ?? "timeout"
                        : found ? null : "element-not-found",
                    ["view"] = options.View,
                    ["root"] = options.RootMode,
                    ["maxDepth"] = options.MaxDepth,
                    ["includeOffscreen"] = options.IncludeOffscreen,
                    ["locator"] = DescribeLocator(locator),
                    ["nameContains"] = options.NameContains,
                    ["matchCount"] = matches.Count,
                    ["matches"] = matches,
                    ["partialResults"] = timedOut.Value ? matches : null,
                    ["menuLikeCandidates"] = !found || menuLikeCandidates.Count > 0 ? menuLikeCandidates : null,
                    ["totalVisited"] = totalVisited
                };
            });

    private object ExecuteDiagnostic(
        string operation,
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken,
        Func<
            NativeUiaTreeDiagnosticOptions,
            DateTime,
            List<object>,
            List<object>,
            StrongBox<bool>,
            StrongBox<string?>,
            Dictionary<string, object?>> buildResult)
    {
        var sw = Stopwatch.StartNew();
        var options = ParseOptions(request);
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        var timedOut = new StrongBox<bool>(false);
        var partialReason = new StrongBox<string?>(null);
        var nodes = new List<object>();
        var menuLikeCandidates = new List<object>();

        _logger.LogInformation(
            "NativeUiaTree {Operation} starting. rootHwnd={RootHwnd}, processId={ProcessId}, view={View}, root={Root}, timeoutMs={TimeoutMs}",
            operation,
            activeWindowHwnd,
            processId,
            options.View,
            options.RootMode,
            timeoutMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            var rootDiagnostics = BuildRootDiagnostics(
                options,
                activeWindowHwnd,
                processId,
                deadlineUtc,
                cancellationToken);

            var payload = buildResult(
                options,
                deadlineUtc,
                nodes,
                menuLikeCandidates,
                timedOut,
                partialReason);

            payload["rootDiagnostics"] = rootDiagnostics;
            payload["elapsedMs"] = sw.ElapsedMilliseconds;
            payload["timeoutMs"] = timeoutMs;

            return payload;
        }
        catch (TimeoutException ex)
        {
            _logger.LogWarning(ex, "NativeUiaTree {Operation} timed out. elapsedMs={ElapsedMs}", operation, sw.ElapsedMilliseconds);

            return new
            {
                operation,
                success = false,
                reason = "timeout",
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds,
                partialResults = nodes.Count > 0 ? nodes : null,
                message = ex.Message
            };
        }
        catch (OperationCanceledException)
        {
            return new
            {
                operation,
                success = false,
                reason = "cancelled",
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds,
                partialResults = nodes.Count > 0 ? nodes : null
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaTree {Operation} failed.", operation);

            return new
            {
                operation,
                success = false,
                reason = "error",
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds,
                partialResults = nodes.Count > 0 ? nodes : null,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private static DumpFilter BuildDumpFilter(NativeUiaTreeDiagnosticOptions options) => new()
    {
        NameContains = options.NameContains,
        ControlType = options.ControlType,
        ClassName = options.ClassName,
        AutomationId = options.AutomationId
    };

    private object BuildRootDiagnostics(
        NativeUiaTreeDiagnosticOptions options,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "root-diagnostics");

        IUIAutomationElement? activeWindowElement = null;
        object? activeWindowSnapshot = null;

        if (activeWindowHwnd is > 0)
        {
            activeWindowElement = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeWindowElement != null)
            {
                activeWindowSnapshot = BuildElementRecord(
                    activeWindowElement,
                    depth: 0,
                    path: "activeWindow",
                    options,
                    nodeIndex: null,
                    parentNodeIndex: null,
                    ordinal: null);
            }
        }

        var processWindows = new List<object>();
        if (processId.HasValue)
        {
            foreach (var window in FindTopLevelWindowsForProcess(processId.Value, deadlineUtc, cancellationToken))
            {
                EnsureWithinDeadline(deadlineUtc, cancellationToken, "process-window-snapshot");
                processWindows.Add(BuildElementRecord(
                    window,
                    depth: 0,
                    path: "processWindow",
                    options,
                    nodeIndex: null,
                    parentNodeIndex: null,
                    ordinal: null));
            }
        }

        // Prefer the live active-window process when the session PID is stale/exited.
        var identityProcessId = processId;
        if (activeWindowSnapshot is Dictionary<string, object?> activeRecord
            && activeRecord.TryGetValue("processId", out var activePidObj)
            && activePidObj is int activePid
            && activePid > 0)
        {
            identityProcessId = activePid;
        }

        string? processName = null;
        int? windowsSessionId = null;
        if (identityProcessId is int pid)
        {
            try
            {
                using var process = Process.GetProcessById(pid);
                processName = BoundString(process.ProcessName);
                windowsSessionId = process.SessionId;
            }
            catch
            {
                // Process may have exited; leave identity fields null.
            }
        }

        return new
        {
            activeWindowHwnd = activeWindowHwnd is > 0 ? activeWindowHwnd.Value.ToInt64() : (long?)null,
            activeWindow = activeWindowSnapshot,
            processId = identityProcessId,
            processName,
            windowsSessionId,
            processWindowCount = processWindows.Count,
            processWindows
        };
    }

    private List<(IUIAutomationElement Element, string Label)> ResolveRoots(
        NativeUiaTreeDiagnosticOptions options,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "resolve-roots");

        return options.RootMode.ToLowerInvariant() switch
        {
            "processwindows" or "process-windows" or "process_windows" => ResolveProcessWindowRoots(processId, deadlineUtc, cancellationToken),
            "desktopchildren" or "desktop-children" or "desktop_children" => ResolveDesktopChildRoots(processId, deadlineUtc, cancellationToken),
            _ => ResolveActiveWindowRoots(activeWindowHwnd, processId, deadlineUtc, cancellationToken)
        };
    }

    private List<(IUIAutomationElement Element, string Label)> ResolveActiveWindowRoots(
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        if (activeWindowHwnd is > 0)
        {
            var root = _uia.FromHandle(activeWindowHwnd.Value);
            if (root != null)
                return [(root, "activeWindow")];
        }

        if (processId.HasValue)
            return ResolveProcessWindowRoots(processId, deadlineUtc, cancellationToken);

        return [];
    }

    private List<(IUIAutomationElement Element, string Label)> ResolveProcessWindowRoots(
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        if (!processId.HasValue)
            return [];

        var roots = new List<(IUIAutomationElement, string)>();
        var index = 0;

        foreach (var window in FindTopLevelWindowsForProcess(processId.Value, deadlineUtc, cancellationToken))
        {
            roots.Add((window, $"processWindow[{index}]"));
            index++;
        }

        return roots;
    }

    private List<(IUIAutomationElement Element, string Label)> ResolveDesktopChildRoots(
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "desktop-children");

        var roots = new List<(IUIAutomationElement, string)>();
        IUIAutomationElementArray children;

        try
        {
            children = _uia.Root.FindAll(TreeScope.TreeScope_Children, _uia.TrueCondition());
        }
        catch
        {
            return roots;
        }

        var childCount = Math.Min(children.Length, MaxDesktopChildren);
        var index = 0;

        for (var i = 0; i < childCount; i++)
        {
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "desktop-children loop");
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var child = children.GetElement(i);
                if (processId.HasValue)
                {
                    var pid = _uia.GetIntProperty(child, UIA_PropertyIds.UIA_ProcessIdPropertyId);
                    if (pid != processId.Value)
                        continue;
                }

                roots.Add((child, $"desktopChild[{index}]"));
                index++;

                if (roots.Count >= MaxProcessWindows)
                    break;
            }
            catch
            {
                // ignore stale element
            }
        }

        return roots;
    }

    private void CollectTreeNodes(
        IUIAutomationElement element,
        string segmentLabel,
        NativeUiaTreeDiagnosticOptions options,
        DumpFilter filter,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        List<object> collected,
        StrongBox<bool> truncated,
        StrongBox<bool> stopTraversal,
        List<string> truncationWarnings,
        int? parentNodeIndex,
        int ordinal,
        ref int totalVisited,
        int depth,
        List<string>? pathSegments)
    {
        if (stopTraversal.Value)
            return;

        if (IsDeadlineNear(deadlineUtc))
            throw new TimeoutException("dumpuia traversal exceeded timeout.");

        cancellationToken.ThrowIfCancellationRequested();
        totalVisited++;

        int? currentNodeIndex = null;
        if (ShouldIncludeElement(element, options))
        {
            var provisionalIndex = collected.Count;
            var record = BuildElementRecord(
                element,
                depth,
                BuildPath(pathSegments, segmentLabel, options.IncludePath),
                options,
                nodeIndex: provisionalIndex,
                parentNodeIndex: parentNodeIndex,
                ordinal: ordinal);

            if (filter.Matches(record))
            {
                if (collected.Count >= options.MaxNodes || collected.Count >= MaxTotalNodes)
                {
                    truncated.Value = true;
                    stopTraversal.Value = true;
                    truncationWarnings.Add(
                        collected.Count >= MaxTotalNodes
                            ? "dumpuia reached absolute max node cap."
                            : "dumpuia reached requested maxNodes cap.");
                    return;
                }

                collected.Add(record);
                currentNodeIndex = provisionalIndex;
            }
        }

        // Informational truncation (maxChildren/maxDepth) must not stop sibling traversal.
        if (stopTraversal.Value)
            return;

        if (depth >= options.MaxDepth)
        {
            if (HasAnyChild(element, options, deadlineUtc, cancellationToken))
            {
                truncated.Value = true;
                var warning = NativeUiaSemanticContract.MaxDepthTruncationWarning(options.MaxDepth);
                if (!truncationWarnings.Contains(warning, StringComparer.Ordinal))
                    truncationWarnings.Add(warning);
            }

            return;
        }

        var childOrdinal = 0;
        foreach (var (child, childLabel, childPath) in GetChildren(
                     element,
                     options,
                     pathSegments,
                     segmentLabel,
                     depth + 1,
                     deadlineUtc,
                     cancellationToken,
                     truncated,
                     truncationWarnings))
        {
            CollectTreeNodes(
                child,
                childLabel,
                options,
                filter,
                deadlineUtc,
                cancellationToken,
                collected,
                truncated,
                stopTraversal,
                truncationWarnings,
                parentNodeIndex: currentNodeIndex,
                ordinal: childOrdinal,
                ref totalVisited,
                depth + 1,
                childPath);

            childOrdinal++;

            if (stopTraversal.Value)
                return;

            if (IsDeadlineNear(deadlineUtc))
                throw new TimeoutException("dumpuia traversal exceeded timeout.");
        }
    }

    private void FindMatchingNodes(
        IUIAutomationElement element,
        string segmentLabel,
        NativeUiaLocator locator,
        NativeUiaTreeDiagnosticOptions options,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        List<object> matches,
        List<object> menuLikeCandidates,
        ref int totalVisited,
        int depth,
        List<string>? pathSegments)
    {
        if (IsDeadlineNear(deadlineUtc))
            throw new TimeoutException("finduia traversal exceeded timeout.");

        cancellationToken.ThrowIfCancellationRequested();
        totalVisited++;

        if (ShouldIncludeElement(element, options))
        {
            if (ElementMatchesLocator(element, locator, options))
            {
                matches.Add(BuildElementRecord(
                    element,
                    depth,
                    BuildPath(pathSegments, segmentLabel, options.IncludePath),
                    options,
                    nodeIndex: matches.Count,
                    parentNodeIndex: null,
                    ordinal: null));
            }

            if (depth <= 3 && IsMenuLikeControlType(_uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId)))
            {
                menuLikeCandidates.Add(BuildElementRecord(
                    element,
                    depth,
                    BuildPath(pathSegments, segmentLabel, options.IncludePath),
                    options,
                    nodeIndex: null,
                    parentNodeIndex: null,
                    ordinal: null));
            }
        }

        if (depth >= options.MaxDepth)
            return;

        var findTruncated = new StrongBox<bool>(false);
        var findWarnings = new List<string>();
        foreach (var (child, childLabel, childPath) in GetChildren(
                     element,
                     options,
                     pathSegments,
                     segmentLabel,
                     depth + 1,
                     deadlineUtc,
                     cancellationToken,
                     findTruncated,
                     findWarnings))
        {
            FindMatchingNodes(
                child,
                childLabel,
                locator,
                options,
                deadlineUtc,
                cancellationToken,
                matches,
                menuLikeCandidates,
                ref totalVisited,
                depth + 1,
                childPath);

            if (matches.Count >= MaxTotalNodes)
                throw new TimeoutException("finduia reached max node cap.");
        }
    }

    private void CollectMenuLikeCandidates(
        NativeUiaTreeDiagnosticOptions options,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        List<object> menuLikeCandidates,
        int maxDepth)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var limitedOptions = options with { MaxDepth = maxDepth };

        foreach (var (root, rootLabel) in ResolveRoots(limitedOptions, activeWindowHwnd, processId, deadlineUtc, cancellationToken))
        {
            CollectMenuLikeRecursive(
                root,
                rootLabel,
                limitedOptions,
                deadlineUtc,
                cancellationToken,
                menuLikeCandidates,
                seen,
                depth: 0,
                pathSegments: limitedOptions.IncludePath ? [rootLabel] : null);

            if (menuLikeCandidates.Count >= 100 || IsDeadlineNear(deadlineUtc))
                break;
        }
    }

    private void CollectMenuLikeRecursive(
        IUIAutomationElement element,
        string segmentLabel,
        NativeUiaTreeDiagnosticOptions options,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        List<object> menuLikeCandidates,
        HashSet<string> seen,
        int depth,
        List<string>? pathSegments)
    {
        if (IsDeadlineNear(deadlineUtc))
            return;

        cancellationToken.ThrowIfCancellationRequested();

        if (ShouldIncludeElement(element, options)
            && IsMenuLikeControlType(_uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId)))
        {
            var key = RuntimeKey(element);
            if (seen.Add(key))
            {
                menuLikeCandidates.Add(BuildElementRecord(
                    element,
                    depth,
                    BuildPath(pathSegments, segmentLabel, options.IncludePath),
                    options,
                    nodeIndex: null,
                    parentNodeIndex: null,
                    ordinal: null));
            }
        }

        if (depth >= options.MaxDepth || menuLikeCandidates.Count >= 100)
            return;

        var menuTruncated = new StrongBox<bool>(false);
        var menuWarnings = new List<string>();
        foreach (var (child, childLabel, childPath) in GetChildren(
                     element,
                     options,
                     pathSegments,
                     segmentLabel,
                     depth + 1,
                     deadlineUtc,
                     cancellationToken,
                     menuTruncated,
                     menuWarnings))
        {
            CollectMenuLikeRecursive(
                child,
                childLabel,
                options,
                deadlineUtc,
                cancellationToken,
                menuLikeCandidates,
                seen,
                depth + 1,
                childPath);
        }
    }

    private bool HasAnyChild(
        IUIAutomationElement element,
        NativeUiaTreeDiagnosticOptions options,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "has-any-child");
        try
        {
            var children = element.FindAll(TreeScope.TreeScope_Children, GetViewCondition(options.View));
            return children.Length > 0;
        }
        catch
        {
            return false;
        }
    }

    private IEnumerable<(IUIAutomationElement Child, string Label, List<string>? PathSegments)> GetChildren(
        IUIAutomationElement element,
        NativeUiaTreeDiagnosticOptions options,
        List<string>? pathSegments,
        string segmentLabel,
        int childDepth,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        StrongBox<bool> truncated,
        List<string> truncationWarnings)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "get-children");

        IUIAutomationElementArray children;
        try
        {
            children = element.FindAll(TreeScope.TreeScope_Children, GetViewCondition(options.View));
        }
        catch
        {
            yield break;
        }

        // maxChildren is informational truncation only — never stopTraversal.
        // Callers must keep visiting the yielded (capped) children and their siblings.
        var warning = NativeUiaSemanticContract.MaxChildrenTruncationWarning(
            children.Length,
            options.MaxChildrenPerNode);
        if (warning is not null)
        {
            truncated.Value = true;
            if (!truncationWarnings.Contains(warning, StringComparer.Ordinal))
                truncationWarnings.Add(warning);
        }

        var childCount = Math.Min(children.Length, options.MaxChildrenPerNode);
        var currentPath = options.IncludePath
            ? AppendPathSegment(pathSegments, segmentLabel)
            : null;

        for (var i = 0; i < childCount; i++)
        {
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "get-children loop");
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

            var childLabel = DescribeChildLabel(child, i);
            yield return (child, childLabel, currentPath);
        }
    }

    private Dictionary<string, object?> BuildElementRecord(
        IUIAutomationElement element,
        int depth,
        string? path,
        NativeUiaTreeDiagnosticOptions options,
        int? nodeIndex,
        int? parentNodeIndex,
        int? ordinal)
    {
        var controlTypeId = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
        var snapshot = _uia.CreateSnapshot(element);
        var (patterns, valueIsReadOnly, interaction) = ReadCheapPatterns(element);

        var record = new Dictionary<string, object?>
        {
            ["depth"] = depth,
            ["name"] = BoundString(snapshot.Name),
            ["automationId"] = BoundString(snapshot.AutomationId),
            ["controlTypeId"] = controlTypeId,
            ["controlType"] = BoundString(snapshot.ControlType),
            ["className"] = BoundString(snapshot.ClassName),
            ["frameworkId"] = BoundString(snapshot.FrameworkId),
            ["processId"] = snapshot.ProcessId,
            ["nativeWindowHandle"] = snapshot.NativeWindowHandle,
            ["isEnabled"] = snapshot.IsEnabled,
            ["isOffscreen"] = snapshot.IsOffscreen,
            ["supportedPatterns"] = patterns,
            ["valueIsReadOnly"] = valueIsReadOnly,
            ["capabilities"] = NativeUiaSemanticContract.MapCapabilities(patterns, valueIsReadOnly)
        };

        if (interaction.Count > 0)
            record["interaction"] = interaction;

        if (options.IncludeBounds)
            record["boundingRectangle"] = NativeUiaDiagnostics.ToRectangleObject(snapshot.BoundingRectangle);

        if (options.IncludePath)
            record["path"] = path ?? string.Empty;

        if (options.IncludeHelpText)
        {
            record["helpText"] = BoundString(
                _uia.GetStringProperty(element, UIA_PropertyIds.UIA_HelpTextPropertyId));
        }

        if (options.IncludeFocusState)
        {
            record["isKeyboardFocusable"] = _uia.GetBoolProperty(
                element,
                UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId);
            record["hasKeyboardFocus"] = _uia.GetBoolProperty(
                element,
                UIA_PropertyIds.UIA_HasKeyboardFocusPropertyId);
        }

        if (options.IncludeRuntimeId)
            record["runtimeId"] = BoundString(snapshot.RuntimeId);

        if (options.IncludeTreeLinks)
        {
            record["nodeIndex"] = nodeIndex;
            record["parentNodeIndex"] = parentNodeIndex;
            record["ordinal"] = ordinal;
        }

        if (options.IncludeWindowState && controlTypeId == WindowControlTypeId)
            record["windowState"] = ReadWindowState(element);

        return record;
    }

    private (List<string> Patterns, bool? ValueIsReadOnly, Dictionary<string, object?> Interaction) ReadCheapPatterns(
        IUIAutomationElement element)
    {
        var patterns = new List<string>();
        bool? valueIsReadOnly = null;
        var interaction = new Dictionary<string, object?>(StringComparer.Ordinal);

        foreach (var patternId in CheapPatternIds)
        {
            try
            {
                var pattern = element.GetCurrentPattern(patternId);
                if (pattern is null)
                    continue;

                patterns.Add(PatternName(patternId));
                if (patternId == UIA_PatternIds.UIA_ValuePatternId
                    && pattern is IUIAutomationValuePattern valuePattern)
                {
                    try
                    {
                        valueIsReadOnly = valuePattern.CurrentIsReadOnly != 0;
                    }
                    catch
                    {
                        // Fail closed: unknown read-only state omits setValue.
                        valueIsReadOnly = null;
                    }
                }
                else if (patternId == UIA_PatternIds.UIA_ExpandCollapsePatternId
                         && pattern is IUIAutomationExpandCollapsePattern expandCollapse)
                {
                    try
                    {
                        interaction["expandCollapse"] = expandCollapse.CurrentExpandCollapseState switch
                        {
                            ExpandCollapseState.ExpandCollapseState_Expanded => "expanded",
                            ExpandCollapseState.ExpandCollapseState_Collapsed => "collapsed",
                            ExpandCollapseState.ExpandCollapseState_PartiallyExpanded => "partiallyExpanded",
                            ExpandCollapseState.ExpandCollapseState_LeafNode => "leaf",
                            _ => "unknown"
                        };
                    }
                    catch
                    {
                        // ignore unstable pattern reads
                    }
                }
                else if (patternId == UIA_PatternIds.UIA_SelectionItemPatternId
                         && pattern is IUIAutomationSelectionItemPattern selectionItem)
                {
                    try
                    {
                        interaction["selected"] = selectionItem.CurrentIsSelected != 0;
                    }
                    catch
                    {
                        // ignore
                    }
                }
                else if (patternId == UIA_PatternIds.UIA_TogglePatternId
                         && pattern is IUIAutomationTogglePattern toggle)
                {
                    try
                    {
                        interaction["toggle"] = toggle.CurrentToggleState switch
                        {
                            ToggleState.ToggleState_On => "on",
                            ToggleState.ToggleState_Off => "off",
                            ToggleState.ToggleState_Indeterminate => "indeterminate",
                            _ => "unknown"
                        };
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        return (patterns, valueIsReadOnly, interaction);
    }

    private object? ReadWindowState(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId) is not IUIAutomationWindowPattern window)
                return null;

            return new Dictionary<string, object?>
            {
                ["isModal"] = window.CurrentIsModal != 0,
                ["canMaximize"] = window.CurrentCanMaximize != 0,
                ["canMinimize"] = window.CurrentCanMinimize != 0,
                ["windowInteractionState"] = window.CurrentWindowInteractionState.ToString(),
                ["windowVisualState"] = window.CurrentWindowVisualState.ToString()
            };
        }
        catch
        {
            return null;
        }
    }

    private static string PatternName(int patternId) => patternId switch
    {
        UIA_PatternIds.UIA_InvokePatternId => "Invoke",
        UIA_PatternIds.UIA_ExpandCollapsePatternId => "ExpandCollapse",
        UIA_PatternIds.UIA_SelectionItemPatternId => "SelectionItem",
        UIA_PatternIds.UIA_TogglePatternId => "Toggle",
        UIA_PatternIds.UIA_ValuePatternId => "Value",
        UIA_PatternIds.UIA_SelectionPatternId => "Selection",
        UIA_PatternIds.UIA_ScrollPatternId => "Scroll",
        UIA_PatternIds.UIA_GridPatternId => "Grid",
        UIA_PatternIds.UIA_WindowPatternId => "Window",
        _ => $"Pattern({patternId})"
    };

    private static string BoundString(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        return value.Length <= MaxDiagnosticStringLength
            ? value
            : value[..MaxDiagnosticStringLength];
    }

    private bool ElementMatchesLocator(
        IUIAutomationElement element,
        NativeUiaLocator locator,
        NativeUiaTreeDiagnosticOptions options)
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

        if (!string.IsNullOrWhiteSpace(options.NameContains)
            && !NativeUiaText.Matches(snapshot.Name, options.NameContains, "contains"))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.AutomationId)
            && !NativeUiaText.Matches(snapshot.AutomationId, locator.AutomationId, locator.MatchMode))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ClassName)
            && !NativeUiaText.Matches(snapshot.ClassName, locator.ClassName, locator.MatchMode))
            return false;

        return true;
    }

    private static bool ShouldIncludeElement(IUIAutomationElement element, NativeUiaTreeDiagnosticOptions options)
    {
        if (options.IncludeOffscreen)
            return true;

        try
        {
            var value = element.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId);
            return value is not bool offscreen || !offscreen;
        }
        catch
        {
            return true;
        }
    }

    private static bool IsMenuLikeControlType(int controlTypeId) =>
        MenuLikeControlTypeIds.Contains(controlTypeId);

    private IUIAutomationCondition GetViewCondition(string view) =>
        view.ToLowerInvariant() switch
        {
            "content" => _uia.ContentViewCondition,
            "raw" => _uia.TrueCondition(),
            _ => _uia.ControlViewCondition
        };

    private List<IUIAutomationElement> FindTopLevelWindowsForProcess(
        int processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "process-windows");

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

        var childCount = Math.Min(children.Length, MaxDesktopChildren);
        for (var i = 0; i < childCount; i++)
        {
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "process-windows loop");
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

    private string DescribeChildLabel(IUIAutomationElement child, int index)
    {
        var name = _uia.GetStringProperty(child, UIA_PropertyIds.UIA_NamePropertyId);
        var automationId = _uia.GetStringProperty(child, UIA_PropertyIds.UIA_AutomationIdPropertyId);
        var controlTypeId = _uia.GetIntProperty(child, UIA_PropertyIds.UIA_ControlTypePropertyId);
        var controlType = NativeUiaText.ControlTypeName(controlTypeId);

        if (!string.IsNullOrWhiteSpace(name))
            return $"{controlType}[{index}]/{name}";

        if (!string.IsNullOrWhiteSpace(automationId))
            return $"{controlType}[{index}]#{automationId}";

        return $"{controlType}[{index}]";
    }

    private static bool HasFindCriteria(UiRequest request, NativeUiaTreeDiagnosticOptions options, NativeUiaLocator locator) =>
        !string.IsNullOrWhiteSpace(options.NameContains)
        || !string.IsNullOrWhiteSpace(locator.Name)
        || !string.IsNullOrWhiteSpace(locator.AutomationId)
        || !string.IsNullOrWhiteSpace(locator.ClassName)
        || !string.IsNullOrWhiteSpace(locator.ControlType)
        || locator.Hwnd is > 0
        || request.Hwnd is > 0;

    private static NativeUiaTreeDiagnosticOptions ParseOptions(UiRequest request)
    {
        var view = request.View ?? request.TreeView ?? "control";
        var root = request.Root ?? request.SearchRoot ?? "activeWindow";

        return new NativeUiaTreeDiagnosticOptions
        {
            View = view,
            RootMode = NormalizeRootMode(root),
            MaxDepth = Math.Clamp(request.MaxDepth ?? request.Depth ?? DefaultMaxDepth, 0, 20),
            MaxChildrenPerNode = Math.Clamp(request.MaxChildren ?? request.Limit ?? DefaultMaxChildrenPerNode, 1, 1000),
            // Unspecified MaxNodes keeps the historical hard cap so capture-page stays compatible.
            MaxNodes = Math.Clamp(request.MaxNodes ?? MaxTotalNodes, 1, MaxTotalNodes),
            IncludeOffscreen = request.IncludeOffscreen ?? false,
            IncludePath = request.IncludePath ?? request.IncludeIdentifiers ?? true,
            IncludeBounds = request.IncludeBounds ?? true,
            IncludeHelpText = request.IncludeHelpText ?? false,
            IncludeFocusState = request.IncludeFocusState ?? false,
            IncludeRuntimeId = request.IncludeRuntimeId ?? false,
            IncludeTreeLinks = request.IncludeTreeLinks ?? false,
            IncludeWindowState = request.IncludeWindowState ?? false,
            NameContains = request.NameContains ?? request.BestMatch,
            ControlType = request.Locator?.ControlType,
            ClassName = request.ClassName ?? request.Locator?.ClassName,
            AutomationId = request.Locator?.AutomationId
        };
    }

    private static string NormalizeRootMode(string root) =>
        root.Trim().ToLowerInvariant() switch
        {
            "activewindow" or "active-window" or "active_window" or "currentwindow" or "current-window" => "activeWindow",
            "processwindows" or "process-windows" or "process_windows" => "processWindows",
            "desktopchildren" or "desktop-children" or "desktop_children" or "desktop" => "desktopChildren",
            _ => root
        };

    private static NativeUiaLocator ToNativeLocator(UiRequest request, int? processId) => new()
    {
        Name = request.Locator?.Name,
        AutomationId = request.Locator?.AutomationId,
        ClassName = request.ClassName ?? request.Locator?.ClassName,
        ControlType = request.Locator?.ControlType,
        Hwnd = request.Hwnd ?? request.Locator?.Hwnd ?? request.Locator?.Handle,
        ProcessId = request.ProcessId ?? request.Locator?.ProcessId ?? processId,
        FoundIndex = request.FoundIndex ?? request.Locator?.FoundIndex ?? request.Locator?.Index,
        MatchMode = string.IsNullOrWhiteSpace(request.MatchMode)
            ? request.Locator?.MatchMode ?? "exact"
            : request.MatchMode!
    };

    private static string DescribeLocator(NativeUiaLocator locator)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            parts.Add($"automationId={locator.AutomationId}");
        if (!string.IsNullOrWhiteSpace(locator.Name))
            parts.Add($"name={locator.Name}");
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
            parts.Add($"controlType={locator.ControlType}");
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
            parts.Add($"className={locator.ClassName}");
        if (!string.IsNullOrWhiteSpace(locator.MatchMode))
            parts.Add($"matchMode={locator.MatchMode}");

        return parts.Count == 0 ? string.Empty : string.Join(", ", parts);
    }

    private static string? BuildPath(List<string>? pathSegments, string segmentLabel, bool includePath)
    {
        if (!includePath)
            return null;

        if (pathSegments == null || pathSegments.Count == 0)
            return segmentLabel;

        return string.Join(" > ", pathSegments.Concat([segmentLabel]));
    }

    private static List<string>? AppendPathSegment(List<string>? pathSegments, string segmentLabel)
    {
        if (pathSegments == null)
            return [segmentLabel];

        var copy = new List<string>(pathSegments) { segmentLabel };
        return copy;
    }

    private static bool IsDeadlineNear(DateTime deadlineUtc) =>
        DateTime.UtcNow.AddMilliseconds(DeadlineSafetyMs) >= deadlineUtc;

    private static void EnsureWithinDeadline(DateTime deadlineUtc, CancellationToken cancellationToken, string stage)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow > deadlineUtc)
            throw new TimeoutException($"{stage} exceeded timeout.");
    }

    private static int ResolveTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }

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

    private sealed record NativeUiaTreeDiagnosticOptions
    {
        public string View { get; init; } = "control";
        public string RootMode { get; init; } = "activeWindow";
        public int MaxDepth { get; init; } = DefaultMaxDepth;
        public int MaxChildrenPerNode { get; init; } = DefaultMaxChildrenPerNode;
        public int MaxNodes { get; init; } = MaxTotalNodes;
        public bool IncludeOffscreen { get; init; }
        public bool IncludePath { get; init; } = true;
        public bool IncludeBounds { get; init; } = true;
        public bool IncludeHelpText { get; init; }
        public bool IncludeFocusState { get; init; }
        public bool IncludeRuntimeId { get; init; }
        public bool IncludeTreeLinks { get; init; }
        public bool IncludeWindowState { get; init; }
        public string? NameContains { get; init; }
        public string? ControlType { get; init; }
        public string? ClassName { get; init; }
        public string? AutomationId { get; init; }
    }

    private sealed class DumpFilter
    {
        public string? NameContains { get; init; }
        public string? ControlType { get; init; }
        public string? ClassName { get; init; }
        public string? AutomationId { get; init; }

        public bool Matches(IDictionary<string, object?> record)
        {
            if (!HasAnyFilter())
                return true;

            if (!string.IsNullOrWhiteSpace(NameContains)
                && record.TryGetValue("name", out var nameObj)
                && !NativeUiaText.Matches(nameObj?.ToString(), NameContains, "contains"))
                return false;

            if (!string.IsNullOrWhiteSpace(ControlType)
                && record.TryGetValue("controlType", out var controlTypeObj)
                && !NativeUiaText.Matches(controlTypeObj?.ToString(), ControlType, "exact"))
                return false;

            if (!string.IsNullOrWhiteSpace(ClassName)
                && record.TryGetValue("className", out var classNameObj)
                && !NativeUiaText.Matches(classNameObj?.ToString(), ClassName, "exact"))
                return false;

            if (!string.IsNullOrWhiteSpace(AutomationId)
                && record.TryGetValue("automationId", out var automationIdObj)
                && !NativeUiaText.Matches(automationIdObj?.ToString(), AutomationId, "exact"))
                return false;

            return true;
        }

        private bool HasAnyFilter() =>
            !string.IsNullOrWhiteSpace(NameContains)
            || !string.IsNullOrWhiteSpace(ControlType)
            || !string.IsNullOrWhiteSpace(ClassName)
            || !string.IsNullOrWhiteSpace(AutomationId);
    }

    private sealed class StrongBox<T>(T value)
    {
        public T Value { get; set; } = value;
    }
}
