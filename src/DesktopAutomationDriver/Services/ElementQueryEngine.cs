using System.Diagnostics;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Central pywinauto-style element resolver — Phase 1.
/// Extends <see cref="UiService"/> as a partial class.
/// </summary>
public partial class UiService
{
    // =========================================================================
    // Central resolver entry points
    // =========================================================================

    internal static ITreeWalker GetTreeWalker(FlaUI.Core.AutomationBase automation, string? treeView)
    {
        var view = treeView?.ToLowerInvariant() ?? "control";
        return view switch
        {
            "content" => automation.TreeWalkerFactory.GetContentViewWalker(),
            "raw" => automation.TreeWalkerFactory.GetRawViewWalker(),
            _ => automation.TreeWalkerFactory.GetControlViewWalker()
        };
    }

    internal static List<AutomationElement> FindDescendantsWithWalker(
        AutomationElement root, int maxDepth, ITreeWalker walker)
    {
        var result = new List<AutomationElement>();
        if (maxDepth < 0) return result;

        var queue = new Queue<(AutomationElement Element, int Depth)>();
        queue.Enqueue((root, 0));

        while (queue.Count > 0)
        {
            var (current, currentDepth) = queue.Dequeue();
            if (currentDepth > maxDepth) continue;

            var children = new List<AutomationElement>();
            try
            {
                var child = walker.GetFirstChild(current);
                while (child != null)
                {
                    children.Add(child);
                    child = walker.GetNextSibling(child);
                }
            }
            catch
            {
                continue;
            }

            foreach (var child in children)
            {
                result.Add(child);
                if (currentDepth < maxDepth)
                    queue.Enqueue((child, currentDepth + 1));
            }
        }

        return result;
    }

    private AutomationElement GetForegroundWindowElement(FlaUI.Core.AutomationBase automation)
    {
        var hwnd = GetForegroundWindow();
        if (hwnd != IntPtr.Zero)
        {
            try
            {
                var el = automation.FromHandle(hwnd);
                if (el != null) return el;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to resolve element from foreground HWND.");
            }
        }
        return automation.GetDesktop();
    }

    private AutomationElement GetActivePopupRoot(AutomationSession session)
    {
        var desktop = session.Automation.GetDesktop();
        try
        {
            var children = desktop.FindAllChildren();
            foreach (var child in children)
            {
                var ct = child.ControlType;
                if ((ct == ControlType.Window || ct == ControlType.Menu || ct == ControlType.List) &&
                    SafeIsOffscreen(child) == false)
                {
                    var cn = SafeElementClassName(child);
                    var aid = SafeElementAutomationId(child);
                    if (cn == "#32768" || cn == "ComboLBox" || cn.Contains("Popup", StringComparison.OrdinalIgnoreCase) || aid.Contains("Popup", StringComparison.OrdinalIgnoreCase))
                    {
                        return child;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to walk desktop children for active popup discovery.");
        }

        return GetForegroundWindowElement(session.Automation);
    }

    internal static int CalculateBestMatchScore(string input, string pattern)
    {
        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(pattern))
            return 0;
        if (string.Equals(input, pattern, StringComparison.OrdinalIgnoreCase))
            return 100;
        if (input.StartsWith(pattern, StringComparison.OrdinalIgnoreCase))
            return 80;
        if (input.Contains(pattern, StringComparison.OrdinalIgnoreCase))
            return 60;
        
        var inputWords = input.Split(' ', '_', '-');
        var patternWords = pattern.Split(' ', '_', '-');
        int matches = 0;
        foreach (var pw in patternWords)
        {
            if (inputWords.Any(iw => string.Equals(iw, pw, StringComparison.OrdinalIgnoreCase)))
                matches++;
        }
        if (matches > 0)
            return 50 + (matches * 10);
        return 0;
    }

    /// <summary>
    /// Single-shot element resolution. Validates the locator, determines the search
    /// root, gathers candidates, filters and scores them, applies indexing, and returns
    /// an <see cref="ElementResolveResult"/> that carries either the found element or
    /// detailed not-found / ambiguity diagnostics.
    /// </summary>
    private ElementResolveResult ResolveElement(
        UiRequest request,
        UiLocator? locator,
        AutomationElement? explicitRoot = null,
        string purpose = "generic",
        bool allowDesktopSearch = false,
        bool allowOffscreen = true,
        bool returnAllCandidates = false)
    {
        if (locator == null && (request.LocatorPath == null || request.LocatorPath.Count == 0))
            return new ElementResolveResult
            {
                Strategy = "null-locator",
                Errors   = ["Locator is null and locatorPath is empty."]
            };

        var session = RequireSession();

        // ── 1. Determine root ───────────────────────────────────────────────
        string rootStrategy;
        AutomationElement root;

        var searchRootName = request.SearchRoot;
        if (explicitRoot != null)
        {
            root         = explicitRoot;
            rootStrategy = "explicit-root";
        }
        else if (!string.IsNullOrWhiteSpace(searchRootName))
        {
            rootStrategy = searchRootName.ToLowerInvariant();
            if (string.Equals(searchRootName, "desktop", StringComparison.OrdinalIgnoreCase) || allowDesktopSearch)
            {
                root = session.Automation.GetDesktop();
            }
            else if (string.Equals(searchRootName, "foreground", StringComparison.OrdinalIgnoreCase))
            {
                root = GetForegroundWindowElement(session.Automation);
            }
            else if (string.Equals(searchRootName, "activepopup", StringComparison.OrdinalIgnoreCase))
            {
                root = GetActivePopupRoot(session);
            }
            else if (string.Equals(searchRootName, "parent", StringComparison.OrdinalIgnoreCase))
            {
                if (request.ParentLocator != null)
                {
                    var windowRoot = GetWindowRoot(session, allowDesktopPopupScan: false);
                    var parentSearch = TryFindElementBySmartStrategy(
                        windowRoot, session, request.ParentLocator,
                        preferAttributes: true, xpathOnly: false);
                    root = parentSearch.Element ?? windowRoot;
                }
                else
                {
                    root = GetWindowRoot(session, allowDesktopPopupScan: false);
                }
            }
            else
            {
                root = GetWindowRoot(session, allowDesktopPopupScan: false);
            }
        }
        else if (request.ParentLocator != null)
        {
            var windowRoot    = GetWindowRoot(session, allowDesktopPopupScan: false);
            var parentSearch  = TryFindElementBySmartStrategy(
                windowRoot, session, request.ParentLocator,
                preferAttributes: true, xpathOnly: false);

            if (parentSearch.Element == null)
                return new ElementResolveResult
                {
                    RootStrategy = "parent-locator-not-found",
                    Strategy     = "parent-locator-not-found",
                    Locator      = locator,
                    Errors       = [$"Parent locator not found: {DescribeLocator(request.ParentLocator)}"]
                };

            root         = parentSearch.Element;
            rootStrategy = "parent-locator";
        }
        else if (request.UseDesktopRoot == true)
        {
            root         = session.Automation.GetDesktop();
            rootStrategy = "desktop";
        }
        else if (request.UseActiveWindowRoot == true)
        {
            root         = GetWindowRoot(session, allowDesktopPopupScan: false);
            rootStrategy = "active-window";
        }
        else
        {
            root         = GetWindowRoot(session, allowDesktopPopupScan: false);
            rootStrategy = session.ActiveWindow != null ? "active-window" : "app-main-window";
        }

        // ── 2. Fast path for simple locators (no new-style fields or request options) ─────────
        bool isNewStyleRequest = (request.LocatorPath != null && request.LocatorPath.Count > 0) ||
                                 !string.IsNullOrWhiteSpace(request.SearchRoot) ||
                                 !string.IsNullOrWhiteSpace(request.TreeView) ||
                                 !string.IsNullOrWhiteSpace(request.Backend) ||
                                 request.ReturnCandidates == true ||
                                 request.Debug == true ||
                                 !string.IsNullOrWhiteSpace(request.Ambiguity);

        if (locator != null && !isNewStyleRequest && !NeedsNewStyleSearch(locator))
        {
            var xpathOnly        = request.XPathOnly == true;
            var preferAttributes = ShouldPreferAttributeSearch(request);

            var result = TryFindElementBySmartStrategy(root, session, locator, preferAttributes, xpathOnly);

            if (result.Element != null)
                return new ElementResolveResult
                {
                    Element        = result.Element,
                    Strategy       = result.Strategy,
                    RootStrategy   = rootStrategy,
                    Locator        = locator,
                    CandidateCount = 1
                };

            // ComboBox relaxed fallback
            if (string.Equals(locator.ControlType, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(locator.AutomationId))
            {
                var cf        = session.Automation.ConditionFactory;
                var candidate = root.FindFirstDescendant(cf.ByAutomationId(locator.AutomationId));
                if (candidate != null)
                {
                    if (candidate.ControlType == ControlType.ComboBox)
                        return new ElementResolveResult
                        {
                            Element        = candidate,
                            Strategy       = "combobox-automationid-relaxed",
                            RootStrategy   = rootStrategy,
                            Locator        = locator,
                            CandidateCount = 1
                        };

                    var ancestor = TryFindComboBoxAncestor(candidate);
                    if (ancestor != null)
                        return new ElementResolveResult
                        {
                            Element        = ancestor,
                            Strategy       = "combobox-ancestor-relaxed",
                            RootStrategy   = rootStrategy,
                            Locator        = locator,
                            CandidateCount = 1
                        };
                }
            }

            return new ElementResolveResult
            {
                RootStrategy = rootStrategy,
                Strategy     = result.Strategy,
                Locator      = locator,
                Errors       = [$"ElementNotFound: Element not found. strategy={result.Strategy}, locator={DescribeLocator(locator)}"]
            };
        }

        // ── 3. New-style: resolve elements or path ────────────────────────
        List<AutomationElement> finalCandidates;
        if (request.LocatorPath != null && request.LocatorPath.Count > 0)
        {
            finalCandidates = ResolveLocatorPath(request, request.LocatorPath, root);
        }
        else
        {
            finalCandidates = ResolveElements(request, locator ?? new UiLocator(), root);
        }

        int rawCollectedCount = finalCandidates.Count;
        bool includeCandidates = returnAllCandidates || request.ReturnCandidates == true || request.Debug == true || request.IncludeDiagnostics == true;

        if (rawCollectedCount == 0)
        {
            var emptyLoc = locator ?? (request.LocatorPath != null && request.LocatorPath.Count > 0 ? request.LocatorPath[^1] : new UiLocator());
            var rawCandidates = BuildCandidateList(root, emptyLoc, session, request);
            var diagnostics = BuildNearMatchDiagnostics(rawCandidates, emptyLoc);

            return new ElementResolveResult
            {
                RootStrategy   = rootStrategy,
                Strategy       = "filtered-not-found",
                Locator        = locator,
                CandidateCount = rawCandidates.Count,
                Candidates     = diagnostics,
                Errors         = [$"ElementNotFound: No matching elements found. collected={rawCandidates.Count}, locator={DescribeLocator(emptyLoc)}"]
            };
        }

        // Check for ambiguity
        var ambiguityMode = request.Ambiguity ?? "error";
        if (rawCollectedCount > 1)
        {
            if (string.Equals(ambiguityMode, "error", StringComparison.OrdinalIgnoreCase))
            {
                var emptyLoc = locator ?? (request.LocatorPath != null && request.LocatorPath.Count > 0 ? request.LocatorPath[^1] : new UiLocator());
                var diagnostics = BuildNearMatchDiagnostics(finalCandidates, emptyLoc);
                return new ElementResolveResult
                {
                    RootStrategy   = rootStrategy,
                    Strategy       = "ambiguity-error",
                    Locator        = locator,
                    CandidateCount = rawCollectedCount,
                    Ambiguous      = true,
                    Candidates     = diagnostics,
                    Errors         = [$"ElementAmbiguous: Multiple matching elements found ({rawCollectedCount})."]
                };
            }
        }

        var picked = finalCandidates[0];
        var scored = ScoreAndSortCandidates(finalCandidates, locator ?? new UiLocator());

        return new ElementResolveResult
        {
            Element        = picked,
            Strategy       = rawCollectedCount > 1 ? "scored-ambiguous" : "scored-unique",
            RootStrategy   = rootStrategy,
            Locator        = locator,
            CandidateCount = rawCollectedCount,
            Ambiguous      = rawCollectedCount > 1,
            Candidates     = includeCandidates
                ? scored.Select((s, i) => BuildCandidateDto(s.Element, i, s.Score, s.Reason)).ToList()
                : []
        };
    }

    private List<AutomationElement> ResolveElements(
        UiRequest request,
        UiLocator locator,
        AutomationElement startRoot)
    {
        var session = RequireSession();

        // 1. BuildCandidateList
        var candidates = BuildCandidateList(startRoot, locator, session, request);

        // 2. CtrlIndex raw pre-filter
        if (locator.CtrlIndex.HasValue)
        {
            candidates = ApplyIndexFilters(candidates, locator, isRawPreFilter: true);
            return candidates;
        }

        // 3. Apply filters
        candidates = ApplyLocatorFilters(candidates, locator);
        candidates = ApplyRegexFilters(candidates, locator);
        candidates = ApplyValueFilters(candidates, locator);
        candidates = ApplyStateFilters(candidates, locator);

        // 4. Score and Sort
        var scored = ScoreAndSortCandidates(candidates, locator);
        var sortedCandidates = scored.Select(s => s.Element).ToList();

        // 5. BestMatch scoring & picking highest above threshold
        var bestMatchPattern = locator.BestMatch ?? request.BestMatch;
        if (!string.IsNullOrWhiteSpace(bestMatchPattern) && sortedCandidates.Count > 0)
        {
            var bestScored = sortedCandidates
                .Select(c => new { Element = c, Score = CalculateBestMatchScore(SafeElementName(c), bestMatchPattern) })
                .Where(x => x.Score >= BestMatchMinimumThreshold) // threshold = 50
                .OrderByDescending(x => x.Score)
                .ToList();

            if (bestScored.Count > 0)
            {
                sortedCandidates = bestScored.Select(x => x.Element).ToList();
            }
        }

        // 6. FoundIndex post-filter
        if (locator.FoundIndex.HasValue)
        {
            sortedCandidates = ApplyIndexFilters(sortedCandidates, locator, isRawPreFilter: false);
        }

        return sortedCandidates;
    }

    private List<AutomationElement> ResolveLocatorPath(
        UiRequest request,
        List<UiLocator> path,
        AutomationElement startRoot)
    {
        var currentRoots = new List<AutomationElement> { startRoot };

        for (int i = 0; i < path.Count; i++)
        {
            var locator = path[i];
            var nextRoots = new List<AutomationElement>();

            foreach (var root in currentRoots)
            {
                var resolved = ResolveElements(request, locator, root);
                nextRoots.AddRange(resolved);
            }

            currentRoots = nextRoots.Distinct().ToList();

            if (currentRoots.Count == 0)
                break;
        }

        return currentRoots;
    }

    private List<AutomationElement> BuildCandidateList(
        AutomationElement root,
        UiLocator locator,
        AutomationSession session,
        UiRequest request)
    {
        var depth = locator.Depth ?? request.Locator?.Depth ?? DefaultCandidateSearchDepth;
        var topLevelOnly = locator.TopLevelOnly == true;

        // ── HWND direct lookup ───────────────────────────────────────────────
        if (locator.Hwnd.HasValue)
        {
            try
            {
                var hwndPtr = new IntPtr(locator.Hwnd.Value);
                var element = session.Automation.FromHandle(hwndPtr);
                return element == null ? [] : [element];
            }
            catch
            {
                return [];
            }
        }

        // ── XPath ────────────────────────────────────────────────────────────
        if (!string.IsNullOrWhiteSpace(locator.XPath))
        {
            var byXPath = FindByXPath(root, session, locator.XPath);
            return byXPath == null ? [] : [byXPath];
        }

        // If request has treeView set (e.g. content or raw), use TreeWalker to fetch candidates
        if (!string.IsNullOrWhiteSpace(request.TreeView) && 
            !string.Equals(request.TreeView, "control", StringComparison.OrdinalIgnoreCase))
        {
            var walker = GetTreeWalker(session.Automation, request.TreeView);
            if (topLevelOnly)
            {
                var children = new List<AutomationElement>();
                try
                {
                    var child = walker.GetFirstChild(root);
                    while (child != null)
                    {
                        children.Add(child);
                        child = walker.GetNextSibling(child);
                    }
                }
                catch { }
                return children;
            }
            return FindDescendantsWithWalker(root, depth, walker);
        }

        // ── TopLevelOnly ─────────────────────────────────────────────────────
        if (topLevelOnly)
            return root.FindAllChildren().ToList();

        // ── ControlType-scoped descendants ──────────────────────────────────
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            try
            {
                var ct = ParseControlType(locator.ControlType);
                return root.FindAllDescendants(
                    session.Automation.ConditionFactory.ByControlType(ct)).ToList();
            }
            catch
            {
                return [];
            }
        }

        // ── All descendants with optional depth cap ──────────────────────────
        return FindDescendantsUpToDepth(root, depth);
    }

    private List<AutomationElement> ApplyLocatorFilters(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var filtered = new List<AutomationElement>();
        foreach (var c in candidates)
        {
            // controlType
            if (!string.IsNullOrWhiteSpace(locator.ControlType))
            {
                var ct = SafeElementControlType(c);
                if (!string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            // className (non-regex)
            if (!string.IsNullOrWhiteSpace(locator.ClassName))
            {
                var cn = SafeElementClassName(c);
                var mode = locator.ClassNameMatchMode ?? locator.MatchMode;
                if (!StringMatchesByMode(cn, locator.ClassName, mode, caseSensitive: false))
                    continue;
            }

            // processId
            if (locator.ProcessId.HasValue)
            {
                var pid = SafeProcessId(c);
                if (pid != locator.ProcessId.Value)
                    continue;
            }

            // frameworkId
            if (!string.IsNullOrWhiteSpace(locator.FrameworkId))
            {
                var fwId = SafeFrameworkId(c);
                if (!string.Equals(fwId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            // runtimeId
            if (!string.IsNullOrWhiteSpace(locator.RuntimeId))
            {
                var rtId = SafeRuntimeIdString(c);
                if (!string.Equals(rtId, locator.RuntimeId, StringComparison.Ordinal))
                    continue;
            }

            // name (non-regex)
            if (!string.IsNullOrWhiteSpace(locator.Name))
            {
                var name = SafeElementName(c);
                var mode = locator.NameMatchMode ?? locator.MatchMode;
                if (!StringMatchesByMode(name, locator.Name, mode, caseSensitive: true))
                    continue;
            }

            // automationId (non-regex)
            if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            {
                var aid = SafeElementAutomationId(c);
                var mode = locator.AutomationIdMatchMode ?? locator.MatchMode;
                if (!StringMatchesByMode(aid, locator.AutomationId, mode, caseSensitive: true))
                    continue;
            }

            filtered.Add(c);
        }
        return filtered;
    }

    private List<AutomationElement> ApplyRegexFilters(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var filtered = new List<AutomationElement>();
        foreach (var c in candidates)
        {
            // nameRegex
            if (!string.IsNullOrWhiteSpace(locator.NameRegex))
            {
                var name = SafeElementName(c);
                if (!SafeRegexIsMatch(name, locator.NameRegex))
                    continue;
            }

            // automationIdRegex
            if (!string.IsNullOrWhiteSpace(locator.AutomationIdRegex))
            {
                var aid = SafeElementAutomationId(c);
                if (!SafeRegexIsMatch(aid, locator.AutomationIdRegex))
                    continue;
            }

            // classNameRegex
            if (!string.IsNullOrWhiteSpace(locator.ClassNameRegex))
            {
                var cn = SafeElementClassName(c);
                if (!SafeRegexIsMatch(cn, locator.ClassNameRegex))
                    continue;
            }

            filtered.Add(c);
        }
        return filtered;
    }

    private List<AutomationElement> ApplyValueFilters(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var filtered = new List<AutomationElement>();
        foreach (var c in candidates)
        {
            // value (non-regex)
            if (!string.IsNullOrWhiteSpace(locator.Value))
            {
                var val = SafeElementValue(c);
                var mode = locator.ValueMatchMode ?? locator.MatchMode;
                if (!StringMatchesByMode(val, locator.Value, mode, caseSensitive: false))
                    continue;
            }

            // valueRegex
            if (!string.IsNullOrWhiteSpace(locator.ValueRegex))
            {
                var val = SafeElementValue(c);
                if (!SafeRegexIsMatch(val, locator.ValueRegex))
                    continue;
            }

            // text
            if (!string.IsNullOrWhiteSpace(locator.Text))
            {
                var text = SafeElementText(c);
                var mode = locator.TextMatchMode ?? locator.MatchMode;
                if (!StringMatchesByMode(text, locator.Text, mode, caseSensitive: false))
                    continue;
            }

            filtered.Add(c);
        }
        return filtered;
    }

    private List<AutomationElement> ApplyStateFilters(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var filtered = new List<AutomationElement>();
        foreach (var c in candidates)
        {
            // visible / offscreen
            if (locator.Visible.HasValue)
            {
                var offscreen = SafeIsOffscreen(c) ?? false;
                var visible = !offscreen;
                if (visible != locator.Visible.Value)
                    continue;
            }

            if (locator.Offscreen.HasValue)
            {
                var offscreen = SafeIsOffscreen(c) ?? false;
                if (offscreen != locator.Offscreen.Value)
                    continue;
            }

            // enabled
            if (locator.Enabled.HasValue)
            {
                var enabled = SafeIsEnabled(c) ?? false;
                if (enabled != locator.Enabled.Value)
                    continue;
            }

            filtered.Add(c);
        }
        return filtered;
    }

    private List<AutomationElement> ApplyIndexFilters(
        List<AutomationElement> candidates,
        UiLocator locator,
        bool isRawPreFilter)
    {
        if (isRawPreFilter)
        {
            if (locator.CtrlIndex.HasValue)
            {
                var ctrlIdx = locator.CtrlIndex.Value;
                if (ctrlIdx >= 0 && ctrlIdx < candidates.Count)
                    return [candidates[ctrlIdx]];
                return [];
            }
        }
        else
        {
            if (locator.FoundIndex.HasValue)
            {
                var foundIdx = locator.FoundIndex.Value;
                if (foundIdx >= 0 && foundIdx < candidates.Count)
                    return [candidates[foundIdx]];
                return [];
            }
        }
        return candidates;
    }

    private List<ElementCandidateDto> BuildNearMatchDiagnostics(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        return BuildCandidateDtos(candidates.Take(10).ToList(), locator);
    }

    /// <summary>
    /// Resolves the element or throws <see cref="InvalidOperationException"/>.
    /// Wraps <see cref="ResolveElement"/> with the standard retry/timeout loop
    /// from the operation policy. Replaces <c>FindWithRetry</c> for operations
    /// that support the new pywinauto-style locator fields.
    /// </summary>
    private AutomationElement ResolveElementOrThrow(
        UiRequest request,
        UiLocator? locator,
        string purpose,
        bool allowDesktopSearch = false,
        bool allowOffscreen     = true)
    {
        if (locator == null && (request.LocatorPath == null || request.LocatorPath.Count == 0))
            throw new ArgumentException("'locator' or 'locatorPath' is required for this operation.");

        var session  = RequireSession();
        var policy   = GetOperationPolicy(request);
        var deadline = DateTime.UtcNow + policy.Timeout;

        var effectiveAllowDesktop = allowDesktopSearch || policy.AllowDesktopPopupScan;

        // Build cache key for fast operations (only when new-style fields are absent).
        var preferAttributes = ShouldPreferAttributeSearch(request);
        var xpathOnly        = request.XPathOnly == true;

        string? cacheKey = null;
        if (locator != null && policy.UseElementCache && !policy.RefreshRootEveryRetry && !NeedsNewStyleSearch(locator))
        {
            var root = GetWindowRoot(session, allowDesktopPopupScan: policy.AllowDesktopPopupScan);
            cacheKey = BuildElementCacheKey(
                session, root, locator, request.ParentLocator,
                preferAttributes: preferAttributes,
                xpathOnly: xpathOnly,
                preferXPath: request.PreferXPath == true);
        }

        if (cacheKey != null &&
            TryGetCachedElement(cacheKey, locator!, out var cached) &&
            cached != null)
        {
            _logger.LogInformation(
                "UI locator resolved (engine). operation={Operation}, strategy=cache, locator={Locator}",
                SanitizeValue(request.Operation),
                DescribeLocator(locator!));
            return cached;
        }

        var sw         = Stopwatch.StartNew();
        ElementResolveResult lastResult = new() { Strategy = "not-started" };

        while (true)
        {
            lastResult = ResolveElement(
                request, locator,
                purpose:            purpose,
                allowDesktopSearch: effectiveAllowDesktop,
                allowOffscreen:     allowOffscreen);

            if (lastResult.Element != null)
            {
                sw.Stop();

                if (cacheKey != null)
                    StoreCachedElement(cacheKey, lastResult.Element);

                _logger.LogInformation(
                    "UI locator resolved (engine). operation={Operation}, strategy={Strategy}, rootStrategy={RootStrategy}, elapsedMs={ElapsedMs}, locator={Locator}",
                    SanitizeValue(request.Operation),
                    lastResult.Strategy,
                    lastResult.RootStrategy,
                    sw.ElapsedMilliseconds,
                    DescribeLocator(locator ?? new UiLocator()));

                return lastResult.Element;
            }

            if (lastResult.Strategy == "ambiguity-error")
            {
                throw new InvalidOperationException(lastResult.Errors.FirstOrDefault() ?? "ElementAmbiguous: Multiple matching elements found.");
            }

            if (DateTime.UtcNow >= deadline)
            {
                sw.Stop();

                _logger.LogWarning(
                    "UI locator not found (engine). operation={Operation}, strategy={Strategy}, rootStrategy={RootStrategy}, elapsedMs={ElapsedMs}, locator={Locator}",
                    SanitizeValue(request.Operation),
                    lastResult.Strategy,
                    lastResult.RootStrategy,
                    sw.ElapsedMilliseconds,
                    DescribeLocator(locator ?? new UiLocator()));

                var specialErr = lastResult.Errors.FirstOrDefault(e => e.StartsWith("ElementNotFound") || e.StartsWith("ElementAmbiguous"));
                if (specialErr != null)
                {
                    throw new InvalidOperationException(specialErr);
                }

                throw new InvalidOperationException(BuildElementNotFoundMessage(lastResult, policy));
            }

            Thread.Sleep(policy.RetryInterval);

            if (policy.RefreshRootEveryRetry)
            {
                // Re-invalidate cache key after root refresh
                cacheKey = null;
            }
        }
    }

    // =========================================================================
    // Root determination helpers (used by the resolver)
    // =========================================================================

    /// <summary>
    /// Returns true when the locator has any field that requires the new-style
    /// candidate-gathering path rather than the existing FlaUI fast-attribute search.
    /// </summary>
    private static bool NeedsNewStyleSearch(UiLocator locator)
    {
        // New fields that bypass the existing fast path:
        if (locator.Hwnd.HasValue)            return true;
        if (locator.ProcessId.HasValue)       return true;
        if (locator.ControlId.HasValue)       return true;
        if (locator.FrameworkId != null)      return true;
        if (locator.RuntimeId != null)        return true;
        if (locator.Value != null)            return true;
        if (locator.ValueRegex != null)       return true;
        if (locator.Text != null)             return true;
        if (locator.Visible.HasValue)         return true;
        if (locator.Enabled.HasValue)         return true;
        if (locator.Offscreen.HasValue)       return true;
        if (locator.FoundIndex.HasValue)      return true;
        if (locator.CtrlIndex.HasValue)       return true;
        if (locator.Depth.HasValue)           return true;
        if (locator.TopLevelOnly == true)     return true;
        if (locator.ActiveOnly.HasValue)      return true;
        if (locator.IncludeOffscreen.HasValue) return true;
        if (locator.NameRegex != null)         return true;
        if (locator.AutomationIdRegex != null) return true;
        if (locator.ClassNameRegex != null)    return true;
        if (locator.BestMatch != null)         return true;

        // Non-exact match modes
        static bool IsNonExact(string? mode) =>
            mode != null && !string.Equals(mode, "exact", StringComparison.OrdinalIgnoreCase);

        if (IsNonExact(locator.MatchMode))             return true;
        if (IsNonExact(locator.NameMatchMode))         return true;
        if (IsNonExact(locator.AutomationIdMatchMode)) return true;
        if (IsNonExact(locator.ClassNameMatchMode))    return true;
        if (locator.ValueMatchMode != null)            return true;
        if (locator.TextMatchMode != null)             return true;

        return false;
    }

    // =========================================================================
    // Candidate collection
    // =========================================================================

    /// <summary>
    /// Gathers element candidates from <paramref name="root"/> according to the
    /// locator's HWND, XPath, ControlType, depth, and top-level-only settings.
    /// </summary>
    private List<AutomationElement> CollectCandidates(
        AutomationElement root,
        UiLocator locator,
        AutomationSession session,
        int? depth,
        bool topLevelOnly)
    {
        // ── HWND direct lookup ───────────────────────────────────────────────
        if (locator.Hwnd.HasValue)
        {
            try
            {
                var hwndPtr = new IntPtr(locator.Hwnd.Value);
                var element = session.Automation.FromHandle(hwndPtr);
                return element == null
                    ? []
                    : [element];
            }
            catch
            {
                return [];
            }
        }

        // ── XPath ────────────────────────────────────────────────────────────
        if (!string.IsNullOrWhiteSpace(locator.XPath))
        {
            var byXPath = FindByXPath(root, session, locator.XPath);
            return byXPath == null ? [] : [byXPath];
        }

        // ── TopLevelOnly ─────────────────────────────────────────────────────
        if (topLevelOnly)
            return root.FindAllChildren().ToList();

        // ── ControlType-scoped descendants ──────────────────────────────────
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            try
            {
                var ct = ParseControlType(locator.ControlType);
                return root.FindAllDescendants(
                    session.Automation.ConditionFactory.ByControlType(ct)).ToList();
            }
            catch
            {
                return [];
            }
        }

        // ── All descendants with optional depth cap ──────────────────────────
        return FindDescendantsUpToDepth(root, depth ?? DefaultCandidateSearchDepth);
    }

    // Default cap to avoid traversing the entire desktop tree unboundedly.
    private const int DefaultCandidateSearchDepth = 20;
    private const int BestMatchMinimumThreshold = 50;

    /// <summary>
    /// BFS traversal of the UIA tree up to <paramref name="maxDepth"/> levels.
    /// Depth 0 = direct children of <paramref name="root"/>.
    /// </summary>
    internal static List<AutomationElement> FindDescendantsUpToDepth(
        AutomationElement root, int maxDepth)
    {
        var result = new List<AutomationElement>();
        if (maxDepth < 0) return result;

        // BFS queue: (element, depth)
        var queue = new Queue<(AutomationElement Element, int Depth)>();
        queue.Enqueue((root, 0));

        while (queue.Count > 0)
        {
            var (current, currentDepth) = queue.Dequeue();
            if (currentDepth > maxDepth) continue;

            AutomationElement[] children;
            try { children = current.FindAllChildren(); }
            catch { continue; }

            foreach (var child in children)
            {
                result.Add(child);
                if (currentDepth < maxDepth)
                    queue.Enqueue((child, currentDepth + 1));
            }
        }

        return result;
    }

    // =========================================================================
    // Candidate filtering
    // =========================================================================

    /// <summary>
    /// Returns <c>true</c> when the element matches all locator properties that are set.
    /// Properties are checked in pywinauto search_order: controlType, className, processId,
    /// visible/offscreen, enabled, name, automationId, frameworkId, runtimeId, value, text.
    /// </summary>
    private static bool MatchesLocator(
        AutomationElement element,
        UiLocator locator,
        out string reason)
    {
        reason = string.Empty;

        // controlType
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var ct = SafeElementControlType(element);
            if (!string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase))
            {
                reason = $"controlType mismatch: '{ct}' vs '{locator.ControlType}'";
                return false;
            }
        }

        // className
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            var cn   = SafeElementClassName(element);
            var mode = locator.ClassNameMatchMode ?? locator.MatchMode;
            if (!StringMatchesByMode(cn, locator.ClassName, mode, caseSensitive: false))
            {
                reason = $"className mismatch: '{cn}'";
                return false;
            }
        }

        // processId
        if (locator.ProcessId.HasValue)
        {
            var pid = SafeProcessId(element);
            if (pid != locator.ProcessId.Value)
            {
                reason = $"processId mismatch: {pid} vs {locator.ProcessId.Value}";
                return false;
            }
        }

        // visible / offscreen
        if (locator.Visible.HasValue)
        {
            var offscreen = SafeIsOffscreen(element) ?? false;
            var visible   = !offscreen;
            if (visible != locator.Visible.Value)
            {
                reason = $"visible mismatch: {visible} vs {locator.Visible.Value}";
                return false;
            }
        }

        if (locator.Offscreen.HasValue)
        {
            var offscreen = SafeIsOffscreen(element) ?? false;
            if (offscreen != locator.Offscreen.Value)
            {
                reason = $"offscreen mismatch: {offscreen} vs {locator.Offscreen.Value}";
                return false;
            }
        }

        // enabled
        if (locator.Enabled.HasValue)
        {
            var enabled = SafeIsEnabled(element) ?? false;
            if (enabled != locator.Enabled.Value)
            {
                reason = $"enabled mismatch: {enabled} vs {locator.Enabled.Value}";
                return false;
            }
        }

        // name
        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            var name = SafeElementName(element);
            var mode = locator.NameMatchMode ?? locator.MatchMode;
            if (!StringMatchesByMode(name, locator.Name, mode, caseSensitive: true))
            {
                reason = $"name mismatch: '{name}'";
                return false;
            }
        }

        // automationId
        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            var aid  = SafeElementAutomationId(element);
            var mode = locator.AutomationIdMatchMode ?? locator.MatchMode;
            if (!StringMatchesByMode(aid, locator.AutomationId, mode, caseSensitive: true))
            {
                reason = $"automationId mismatch: '{aid}'";
                return false;
            }
        }

        // frameworkId
        if (!string.IsNullOrWhiteSpace(locator.FrameworkId))
        {
            var fwId = SafeFrameworkId(element);
            if (!string.Equals(fwId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
            {
                reason = $"frameworkId mismatch: '{fwId}'";
                return false;
            }
        }

        // runtimeId (compared as comma-joined string for simplicity)
        if (!string.IsNullOrWhiteSpace(locator.RuntimeId))
        {
            var rtId = SafeRuntimeIdString(element);
            if (!string.Equals(rtId, locator.RuntimeId, StringComparison.Ordinal))
            {
                reason = $"runtimeId mismatch: '{rtId}'";
                return false;
            }
        }

        // value (ValuePattern)
        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var val  = SafeElementValue(element);
            var mode = locator.ValueMatchMode ?? locator.MatchMode;
            if (!StringMatchesByMode(val, locator.Value, mode, caseSensitive: false))
            {
                reason = $"value mismatch: '{val}'";
                return false;
            }
        }

        // text (TextPattern or name fallback)
        if (!string.IsNullOrWhiteSpace(locator.Text))
        {
            var text = SafeElementText(element);
            var mode = locator.TextMatchMode ?? locator.MatchMode;
            if (!StringMatchesByMode(text, locator.Text, mode, caseSensitive: false))
            {
                reason = $"text mismatch: '{text}'";
                return false;
            }
        }

        return true;
    }

    // =========================================================================
    // Scoring
    // =========================================================================

    private sealed record ScoredElement(AutomationElement Element, int Score, string Reason);

    private static List<ScoredElement> ScoreAndSortCandidates(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var scored = new List<ScoredElement>(candidates.Count);
        foreach (var e in candidates)
        {
            var (score, reason) = ScoreElement(e, locator);
            scored.Add(new ScoredElement(e, score, reason));
        }

        scored.Sort((a, b) => b.Score.CompareTo(a.Score));
        return scored;
    }

    private static (int score, string reason) ScoreElement(AutomationElement element, UiLocator locator)
    {
        var score   = 50; // base score
        var reasons = new List<string>(4);

        // Name exact gets highest boost
        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            var name = SafeElementName(element);
            if (string.Equals(name, locator.Name, StringComparison.Ordinal))
            {
                score += 50;
                reasons.Add("name-exact");
            }
            else if (name.Contains(locator.Name, StringComparison.OrdinalIgnoreCase))
            {
                score += 20;
                reasons.Add("name-contains");
            }
        }

        // AutomationId exact
        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            var aid = SafeElementAutomationId(element);
            if (string.Equals(aid, locator.AutomationId, StringComparison.Ordinal))
            {
                score += 40;
                reasons.Add("automationid-exact");
            }
        }

        // ControlType match
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var ct = SafeElementControlType(element);
            if (string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase))
            {
                score += 10;
                reasons.Add("controltype");
            }
        }

        // Value / text boosts
        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var val = SafeElementValue(element);
            if (string.Equals(val, locator.Value, StringComparison.OrdinalIgnoreCase))
            {
                score += 30;
                reasons.Add("value-exact");
            }
        }

        if (!string.IsNullOrWhiteSpace(locator.Text))
        {
            var text = SafeElementText(element);
            if (string.Equals(text, locator.Text, StringComparison.OrdinalIgnoreCase))
            {
                score += 25;
                reasons.Add("text-exact");
            }
        }

        return (score, string.Join(";", reasons));
    }

    // =========================================================================
    // Value / text helpers (Phase 1.8)
    // =========================================================================

    // NOTE: SafeElementValue is defined in UiService.cs (reads ValuePattern then TextPattern).
    // SafeElementText below reads TextPattern first then falls back to element Name.

    internal static string SafeElementText(AutomationElement element)
    {
        try
        {
            var tp = element.Patterns.Text.PatternOrDefault;
            if (tp != null)
            {
                var text = tp.DocumentRange.GetText(-1);
                if (!string.IsNullOrWhiteSpace(text))
                    return text.Trim();
            }
        }
        catch { /* ignore */ }

        // Fallback to name
        return SafeElementName(element);
    }

    internal static string SafeFrameworkId(AutomationElement element)
    {
        try { return element.Properties.FrameworkId.ValueOrDefault ?? string.Empty; }
        catch { return string.Empty; }
    }

    internal static string SafeRuntimeIdString(AutomationElement element)
    {
        try
        {
            var rtId = element.Properties.RuntimeId.ValueOrDefault;
            return rtId == null ? string.Empty : string.Join(",", rtId);
        }
        catch { return string.Empty; }
    }

    // =========================================================================
    // String matching helper
    // =========================================================================

    /// <summary>
    /// Compares <paramref name="value"/> against <paramref name="pattern"/> using the
    /// specified <paramref name="mode"/>.
    /// </summary>
    internal static bool StringMatchesByMode(
        string value, string pattern, string? mode, bool caseSensitive)
    {
        var comparison = caseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        return (mode?.ToLowerInvariant() ?? "exact") switch
        {
            "contains"   => value.Contains(pattern, StringComparison.OrdinalIgnoreCase),
            "startswith" => value.StartsWith(pattern, StringComparison.OrdinalIgnoreCase),
            // Wrap user-supplied regex in a non-capturing group and apply a match timeout to
            // prevent ReDoS and injection of alternation/anchors that alter semantics.
            "regex"      => SafeRegexIsMatch(value, pattern),
            _            => string.Equals(value, pattern, comparison) // "exact" or unrecognised
        };
    }

    /// <summary>
    /// Safely evaluates a user-supplied regex against <paramref name="input"/>.
    /// The pattern is wrapped in a non-capturing group so that top-level alternations
    /// in the user string cannot escape the intended scope. A 1-second match timeout
    /// prevents ReDoS from pathological patterns.
    /// Returns false on any <see cref="RegexMatchTimeoutException"/> or
    /// <see cref="ArgumentException"/> (invalid pattern).
    /// </summary>
    internal static bool SafeRegexIsMatch(string input, string pattern)
    {
        try
        {
            // Wrap in non-capturing group so user alternation stays contained.
            var safePattern = $"(?:{pattern})";
            return Regex.IsMatch(input, safePattern,
                RegexOptions.IgnoreCase,
                matchTimeout: TimeSpan.FromSeconds(1));
        }
        catch (RegexMatchTimeoutException)
        {
            return false; // treat timeout as no-match; caller can retry with simpler pattern
        }
        catch (ArgumentException)
        {
            return false; // invalid pattern — treat as no-match
        }
    }

    // =========================================================================
    // Diagnostics helpers
    // =========================================================================

    private static List<ElementCandidateDto> BuildCandidateDtos(
        List<AutomationElement> candidates,
        UiLocator locator)
    {
        var dtos = new List<ElementCandidateDto>(candidates.Count);
        for (var i = 0; i < candidates.Count; i++)
        {
            var (score, reason) = ScoreElement(candidates[i], locator);
            dtos.Add(BuildCandidateDto(candidates[i], i, score, reason));
        }
        return dtos;
    }

    private static ElementCandidateDto BuildCandidateDto(
        AutomationElement element, int index, int score, string reason)
    {
        long? hwnd = null;
        try { hwnd = element.Properties.NativeWindowHandle.ValueOrDefault.ToInt64(); } catch { }

        return new ElementCandidateDto
        {
            Index       = index,
            Name        = SafeElementName(element),
            AutomationId = SafeElementAutomationId(element),
            ClassName   = SafeElementClassName(element),
            ControlType = SafeElementControlType(element),
            FrameworkId = SafeFrameworkId(element),
            ProcessId   = SafeProcessId(element),
            Hwnd        = hwnd,
            Value       = SafeElementValue(element),
            Text        = SafeElementText(element),
            IsEnabled   = SafeIsEnabled(element),
            IsOffscreen = SafeIsOffscreen(element),
            Rectangle   = SafeBoundingRectangleObject(element),
            Score       = score,
            Reason      = reason
        };
    }

    private string BuildElementNotFoundMessage(
        ElementResolveResult result, UiOperationPolicy policy)
    {
        var timeoutDesc = policy.Timeout.TotalMilliseconds >= 1000
            ? $"{(int)policy.Timeout.TotalSeconds}s"
            : $"{(int)policy.Timeout.TotalMilliseconds}ms";

        var sb = new System.Text.StringBuilder();
        sb.Append($"Element not found within {timeoutDesc}. ");
        sb.Append($"policy={policy.PolicyName}, strategy={result.Strategy}, ");
        sb.Append($"rootStrategy={result.RootStrategy}, ");
        sb.Append($"locator={DescribeLocator(result.Locator ?? new UiLocator())}");

        if (result.CandidateCount > 0)
            sb.Append($", candidates-collected={result.CandidateCount}");

        if (result.Errors.Count > 0)
            sb.Append($", errors=[{string.Join("; ", result.Errors)}]");

        if (result.Candidates.Count > 0)
        {
            sb.Append(", top-candidates=[");
            sb.Append(string.Join(", ",
                result.Candidates.Take(5).Select(c =>
                    $"'{c.Name}'/{c.ControlType}/{c.AutomationId}")));
            sb.Append(']');
        }

        return sb.ToString();
    }

    // =========================================================================
    // Operation: findall
    // =========================================================================

    private object? FindAll(UiRequest req)
    {
        var session = RequireSession();
        var locator = req.Locator ?? new UiLocator();
        var maxMatches = req.MaxMatches ?? GetListResponseLimit(req);

        // Resolve via ElementResolver
        var resolvedList = _elementResolver.ResolveAll(req);
        var limited = resolvedList.Take(maxMatches).ToList();

        var items = limited.Select((r, i) => new
        {
            index = i,
            name = SafeElementName(r.Element!),
            automationId = SafeElementAutomationId(r.Element!),
            className = SafeElementClassName(r.Element!),
            controlType = SafeElementControlType(r.Element!),
            frameworkId = SafeFrameworkId(r.Element!),
            runtimeId = SafeRuntimeIdString(r.Element!),
            processId = SafeProcessId(r.Element!),
            hwnd = r.Element!.Properties.NativeWindowHandle.ValueOrDefault != IntPtr.Zero ? r.Element!.Properties.NativeWindowHandle.ValueOrDefault.ToInt64() : (long?)null,
            rectangle = SafeBoundingRectangleObject(r.Element!),
            value = SafeElementValue(r.Element!),
            text = SafeElementText(r.Element!),
            isEnabled = SafeIsEnabled(r.Element!),
            isOffscreen = SafeIsOffscreen(r.Element!),
            score = 100,
            reason = "Exact match"
        }).ToList();

        return new
        {
            operation    = "findall",
            count        = resolvedList.Count,
            returned     = items.Count,
            rootStrategy = resolvedList.Count > 0 ? resolvedList[0].RootStrategy : "unknown",
            locator      = DescribeLocatorAsObject(req.Locator),
            items
        };
    }

    // =========================================================================
    // Operation: selectopendropdownitem / clickopendropdownitem (Phase 1.12)
    // =========================================================================

    /// <summary>
    /// Searches the desktop for a visible popup/dropdown item matching the request value,
    /// using broad control types: ListItem, CheckBox, RadioButton, MenuItem, Text, DataItem,
    /// TreeItem, Custom, Button.
    /// </summary>
    private object? SelectOpenDropdownItem(UiRequest req)
    {
        var matchValue = req.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(matchValue))
            throw new ArgumentException("'value' is required for selectopendropdownitem.");

        var session    = RequireSession();
        var desktop    = session.Automation.GetDesktop();
        var cf         = session.Automation.ConditionFactory;
        var matchMode  = req.Locator?.MatchMode ?? req.MatchMode ?? "contains";
        var itemRegion = ParseDropdownItemRegion(req.ItemRegion);
        var timeoutMs  = req.TimeoutMs ?? 5000;
        var deadline   = DateTime.UtcNow + TimeSpan.FromMilliseconds(timeoutMs);

        // Broad control types for popup/dropdown items
        var itemControlTypes = new[]
        {
            ControlType.ListItem,
            ControlType.CheckBox,
            ControlType.RadioButton,
            ControlType.MenuItem,
            ControlType.Text,
            ControlType.DataItem,
            ControlType.TreeItem,
            ControlType.Custom,
            ControlType.Button
        };

        do
        {
            foreach (var ct in itemControlTypes)
            {
                AutomationElement[] candidates;
                try { candidates = desktop.FindAllDescendants(cf.ByControlType(ct)); }
                catch { continue; }

                foreach (var candidate in candidates)
                {
                    try
                    {
                        // Skip clearly offscreen elements
                        if (SafeIsOffscreen(candidate) == true)
                            continue;

                        var name  = SafeElementName(candidate);
                        var value = SafeElementValue(candidate);
                        var text  = SafeElementText(candidate);

                        if (MatchesAnyText(name, value, text, matchValue, matchMode))
                        {
                            _logger.LogInformation(
                                "selectopendropdownitem candidate found. name={Name}, ct={ControlType}, matchMode={MatchMode}",
                                name, SafeElementControlType(candidate), SanitizeValue(matchMode));

                            var activated = ActivateOpenDropdownItem(candidate, matchValue, itemRegion, req.SoftVerification == true);
                            if (activated)
                                return new
                                {
                                    selected    = matchValue,
                                    name,
                                    controlType = SafeElementControlType(candidate),
                                    itemRegion  = itemRegion.ToString()
                                };
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug(ex, "selectopendropdownitem: skipping unstable candidate.");
                    }
                }
            }

            if (DateTime.UtcNow < deadline)
                Thread.Sleep(100);

        } while (DateTime.UtcNow < deadline);

        throw new InvalidOperationException(
            $"Open dropdown item '{matchValue}' was not found within {timeoutMs} ms. " +
            $"matchMode={matchMode}. Ensure the dropdown is open and the item is visible.");
    }

    /// <summary>
    /// Lists visible popup/dropdown items from the desktop using broad control types.
    /// </summary>
    private object? ListOpenDropdownItems(UiRequest req)
    {
        var session = RequireSession();
        var desktop = session.Automation.GetDesktop();
        var cf      = session.Automation.ConditionFactory;
        var limit   = GetListResponseLimit(req);

        var itemControlTypes = new[]
        {
            ControlType.ListItem,
            ControlType.CheckBox,
            ControlType.RadioButton,
            ControlType.MenuItem,
            ControlType.Text,
            ControlType.DataItem,
            ControlType.TreeItem,
            ControlType.Custom,
            ControlType.Button
        };

        var items = new List<object>(limit);

        foreach (var ct in itemControlTypes)
        {
            AutomationElement[] candidates;
            try { candidates = desktop.FindAllDescendants(cf.ByControlType(ct)); }
            catch { continue; }

            foreach (var candidate in candidates)
            {
                if (items.Count >= limit) break;

                try
                {
                    if (SafeIsOffscreen(candidate) == true) continue;

                    items.Add(new
                    {
                        name         = SafeElementName(candidate),
                        controlType  = SafeElementControlType(candidate),
                        className    = SafeElementClassName(candidate),
                        value        = SafeElementValue(candidate),
                        text         = SafeElementText(candidate),
                        automationId = SafeElementAutomationId(candidate),
                        isOffscreen  = SafeIsOffscreen(candidate),
                        isEnabled    = SafeIsEnabled(candidate),
                        bounds       = SafeBoundingRectangleObject(candidate)
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "listopendropdownitems: skipping unstable element.");
                }
            }

            if (items.Count >= limit) break;
        }

        return new
        {
            operation = "listopendropdownitems",
            count     = items.Count,
            items
        };
    }

    // =========================================================================
    // Open dropdown item activation helper
    // =========================================================================

    private bool ActivateOpenDropdownItem(
        AutomationElement item,
        string itemName,
        DropdownItemClickRegion region,
        bool softVerification)
    {
        // 1. Physical click
        try
        {
            var rect  = item.BoundingRectangle;
            var point = GetDropdownItemClickPoint(rect, region);

            if (SendInstantLeftClick(point, $"SelectOpenDropdownItem {SanitizeValue(itemName)}"))
            {
                Thread.Sleep(DropdownItemPhysicalClickSettleMs);

                if (softVerification)
                    return true;

                if (VerifyHeaderDropdownItemSelection(item, itemName))
                    return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ActivateOpenDropdownItem: physical click failed.");
        }

        // 2. SelectionItem pattern
        try
        {
            if (item.Patterns.SelectionItem.IsSupported)
            {
                item.Patterns.SelectionItem.Pattern.Select();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ActivateOpenDropdownItem: SelectionItem failed.");
        }

        // 3. Toggle pattern (CheckBox items)
        try
        {
            if (item.Patterns.Toggle.IsSupported)
            {
                item.Patterns.Toggle.Pattern.Toggle();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ActivateOpenDropdownItem: Toggle failed.");
        }

        // 4. Invoke pattern
        try
        {
            if (item.Patterns.Invoke.IsSupported)
            {
                item.Patterns.Invoke.Pattern.Invoke();
                Thread.Sleep(DropdownItemFallbackDelayMs);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ActivateOpenDropdownItem: Invoke failed.");
        }

        // 5. Space / Enter fallback
        try
        {
            item.Focus();
            Thread.Sleep(MenuFocusDelayMs);
            Keyboard.Press(VirtualKeyShort.SPACE);
            Thread.Sleep(DropdownItemFallbackDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ActivateOpenDropdownItem: Focus+Space failed.");
        }

        return false;
    }

    private static bool MatchesAnyText(
        string name, string value, string text, string pattern, string matchMode)
    {
        return StringMatchesByMode(name,  pattern, matchMode, caseSensitive: false) ||
               StringMatchesByMode(value, pattern, matchMode, caseSensitive: false) ||
               StringMatchesByMode(text,  pattern, matchMode, caseSensitive: false);
    }

    // =========================================================================
    // Updated ResolveForStateQuery — delegates to central resolver
    // =========================================================================

    /// <summary>
    /// Resolves the target element for state-query operations (isenabled, isvisible,
    /// exists, wait, etc.). Single-shot (no retry). Throws when the element cannot be found.
    /// Now delegates to <see cref="ResolveElement"/> so that new locator fields work in
    /// state-query operations as well.
    /// </summary>
    private ResolvedElement ResolveForStateQueryViaEngine(UiRequest request)
    {
        var result = ResolveElement(
            request,
            request.Locator,
            purpose:            "state-query",
            allowDesktopSearch: false,
            allowOffscreen:     true);

        if (result.Element != null)
            return new ResolvedElement
            {
                Element  = result.Element,
                Strategy = result.Strategy
            };

        throw new InvalidOperationException(
            $"Element not found for state query: strategy={result.Strategy}, " +
            $"rootStrategy={result.RootStrategy}, " +
            $"locator={DescribeLocator(request.Locator ?? new UiLocator())}" +
            (result.Errors.Count > 0 ? $", errors=[{string.Join("; ", result.Errors)}]" : ""));
    }
}
