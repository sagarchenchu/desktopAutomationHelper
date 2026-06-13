using System.Diagnostics;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models.Request;
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
        if (locator == null)
            return new ElementResolveResult
            {
                Strategy = "null-locator",
                Errors   = ["Locator is null."]
            };

        var session = RequireSession();

        // ── 1. Determine root ───────────────────────────────────────────────
        string rootStrategy;
        AutomationElement root;

        if (explicitRoot != null)
        {
            root         = explicitRoot;
            rootStrategy = "explicit-root";
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
        else if (request.UseDesktopRoot == true || allowDesktopSearch)
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

        // ── 2. Fast path for simple locators (no new-style fields) ─────────
        if (!NeedsNewStyleSearch(locator))
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
                Errors       = [$"Element not found. strategy={result.Strategy}, locator={DescribeLocator(locator)}"]
            };
        }

        // ── 3. New-style: gather candidates ────────────────────────────────
        var topLevelOnly = locator.TopLevelOnly == true;
        var depth        = locator.Depth;

        List<AutomationElement> candidates;
        try
        {
            candidates = CollectCandidates(root, locator, session, depth, topLevelOnly);
        }
        catch (Exception ex)
        {
            return new ElementResolveResult
            {
                RootStrategy = rootStrategy,
                Strategy     = "collect-exception",
                Locator      = locator,
                Errors       = [$"Candidate collection failed: {ex.Message}"]
            };
        }

        // CtrlIndex: raw positional pick before filtering
        if (locator.CtrlIndex.HasValue)
        {
            var ctrlIdx = locator.CtrlIndex.Value;
            if (ctrlIdx >= 0 && ctrlIdx < candidates.Count)
                return new ElementResolveResult
                {
                    Element        = candidates[ctrlIdx],
                    Strategy       = "ctrl-index",
                    RootStrategy   = rootStrategy,
                    Locator        = locator,
                    CandidateCount = candidates.Count
                };

            return new ElementResolveResult
            {
                RootStrategy   = rootStrategy,
                Strategy       = "ctrl-index-out-of-range",
                Locator        = locator,
                CandidateCount = candidates.Count,
                Candidates     = BuildCandidateDtos(candidates.Take(10).ToList(), locator),
                Errors         = [$"ctrlIndex {ctrlIdx} out of range (collected {candidates.Count})"]
            };
        }

        // ── 4. Filter candidates ────────────────────────────────────────────
        var filtered = new List<AutomationElement>();
        foreach (var candidate in candidates)
        {
            if (MatchesLocator(candidate, locator, out _))
                filtered.Add(candidate);
        }

        if (filtered.Count == 0)
            return new ElementResolveResult
            {
                RootStrategy   = rootStrategy,
                Strategy       = "filtered-not-found",
                Locator        = locator,
                CandidateCount = candidates.Count,
                Candidates     = BuildCandidateDtos(candidates.Take(10).ToList(), locator),
                Errors         = [$"No candidates matched filters. collected={candidates.Count}, locator={DescribeLocator(locator)}"]
            };

        // ── 5. Score candidates ─────────────────────────────────────────────
        var scored = ScoreAndSortCandidates(filtered, locator);

        // ── 6. FoundIndex: pick by position after scoring ────────────────────
        if (locator.FoundIndex.HasValue)
        {
            var foundIdx = locator.FoundIndex.Value;
            if (foundIdx >= 0 && foundIdx < scored.Count)
                return new ElementResolveResult
                {
                    Element        = scored[foundIdx].Element,
                    Strategy       = "found-index",
                    RootStrategy   = rootStrategy,
                    Locator        = locator,
                    CandidateCount = scored.Count,
                    Ambiguous      = scored.Count > 1,
                    Candidates     = returnAllCandidates
                        ? scored.Select((s, i) => BuildCandidateDto(s.Element, i, s.Score, s.Reason)).ToList()
                        : []
                };

            return new ElementResolveResult
            {
                RootStrategy   = rootStrategy,
                Strategy       = "found-index-out-of-range",
                Locator        = locator,
                CandidateCount = scored.Count,
                Candidates     = scored.Select((s, i) => BuildCandidateDto(s.Element, i, s.Score, s.Reason)).ToList(),
                Errors         = [$"foundIndex {foundIdx} out of range (found {scored.Count} matching candidates)"]
            };
        }

        // ── 7. Return best match ─────────────────────────────────────────────
        return new ElementResolveResult
        {
            Element        = scored[0].Element,
            Strategy       = scored.Count > 1 ? "scored-ambiguous" : "scored-unique",
            RootStrategy   = rootStrategy,
            Locator        = locator,
            CandidateCount = filtered.Count,
            Ambiguous      = scored.Count > 1,
            Candidates     = returnAllCandidates
                ? scored.Select((s, i) => BuildCandidateDto(s.Element, i, s.Score, s.Reason)).ToList()
                : []
        };
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
        if (locator == null)
            throw new ArgumentException("'locator' is required for this operation.");

        var session  = RequireSession();
        var policy   = GetOperationPolicy(request);
        var deadline = DateTime.UtcNow + policy.Timeout;

        var effectiveAllowDesktop = allowDesktopSearch || policy.AllowDesktopPopupScan;

        // Build cache key for fast operations (only when new-style fields are absent).
        var preferAttributes = ShouldPreferAttributeSearch(request);
        var xpathOnly        = request.XPathOnly == true;

        string? cacheKey = null;
        if (policy.UseElementCache && !policy.RefreshRootEveryRetry && !NeedsNewStyleSearch(locator))
        {
            var root = GetWindowRoot(session, allowDesktopPopupScan: policy.AllowDesktopPopupScan);
            cacheKey = BuildElementCacheKey(
                session, root, locator, request.ParentLocator,
                preferAttributes: preferAttributes,
                xpathOnly: xpathOnly,
                preferXPath: request.PreferXPath == true);
        }

        if (cacheKey != null &&
            TryGetCachedElement(cacheKey, locator, out var cached) &&
            cached != null)
        {
            _logger.LogInformation(
                "UI locator resolved (engine). operation={Operation}, strategy=cache, locator={Locator}",
                SanitizeValue(request.Operation),
                DescribeLocator(locator));
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
                    DescribeLocator(locator));

                return lastResult.Element;
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
                    DescribeLocator(locator));

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

    /// <summary>
    /// BFS traversal of the UIA tree up to <paramref name="maxDepth"/> levels.
    /// Depth 0 = direct children of <paramref name="root"/>.
    /// </summary>
    private static List<AutomationElement> FindDescendantsUpToDepth(
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

    private static string SafeElementText(AutomationElement element)
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

    private static string SafeFrameworkId(AutomationElement element)
    {
        try { return element.Properties.FrameworkId.ValueOrDefault ?? string.Empty; }
        catch { return string.Empty; }
    }

    private static string SafeRuntimeIdString(AutomationElement element)
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
    private static bool StringMatchesByMode(
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
    private static bool SafeRegexIsMatch(string input, string pattern)
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
        var session  = RequireSession();
        var locator  = req.Locator ?? new UiLocator();

        // Determine root
        string rootStrategy;
        AutomationElement root;

        if (req.UseDesktopRoot == true)
        {
            root         = session.Automation.GetDesktop();
            rootStrategy = "desktop";
        }
        else if (req.ParentLocator != null)
        {
            var windowRoot   = GetWindowRoot(session, allowDesktopPopupScan: false);
            var parentResult = TryFindElementBySmartStrategy(
                windowRoot, session, req.ParentLocator,
                preferAttributes: true, xpathOnly: false);

            root         = parentResult.Element ?? windowRoot;
            rootStrategy = parentResult.Element != null ? "parent-locator" : "app-main-window";
        }
        else
        {
            root         = GetWindowRoot(session, allowDesktopPopupScan: false);
            rootStrategy = session.ActiveWindow != null ? "active-window" : "app-main-window";
        }

        var topLevelOnly = locator.TopLevelOnly == true;
        var depth        = locator.Depth;
        var maxMatches   = req.MaxMatches ?? GetListResponseLimit(req);

        var candidates = CollectCandidates(root, locator, session, depth, topLevelOnly);

        // Filter (skip when locator is completely empty)
        List<AutomationElement> filtered;
        if (IsEmptyLocator(locator))
            filtered = candidates;
        else
        {
            filtered = new List<AutomationElement>(candidates.Count);
            foreach (var c in candidates)
                if (MatchesLocator(c, locator, out _))
                    filtered.Add(c);
        }

        var scored = ScoreAndSortCandidates(filtered, locator);
        var limited = scored.Take(maxMatches).ToList();
        var items   = limited.Select((s, i) => BuildCandidateDto(s.Element, i, s.Score, s.Reason)).ToList();

        return new
        {
            operation    = "findall",
            count        = filtered.Count,
            returned     = items.Count,
            rootStrategy,
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
