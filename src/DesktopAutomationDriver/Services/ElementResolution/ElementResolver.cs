using System.Diagnostics;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.Resolution;
using FlaUI.Core.AutomationElements;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services.ElementResolution;

/// <summary>
/// Central pywinauto-style element identification engine.
/// Delegates collection/matching/scoring to <see cref="Resolution.ElementResolver"/>
/// and returns rich diagnostics for all operations.
/// </summary>
public sealed class ElementResolver
{
    private readonly Resolution.ElementResolver _inner;
    private readonly ILogger _logger;

    public ElementResolver(
        IUiSessionContext ctx,
        ILogger logger,
        Func<AutomationSession, bool, AutomationElement>? getWindowRoot = null)
    {
        _inner = getWindowRoot == null
            ? new Resolution.ElementResolver(ctx, logger)
            : new Resolution.ElementResolver(ctx, logger, getWindowRoot);
        _logger = logger;
    }

    public ElementResolveResult ResolveOne(
        UiRequest request,
        UiLocator? locator,
        ResolveOptions options,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        locator ??= request.Locator;

        if (locator == null || UiService.IsEmptyLocator(locator))
        {
            return Fail("Locator is empty", locator, null, options, sw.ElapsedMilliseconds);
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            var prepared = PrepareRequest(request, locator, options);
            var innerResult = _inner.ResolveOne(locator, prepared, MapPurpose(options));
            return MapSuccess(innerResult, locator, prepared, options, sw.ElapsedMilliseconds);
        }
        catch (ElementResolutionException ex)
        {
            return MapFailure(ex, locator, request, options, sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ElementResolver.ResolveOne failed.");
            return Fail(ex.Message, locator, BuildSearchRootSnapshot(request), options, sw.ElapsedMilliseconds);
        }
    }

    public IReadOnlyList<ElementResolveResult> ResolveAll(
        UiRequest request,
        UiLocator? locator,
        ResolveOptions options,
        CancellationToken cancellationToken = default)
    {
        locator ??= request.Locator;
        if (locator == null || UiService.IsEmptyLocator(locator))
            return Array.Empty<ElementResolveResult>();

        cancellationToken.ThrowIfCancellationRequested();
        var prepared = PrepareRequest(request, locator, options);
        var candidates = _inner.ResolveAll(locator, prepared, MapPurpose(options));

        return candidates
            .Select((c, index) => new ElementResolveResult
            {
                Success = true,
                Element = c.Element,
                Snapshot = ElementSnapshot.FromResolutionSnapshot(c.Snapshot),
                Strategy = "pywinauto-style-resolver-all",
                CandidateCount = candidates.Count,
                Criteria = ElementSearchCriteria.FromLocator(locator),
                SearchRoot = BuildSearchRootSnapshot(prepared),
                ElapsedMs = 0,
                RawCandidates = new[] { ElementCandidate.FromResolutionCandidate(c, index) }
            })
            .ToList();
    }

    public ElementResolveResult ResolveWithRetry(
        UiRequest request,
        UiLocator? locator,
        ResolveOptions options,
        TimeSpan timeout,
        TimeSpan pollInterval,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        ElementResolveResult? last = null;

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();

            last = ResolveOne(request, locator, options, cancellationToken);
            if (last.Success)
            {
                return new ElementResolveResult
                {
                    Success = last.Success,
                    Element = last.Element,
                    Snapshot = last.Snapshot,
                    Strategy = last.Strategy,
                    Error = last.Error,
                    CandidateCount = last.CandidateCount,
                    Candidates = last.Candidates,
                    Criteria = last.Criteria,
                    SearchRoot = last.SearchRoot,
                    ElapsedMs = sw.ElapsedMilliseconds,
                    Ambiguous = last.Ambiguous,
                    FallbackUsed = last.FallbackUsed,
                    RawCandidates = last.RawCandidates
                };
            }

            if (pollInterval > TimeSpan.Zero)
                cancellationToken.WaitHandle.WaitOne(pollInterval);
        }

        return last ?? Fail(
            "Element not found",
            locator ?? request.Locator,
            BuildSearchRootSnapshot(request),
            options,
            sw.ElapsedMilliseconds,
            strategy: "pywinauto-style-resolver-timeout");
    }

    private static UiRequest PrepareRequest(UiRequest request, UiLocator locator, ResolveOptions options)
    {
        var clone = new UiRequest
        {
            Operation = request.Operation,
            Locator = locator,
            ParentLocator = request.ParentLocator,
            TimeoutMs = request.TimeoutMs,
            PollIntervalMs = request.PollIntervalMs,
            SearchScope = request.SearchScope ?? locator.SearchScope,
            MatchMode = request.MatchMode ?? locator.MatchMode,
            FoundIndex = request.FoundIndex ?? locator.FoundIndex ?? locator.Index,
            CtrlIndex = request.CtrlIndex ?? locator.CtrlIndex,
            Depth = request.Depth ?? locator.Depth,
            TopLevelOnly = request.TopLevelOnly ?? locator.TopLevelOnly,
            ActiveOnly = request.ActiveOnly ?? locator.ActiveOnly,
            IncludeOffscreen = options.AllowOffscreen || options.IncludeHidden
                ? true
                : request.IncludeOffscreen ?? locator.IncludeOffscreen ?? false,
            IncludeHidden = options.IncludeHidden,
            ThrowIfAmbiguous = options.ThrowIfAmbiguous || request.ThrowIfAmbiguous == true,
            TreeView = request.TreeView,
            FallbackToWindowRootIfParentChildNotFound = request.FallbackToWindowRootIfParentChildNotFound,
            BestMatch = request.BestMatch ?? locator.BestMatch,
            DesktopSearch = request.DesktopSearch,
            SearchRoot = request.SearchRoot
        };

        if (!options.AllowDisabled)
        {
            clone.Locator = CloneLocator(locator);
            clone.Locator.Enabled = true;
        }

        return clone;
    }

    private static UiLocator CloneLocator(UiLocator locator) =>
        new()
        {
            Mode = locator.Mode,
            Name = locator.Name,
            NameRegex = locator.NameRegex,
            AutomationId = locator.AutomationId,
            AutomationIdRegex = locator.AutomationIdRegex,
            ClassName = locator.ClassName,
            ClassNameRegex = locator.ClassNameRegex,
            ControlType = locator.ControlType,
            XPath = locator.XPath,
            Hwnd = locator.Hwnd,
            ProcessId = locator.ProcessId,
            FrameworkId = locator.FrameworkId,
            RuntimeId = locator.RuntimeId,
            Value = locator.Value,
            ValueRegex = locator.ValueRegex,
            Text = locator.Text,
            MatchMode = locator.MatchMode,
            FoundIndex = locator.FoundIndex,
            CtrlIndex = locator.CtrlIndex,
            Depth = locator.Depth,
            TopLevelOnly = locator.TopLevelOnly,
            ActiveOnly = locator.ActiveOnly,
            IncludeOffscreen = locator.IncludeOffscreen,
            IncludeDisabled = locator.IncludeDisabled,
            BestMatch = locator.BestMatch,
            SearchScope = locator.SearchScope,
            Visible = locator.Visible,
            Enabled = locator.Enabled,
            Offscreen = locator.Offscreen
        };

    private static string MapPurpose(ResolveOptions options) =>
        string.IsNullOrWhiteSpace(options.Purpose) ? "query" : options.Purpose;

    private ElementResolveResult MapSuccess(
        Resolution.ElementResolveResult inner,
        UiLocator locator,
        UiRequest request,
        ResolveOptions options,
        long elapsedMs)
    {
        var rawCandidates = inner.Candidates
            .Select((c, i) => ElementCandidate.FromResolutionCandidate(c, i))
            .ToList();

        var snapshots = rawCandidates.Select(c => c.Snapshot).ToList();
        var snapshot = inner.Element == null
            ? null
            : ElementSnapshot.FromResolutionSnapshot(
                Resolution.ElementResolver.CreateSnapshot(inner.Element));

        var ambiguous = inner.Diagnostics?.Status == "Ambiguous" || inner.Candidates.Count > 1;
        var strategy = string.IsNullOrWhiteSpace(inner.Strategy)
            ? "pywinauto-style-resolver"
            : inner.Strategy;

        if (options.ThrowIfNotFound && inner.Element == null)
            throw new InvalidOperationException($"Element not found for {UiService.DescribeLocator(locator)}.");

        if (options.ThrowIfAmbiguous && ambiguous)
            throw new InvalidOperationException(
                $"Multiple elements matched for {UiService.DescribeLocator(locator)}. candidateCount={snapshots.Count}");

        return new ElementResolveResult
        {
            Success = inner.Element != null,
            Element = inner.Element,
            Snapshot = snapshot,
            Strategy = strategy,
            CandidateCount = snapshots.Count,
            Candidates = options.ReturnCandidates ? snapshots : new List<ElementSnapshot>(),
            Criteria = ElementSearchCriteria.FromLocator(locator),
            SearchRoot = BuildSearchRootSnapshot(request),
            ElapsedMs = elapsedMs,
            Ambiguous = ambiguous,
            RawCandidates = rawCandidates
        };
    }

    private ElementResolveResult MapFailure(
        ElementResolutionException ex,
        UiLocator? locator,
        UiRequest request,
        ResolveOptions options,
        long elapsedMs)
    {
        var rawCandidates = ex.TopRejectedCandidates
            .Select((c, i) => ElementCandidate.FromResolutionCandidate(c, i))
            .ToList();

        var snapshots = rawCandidates.Select(c => c.Snapshot).ToList();
        var ambiguous = ex.Message.Contains("Multiple elements matched", StringComparison.OrdinalIgnoreCase);

        if (options.ThrowIfNotFound && !ambiguous)
            throw new InvalidOperationException(ex.Message, ex);

        if (options.ThrowIfAmbiguous && ambiguous)
            throw new InvalidOperationException(ex.Message, ex);

        return new ElementResolveResult
        {
            Success = false,
            Error = ex.Message,
            Strategy = ambiguous ? "pywinauto-style-resolver-ambiguous" : "pywinauto-style-resolver-not-found",
            CandidateCount = ex.CandidatesScanned,
            Candidates = options.ReturnCandidates ? snapshots : new List<ElementSnapshot>(),
            Criteria = ElementSearchCriteria.FromLocator(locator),
            SearchRoot = BuildSearchRootSnapshot(request),
            ElapsedMs = elapsedMs,
            Ambiguous = ambiguous,
            RawCandidates = rawCandidates
        };
    }

    private object? BuildSearchRootSnapshot(UiRequest request)
    {
        try
        {
            var root = _inner.ResolveSearchRoot(request);
            var snap = Resolution.ElementResolver.CreateSnapshot(root);
            return new
            {
                name = snap.Name,
                controlType = snap.ControlType,
                automationId = snap.AutomationId,
                hwnd = snap.Hwnd,
                rectangle = snap.Rectangle
            };
        }
        catch
        {
            return new { name = request.SearchScope ?? "currentWindow" };
        }
    }

    private static ElementResolveResult Fail(
        string error,
        UiLocator? locator,
        object? searchRoot,
        ResolveOptions options,
        long elapsedMs,
        string strategy = "pywinauto-style-resolver")
    {
        return new ElementResolveResult
        {
            Success = false,
            Error = error,
            Strategy = strategy,
            Criteria = ElementSearchCriteria.FromLocator(locator),
            SearchRoot = searchRoot,
            ElapsedMs = elapsedMs,
            Candidates = new List<ElementSnapshot>()
        };
    }
}
