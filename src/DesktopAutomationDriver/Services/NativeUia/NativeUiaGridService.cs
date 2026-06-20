using System.Diagnostics;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Native UIA grid/table read and row selection. Bounded by request timeoutMs.
/// </summary>
internal sealed class NativeUiaGridService : INativeUiaGridService
{
    private const int DataItemControlTypeId = 50029;
    private const int HeaderControlTypeId = 50034;
    private const int CustomControlTypeId = 50025;
    private const int DefaultTimeoutMs = 5000;
    private const int MaxTimeoutMs = 15000;
    private const int ActionDelayMs = 80;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaElementResolver _resolver;
    private readonly NativeUiaTextReader _textReader;
    private readonly ILogger<NativeUiaGridService> _logger;

    public NativeUiaGridService(ILogger<NativeUiaGridService> logger)
        : this(new NativeUiaAutomation(), logger)
    {
    }

    internal NativeUiaGridService(NativeUiaAutomation uia, ILogger<NativeUiaGridService> logger)
    {
        _uia = uia;
        _textReader = new NativeUiaTextReader(_uia);
        _resolver = new NativeUiaElementResolver(_uia);
        _logger = logger;
    }

    public object GetGrid(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        const string operation = "getgriduia";
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            if (!HasLocator(request))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Locator is required for getgriduia.",
                    sw.ElapsedMilliseconds);
            }

            if (string.IsNullOrWhiteSpace(request.Operation))
                request.Operation = operation;

            var resolveResult = _resolver.ResolveElement(
                request,
                activeWindowHwnd,
                processId,
                deadlineUtc,
                cancellationToken);

            if (resolveResult.Element == null)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = resolveResult.Stage ?? "element-not-found",
                    stage = resolveResult.Stage,
                    elapsedMs = sw.ElapsedMilliseconds,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                };
            }

            if (resolveResult.IsAmbiguous)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = "ambiguous-element",
                    stage = resolveResult.Stage,
                    elapsedMs = sw.ElapsedMilliseconds,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                };
            }

            var (headers, rows, strategy) = ReadGridData(resolveResult.Element, request);
            if (headers.Count == 0 && rows.Count == 0)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = "grid-empty",
                    stage = resolveResult.Stage,
                    strategy,
                    elapsedMs = sw.ElapsedMilliseconds,
                    message = "No grid headers or rows could be read from the resolved element."
                };
            }

            return new
            {
                operation,
                success = true,
                reason = (string?)null,
                strategy,
                stage = resolveResult.Stage,
                headers,
                rows,
                rowCount = rows.Count,
                columnCount = headers.Count > 0 ? headers.Count : rows.FirstOrDefault()?.Count ?? 0,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
        catch (TimeoutException ex)
        {
            return BuildFailure(operation, "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure(operation, "cancelled", $"{operation} was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaGrid getgriduia failed with exception.");

            return new
            {
                operation,
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    public object SelectGridRow(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        const string operation = "selectgridrowuia";
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            if (!HasLocator(request))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Locator is required for selectgridrowuia.",
                    sw.ElapsedMilliseconds);
            }

            if (request.Index == null && string.IsNullOrWhiteSpace(request.Value))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Either index or value is required for selectgridrowuia.",
                    sw.ElapsedMilliseconds);
            }

            if (string.IsNullOrWhiteSpace(request.Operation))
                request.Operation = operation;

            var resolveResult = _resolver.ResolveElement(
                request,
                activeWindowHwnd,
                processId,
                deadlineUtc,
                cancellationToken);

            if (resolveResult.Element == null)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = resolveResult.Stage ?? "element-not-found",
                    stage = resolveResult.Stage,
                    elapsedMs = sw.ElapsedMilliseconds,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                };
            }

            if (resolveResult.IsAmbiguous)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = "ambiguous-element",
                    stage = resolveResult.Stage,
                    elapsedMs = sw.ElapsedMilliseconds,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                };
            }

            var (success, strategy, reason, rowIndex, attemptedStrategies) =
                SelectRow(resolveResult.Element, request, deadlineUtc, cancellationToken);

            if (success)
            {
                return new
                {
                    operation,
                    success = true,
                    strategy,
                    rowIndex,
                    stage = resolveResult.Stage,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            return new
            {
                operation,
                success = false,
                reason = reason ?? "action-failed",
                attemptedStrategies,
                rowIndex,
                stage = resolveResult.Stage,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
        catch (TimeoutException ex)
        {
            return BuildFailure(operation, "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure(operation, "cancelled", $"{operation} was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaGrid selectgridrowuia failed with exception.");

            return new
            {
                operation,
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private (List<string> headers, List<List<string>> rows, string strategy) ReadGridData(
        IUIAutomationElement element,
        UiRequest request)
    {
        var headers = new List<string>();
        var rows = new List<List<string>>();

        if (_uia.TryGetTablePattern(element, out var tablePattern))
        {
            try
            {
                var headersArray = tablePattern!.GetCurrentColumnHeaders();
                if (headersArray != null)
                {
                    for (var i = 0; i < headersArray.Length; i++)
                    {
                        var header = headersArray.GetElement(i);
                        headers.Add(header != null
                            ? _uia.GetStringProperty(header, UIA_PropertyIds.UIA_NamePropertyId)
                            : string.Empty);
                    }
                }
            }
            catch
            {
                // fall through to other strategies
            }
        }

        if (_uia.TryGetGridPattern(element, out var gridPattern))
        {
            try
            {
                var rowCount = gridPattern!.CurrentRowCount;
                var colCount = gridPattern.CurrentColumnCount;

                for (var row = 0; row < rowCount; row++)
                {
                    var rowValues = new List<string>();
                    for (var col = 0; col < colCount; col++)
                    {
                        try
                        {
                            var cell = gridPattern.GetItem(row, col);
                            rowValues.Add(cell != null ? ReadCellText(cell) : string.Empty);
                        }
                        catch
                        {
                            rowValues.Add(string.Empty);
                        }
                    }

                    rows.Add(rowValues);
                }

                if (headers.Count == 0 && colCount > 0)
                {
                    for (var col = 0; col < colCount; col++)
                        headers.Add($"Column{col + 1}");
                }

                if (rows.Count > 0 || headers.Count > 0)
                    return (headers, rows, "grid-pattern");
            }
            catch
            {
                // fall through
            }
        }

        if (headers.Count == 0)
            HeadersFromDescendants(element, headers);

        if (rows.Count == 0)
            DataItemsFromDescendants(element, rows);

        if (headers.Count > 0 || rows.Count > 0)
            return (headers, rows, "dataitem-descendants");

        CustomControlDeepDive(element, headers, rows);
        return (headers, rows, "custom-control-deep-dive");
    }

    private void HeadersFromDescendants(IUIAutomationElement element, List<string> headers)
    {
        var header = _uia.FindFirstDescendant(
            element,
            _uia.ControlTypeCondition(HeaderControlTypeId));

        if (header == null)
            return;

        foreach (var child in _uia.GetChildren(header, maxChildren: 200))
            headers.Add(_uia.GetStringProperty(child, UIA_PropertyIds.UIA_NamePropertyId));
    }

    private void DataItemsFromDescendants(IUIAutomationElement element, List<List<string>> rows)
    {
        var dataItems = _uia.FindAllDescendants(
            element,
            _uia.ControlTypeCondition(DataItemControlTypeId),
            limit: 500);

        foreach (var dataItem in dataItems)
        {
            var cells = _uia.GetChildren(dataItem, maxChildren: 100)
                .Select(ReadCellText)
                .ToList();

            if (cells.Count > 0)
                rows.Add(cells);
        }
    }

    private void CustomControlDeepDive(
        IUIAutomationElement root,
        List<string> headers,
        List<List<string>> rows)
    {
        var customElements = _uia.FindAllDescendants(
            root,
            _uia.ControlTypeCondition(CustomControlTypeId),
            limit: 100);

        foreach (var custom in customElements)
        {
            if (headers.Count == 0)
                HeadersFromDescendants(custom, headers);

            if (rows.Count == 0)
                DataItemsFromDescendants(custom, rows);

            if (headers.Count > 0 && rows.Count > 0)
                break;
        }
    }

    private (bool success, string? strategy, string? reason, int? rowIndex, List<string>? attemptedStrategies) SelectRow(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var attempted = new List<string>();
        var rowIndex = request.Index;

        if (_uia.TryGetGridPattern(element, out var gridPattern))
        {
            attempted.Add("grid-pattern");
            try
            {
                var rowCount = gridPattern!.CurrentRowCount;
                var resolvedRow = ResolveRowIndex(element, request, rowCount, out var matchStrategy);
                if (resolvedRow == null)
                    return (false, null, "row-not-found", rowIndex, attempted);

                rowIndex = resolvedRow;
                if (rowCount > 0 && rowIndex >= rowCount)
                    return (false, null, "row-out-of-range", rowIndex, attempted);

                var cell = gridPattern.GetItem(rowIndex.Value, 0);
                if (cell != null && ActOnGridTarget(cell, deadlineUtc, cancellationToken, out var cellStrategy))
                    return (true, $"grid-pattern:{cellStrategy}", null, rowIndex, attempted);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "GridPattern row selection failed.");
            }
        }

        attempted.Add("dataitem-row");
        var dataItems = _uia.FindAllChildren(
            element,
            _uia.ControlTypeCondition(DataItemControlTypeId),
            limit: 500);

        if (dataItems.Count == 0
            && _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId) == DataItemControlTypeId)
        {
            try
            {
                var parent = _uia.Automation.RawViewWalker.GetParentElement(element);
                if (parent != null)
                {
                    dataItems = _uia.FindAllChildren(
                        parent,
                        _uia.ControlTypeCondition(DataItemControlTypeId),
                        limit: 500);
                }
            }
            catch
            {
                // fall through
            }
        }

        if (dataItems.Count > 0)
        {
            var resolvedRow = ResolveRowIndexFromDataItems(dataItems, request, out _);
            if (resolvedRow == null)
                return (false, null, "row-not-found", rowIndex, attempted);

            rowIndex = resolvedRow;
            if (ActOnGridTarget(dataItems[rowIndex.Value], deadlineUtc, cancellationToken, out var rowStrategy))
                return (true, $"dataitem-row:{rowStrategy}", null, rowIndex, attempted);
        }

        return (false, null, "grid-row-unavailable", rowIndex, attempted);
    }

    private int? ResolveRowIndex(
        IUIAutomationElement grid,
        UiRequest request,
        int rowCount,
        out string matchStrategy)
    {
        matchStrategy = "row-index";
        if (request.Index is >= 0)
            return request.Index;

        matchStrategy = "row-name";
        if (_uia.TryGetGridPattern(grid, out var gridPattern))
        {
            var colCount = gridPattern!.CurrentColumnCount;
            var matchMode = request.Locator?.MatchMode ?? "contains";

            for (var row = 0; row < rowCount; row++)
            {
                for (var col = 0; col < Math.Max(colCount, 1); col++)
                {
                    try
                    {
                        var cell = gridPattern.GetItem(row, col);
                        if (cell == null)
                            continue;

                        if (NativeUiaText.Matches(ReadCellText(cell), request.Value, matchMode))
                            return row;
                    }
                    catch
                    {
                        // ignore stale cells
                    }
                }
            }
        }

        return null;
    }

    private static int? ResolveRowIndexFromDataItems(
        List<IUIAutomationElement> dataItems,
        UiRequest request,
        out string matchStrategy)
    {
        matchStrategy = "row-index";
        if (request.Index is >= 0)
            return request.Index.Value < dataItems.Count ? request.Index : null;

        matchStrategy = "row-name";
        var matchMode = request.Locator?.MatchMode ?? "contains";
        for (var i = 0; i < dataItems.Count; i++)
        {
            var name = dataItems[i].GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();
            if (NativeUiaText.Matches(name, request.Value, matchMode))
                return i;
        }

        return null;
    }

    private bool ActOnGridTarget(
        IUIAutomationElement target,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        out string strategy)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "selectgridrow");

        if (_uia.TryGetSelectionItemPattern(target, out var selectionItem))
        {
            try
            {
                selectionItem!.Select();
                Thread.Sleep(ActionDelayMs);
                strategy = "selectionitem-select";
                return true;
            }
            catch
            {
                // fall through
            }
        }

        if (_uia.TryGetInvokePattern(target, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                Thread.Sleep(ActionDelayMs);
                strategy = "invoke-pattern";
                return true;
            }
            catch
            {
                // fall through
            }
        }

        var rect = _uia.GetBoundingRectangle(target);
        if (rect.HasValue)
        {
            var center = new System.Drawing.Point(
                rect.Value.Left + rect.Value.Width / 2,
                rect.Value.Top + rect.Value.Height / 2);

            if (NativeUiaInput.ClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                strategy = "physical-click";
                return true;
            }
        }

        strategy = string.Empty;
        return false;
    }

    private string ReadCellText(IUIAutomationElement cell)
    {
        var value = _uia.GetValuePatternText(cell);
        if (!string.IsNullOrWhiteSpace(value))
            return value;

        var text = _textReader.ReadElementText(cell);
        if (!string.IsNullOrWhiteSpace(text))
            return text;

        return _uia.GetStringProperty(cell, UIA_PropertyIds.UIA_NamePropertyId);
    }

    private static bool HasLocator(UiRequest request) =>
        request.Hwnd is > 0
        || request.Locator?.Hwnd is > 0
        || request.Locator?.Handle is > 0
        || !string.IsNullOrWhiteSpace(request.Locator?.AutomationId)
        || !string.IsNullOrWhiteSpace(request.Locator?.Name)
        || !string.IsNullOrWhiteSpace(request.Locator?.ClassName)
        || !string.IsNullOrWhiteSpace(request.Locator?.ControlType);

    private static int ResolveTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }

    private static void EnsureWithinDeadline(DateTime deadlineUtc, CancellationToken cancellationToken, string stage)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow > deadlineUtc)
            throw new TimeoutException($"{stage} exceeded timeout.");
    }

    private static object BuildFailure(string operation, string reason, string message, long elapsedMs) => new
    {
        operation,
        success = false,
        reason,
        elapsedMs,
        message
    };
}
