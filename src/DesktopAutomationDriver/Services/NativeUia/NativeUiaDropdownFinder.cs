using System.Drawing;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

internal sealed class NativeUiaDropdownFinder
{
    private static readonly int[] ItemControlTypeIds =
        [50007, 50010, 50020, 50029, 50025, 50002, 50013, 50000];

    private static readonly int[] ContainerControlTypeIds =
        [50008, 50009, 50024, 50033, 50032, 50028, 50003];

    private const int MaxItemsPerRoot = 150;
    private const int MaxTotalItems = 250;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaTextReader _textReader;

    public NativeUiaDropdownFinder(NativeUiaAutomation uia, NativeUiaTextReader textReader)
    {
        _uia = uia;
        _textReader = textReader;
    }

    public IEnumerable<IUIAutomationElement> EnumerateItemCandidates(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken,
        DateTime deadline)
    {
        var seen = new HashSet<string>();
        var count = 0;

        foreach (var root in BuildSearchRoots(combo, activeWindowHwnd, processId))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (DateTime.UtcNow >= deadline || count >= MaxTotalItems)
                yield break;

            foreach (var typeId in ItemControlTypeIds)
            {
                foreach (var element in _uia.FindAllDescendants(
                             root,
                             _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, typeId),
                             MaxItemsPerRoot))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (DateTime.UtcNow >= deadline || count >= MaxTotalItems)
                        yield break;

                    var key = RuntimeKey(element);
                    if (!seen.Add(key))
                        continue;

                    count++;
                    yield return element;
                }
            }
        }
    }

    public List<string> CollectVisibleItemTexts(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken,
        DateTime deadline,
        int limit)
    {
        return EnumerateItemCandidates(combo, activeWindowHwnd, processId, cancellationToken, deadline)
            .Select(_textReader.ReadElementText)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(limit)
            .ToList();
    }

    public List<object> BuildItemsPreview(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken,
        DateTime deadline,
        int limit)
    {
        var items = new List<object>();
        foreach (var element in EnumerateItemCandidates(combo, activeWindowHwnd, processId, cancellationToken, deadline))
        {
            var snapshot = _uia.CreateSnapshot(element);
            items.Add(new
            {
                index = items.Count,
                text = _textReader.ReadElementText(element),
                name = snapshot.Name,
                controlType = snapshot.ControlType,
                patterns = new
                {
                    selectionItem = snapshot.SupportedPatterns.Contains("SelectionItem"),
                    invoke = snapshot.SupportedPatterns.Contains("Invoke"),
                    toggle = snapshot.SupportedPatterns.Contains("Toggle")
                }
            });

            if (items.Count >= limit)
                break;
        }

        return items;
    }

    private List<IUIAutomationElement> BuildSearchRoots(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var roots = new List<IUIAutomationElement> { combo };

        var listChild = _uia.FindFirstDescendant(
            combo,
            _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50008));
        if (listChild != null)
            roots.Add(listChild);

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        var comboRect = _uia.GetBoundingRectangle(combo);
        if (processId.HasValue)
        {
            foreach (var popup in _uia.FindAllDescendants(
                         _uia.Root,
                         _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                         80))
            {
                if (IsNearCombo(popup, comboRect))
                    roots.Add(popup);
            }
        }

        var foreground = NativeUiaInput.ForegroundWindowHandle();
        if (foreground != IntPtr.Zero)
        {
            var fgRoot = _uia.FromHandle(foreground);
            if (fgRoot != null && IsNearCombo(fgRoot, comboRect))
                roots.Add(fgRoot);
        }

        return roots.DistinctBy(RuntimeKey).ToList();
    }

    private bool IsNearCombo(IUIAutomationElement element, Rectangle? comboRect)
    {
        if (!comboRect.HasValue)
            return true;

        var rect = _uia.GetBoundingRectangle(element);
        if (!rect.HasValue)
            return false;

        return rect.Value.Top >= comboRect.Value.Bottom - 30;
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
}
