using System.Drawing;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

internal sealed class NativeUiaDropdownFinder
{
    private static readonly int[] ItemControlTypeIds =
    [
        50007, 50010, 50020, 50029, 50025, 50002, 50013, 50000
    ];

    private static readonly int[] ContainerControlTypeIds =
    [
        50008, 50009, 50024, 50033, 50032, 50028, 50003
    ];

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaTextReader _textReader;

    public NativeUiaDropdownFinder(NativeUiaAutomation uia, NativeUiaTextReader textReader)
    {
        _uia = uia;
        _textReader = textReader;
    }

    public List<IUIAutomationElement> BuildSearchRoots(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var roots = new List<IUIAutomationElement> { combo };

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        roots.AddRange(_uia.FindAllChildren(_uia.Root, _uia.TrueCondition(), 200));

        if (processId.HasValue)
        {
            var sameProcess = _uia.FindAllDescendants(
                _uia.Root,
                _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                500);
            roots.AddRange(sameProcess.Where(e =>
                ContainerControlTypeIds.Contains(_uia.GetIntProperty(e, UIA_PropertyIds.UIA_ControlTypePropertyId))));
        }

        var foreground = NativeUiaInput.ForegroundWindowHandle();
        if (foreground != IntPtr.Zero)
        {
            var fgRoot = _uia.FromHandle(foreground);
            if (fgRoot != null)
                roots.Add(fgRoot);
        }

        return roots.DistinctBy(RuntimeKey).ToList();
    }

    public IUIAutomationElement? FindBestContainer(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var comboRect = _uia.GetBoundingRectangle(combo);
        IUIAutomationElement? best = null;
        var bestScore = int.MinValue;

        foreach (var root in BuildSearchRoots(combo, activeWindowHwnd, processId))
        {
            foreach (var typeId in ContainerControlTypeIds)
            {
                foreach (var container in _uia.FindAllDescendants(
                             root,
                             _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, typeId),
                             100))
                {
                    var score = ScoreContainer(container, comboRect, _uia);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = container;
                    }
                }
            }
        }

        return best;
    }

    public IEnumerable<IUIAutomationElement> EnumerateItemCandidates(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var seen = new HashSet<string>();
        foreach (var root in BuildSearchRoots(combo, activeWindowHwnd, processId))
        {
            foreach (var typeId in ItemControlTypeIds)
            {
                foreach (var element in _uia.FindAllDescendants(
                             root,
                             _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, typeId),
                             500))
                {
                    var key = RuntimeKey(element);
                    if (!seen.Add(key))
                        continue;

                    yield return element;
                }
            }
        }
    }

    public List<string> CollectVisibleItemTexts(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        int limit)
    {
        return EnumerateItemCandidates(combo, activeWindowHwnd, processId)
            .Select(_textReader.ReadElementText)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(limit)
            .ToList();
    }

    public List<object> CollectItemDiagnostics(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        int limit)
    {
        return EnumerateItemCandidates(combo, activeWindowHwnd, processId)
            .Take(limit)
            .Select((element, index) =>
            {
                var snapshot = _uia.CreateSnapshot(element);
                return NativeUiaDiagnostics.CandidateDiagnostic(index, snapshot);
            })
            .Cast<object>()
            .ToList();
    }

    private static int ScoreContainer(IUIAutomationElement container, Rectangle? comboRect, NativeUiaAutomation uia)
    {
        if (!comboRect.HasValue)
            return 0;

        var rect = uia.GetBoundingRectangle(container);
        if (!rect.HasValue)
            return 0;

        if (rect.Value.Top >= comboRect.Value.Bottom - 5)
            return 100;

        return 10;
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
