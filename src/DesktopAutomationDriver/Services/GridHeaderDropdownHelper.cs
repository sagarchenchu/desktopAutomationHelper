using System.Drawing;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;

namespace DesktopAutomationDriver.Services;

public enum HeaderDropdownRegion
{
    UpperLeft,
    UpperRight,
    LowerLeft,
    LowerRight,
    CenterLeft,
    CenterRight,
    Center,
    ProbeAll
}

internal static class GridHeaderDropdownHelper
{
    public const int DropdownOpenDelayMs = 300;
    public const int DropdownRetryDelayMs = 250;
    public const int HeaderListVerticalTolerancePx = 10;
    public const int HeaderListHorizontalTolerancePx = 50;

    public static bool IsGridHeaderElement(AutomationElement element)
    {
        try
        {
            if (element.ControlType == ControlType.Header ||
                element.ControlType == ControlType.HeaderItem)
            {
                return true;
            }

            var name = SafeElementName(element);
            var className = SafeElementClassName(element);

            return name.Contains("Header", StringComparison.OrdinalIgnoreCase) ||
                   className.Contains("Header", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    public static HeaderDropdownRegion ParseRegion(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return HeaderDropdownRegion.LowerRight;

        var normalized = value
            .Trim()
            .Replace("-", string.Empty)
            .Replace("_", string.Empty)
            .Replace(" ", string.Empty)
            .ToLowerInvariant();

        return normalized switch
        {
            "upperleft" or "topleft" => HeaderDropdownRegion.UpperLeft,
            "upperright" or "topright" => HeaderDropdownRegion.UpperRight,
            "lowerleft" or "bottomleft" => HeaderDropdownRegion.LowerLeft,
            "lowerright" or "bottomright" => HeaderDropdownRegion.LowerRight,
            "centerleft" or "middleleft" => HeaderDropdownRegion.CenterLeft,
            "centerright" or "middleright" => HeaderDropdownRegion.CenterRight,
            "center" or "middle" => HeaderDropdownRegion.Center,
            "probeall" or "all" or "auto" => HeaderDropdownRegion.ProbeAll,
            _ => HeaderDropdownRegion.LowerRight
        };
    }

    public static Point GetClickPoint(RectangleF rect, HeaderDropdownRegion region)
    {
        if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            throw new InvalidOperationException("Header has invalid bounding rectangle.");

        var padX = Math.Max(4, Math.Min(12, rect.Width / 8));
        var padY = Math.Max(3, Math.Min(8, rect.Height / 4));

        return region switch
        {
            HeaderDropdownRegion.UpperLeft => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Top + padY)),

            HeaderDropdownRegion.UpperRight => new Point(
                (int)Math.Round(rect.Right - padX),
                (int)Math.Round(rect.Top + padY)),

            HeaderDropdownRegion.LowerLeft => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Bottom - padY)),

            HeaderDropdownRegion.LowerRight => new Point(
                (int)Math.Round(rect.Right - padX),
                (int)Math.Round(rect.Bottom - padY)),

            HeaderDropdownRegion.CenterLeft => new Point(
                (int)Math.Round(rect.Left + padX),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            HeaderDropdownRegion.CenterRight => new Point(
                (int)Math.Round(rect.Right - padX),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            HeaderDropdownRegion.Center => new Point(
                (int)Math.Round(rect.Left + rect.Width / 2),
                (int)Math.Round(rect.Top + rect.Height / 2)),

            _ => new Point(
                (int)Math.Round(rect.Right - padX),
                (int)Math.Round(rect.Bottom - padY))
        };
    }

    public static IReadOnlyList<(HeaderDropdownRegion Region, Point Point)> GetCandidatePoints(
        RectangleF rect,
        HeaderDropdownRegion region)
    {
        if (region != HeaderDropdownRegion.ProbeAll)
        {
            return new[]
            {
                (region, GetClickPoint(rect, region))
            };
        }

        var order = new[]
        {
            HeaderDropdownRegion.LowerRight,
            HeaderDropdownRegion.UpperRight,
            HeaderDropdownRegion.CenterRight,
            HeaderDropdownRegion.LowerLeft,
            HeaderDropdownRegion.UpperLeft,
            HeaderDropdownRegion.CenterLeft,
            HeaderDropdownRegion.Center
        };

        return order
            .Select(r => (r, GetClickPoint(rect, r)))
            .ToList();
    }

    public static IReadOnlyList<Point> GetDropdownClickPoints(RectangleF rect)
    {
        return GetCandidatePoints(rect, HeaderDropdownRegion.ProbeAll)
            .Select(x => x.Point)
            .ToArray();
    }

    public static bool IsListNearHeader(RectangleF listRect, RectangleF headerRect)
    {
        return listRect.Top >= headerRect.Bottom - HeaderListVerticalTolerancePx &&
               listRect.Left <= headerRect.Right + HeaderListHorizontalTolerancePx &&
               listRect.Right >= headerRect.Left - HeaderListHorizontalTolerancePx;
    }

    private static string SafeElementName(AutomationElement element)
    {
        try { return element.Name ?? string.Empty; }
        catch { return string.Empty; }
    }

    private static string SafeElementClassName(AutomationElement element)
    {
        try { return element.ClassName ?? string.Empty; }
        catch { return string.Empty; }
    }
}
