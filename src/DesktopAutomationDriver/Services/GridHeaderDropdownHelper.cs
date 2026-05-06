using System.Drawing;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;

namespace DesktopAutomationDriver.Services;

internal static class GridHeaderDropdownHelper
{
    public const int DropdownOpenDelayMs = 300;
    public const int DropdownRetryDelayMs = 250;
    public const int HeaderListVerticalTolerancePx = 10;
    public const int HeaderListHorizontalTolerancePx = 50;

    private static readonly int[] DropdownIconOffsetsPx = [8, 12, 16, 22, 28];

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

    public static IReadOnlyList<Point> GetDropdownClickPoints(RectangleF rect)
    {
        return DropdownIconOffsetsPx
            .Select(offset => new Point(
                (int)Math.Round(rect.Right - Math.Min(offset, Math.Max(1, rect.Width - 1))),
                (int)Math.Round(rect.Top + rect.Height / 2.0)))
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
