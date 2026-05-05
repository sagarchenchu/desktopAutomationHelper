using System.Drawing;
using System.Globalization;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services;

internal static class WinFormsDateTimePickerHelper
{
    private const int MinimumMonthClickOffsetPx = 8;
    private const double MonthClickWidthDivisor = 10.0;

    public static bool IsDateTimePicker(AutomationElement element)
    {
        return IsDateTimePickerClassName(SafeElementClassName(element));
    }

    public static bool IsDateTimePickerClassName(string? className)
    {
        return !string.IsNullOrWhiteSpace(className) &&
               className.Contains("SysDateTimePick32", StringComparison.OrdinalIgnoreCase);
    }

    public static Point GetMonthSectionPoint(RectangleF rect)
    {
        // Click slightly inside the left edge so focus lands on the month segment,
        // while keeping a minimum offset for very narrow controls.
        return new Point(
            (int)Math.Round(rect.Left + Math.Max(MinimumMonthClickOffsetPx, rect.Width / MonthClickWidthDivisor)),
            (int)Math.Round(rect.Top + rect.Height / 2.0));
    }

    public static bool TryParseDateParts(
        string input,
        out string month,
        out string day,
        out string year)
    {
        month = string.Empty;
        day = string.Empty;
        year = string.Empty;

        if (string.IsNullOrWhiteSpace(input))
            return false;

        var cleaned = input.Trim();

        var formats = new[]
        {
            "MM/dd/yyyy",
            "M/d/yyyy",
            "MM-dd-yyyy",
            "M-d-yyyy"
        };

        if (!DateTime.TryParseExact(
                cleaned,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date))
        {
            if (!DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                return false;
        }

        month = date.Month.ToString("00", CultureInfo.InvariantCulture);
        day = date.Day.ToString("00", CultureInfo.InvariantCulture);
        year = date.Year.ToString("0000", CultureInfo.InvariantCulture);

        return true;
    }

    private static string SafeElementClassName(AutomationElement element)
    {
        try
        {
            return element.ClassName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }
}
