using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services;

public enum DatePickerFormatOrder
{
    Unknown = 0,
    MonthDayYear,
    DayMonthYear,
    YearMonthDay
}

public sealed class DatePickerFormatInfo
{
    public DatePickerFormatOrder Order { get; init; } = DatePickerFormatOrder.Unknown;
    public string DisplayFormat { get; init; } = "MM/DD/YYYY";
    public string Separator { get; init; } = "/";
    public string Source { get; init; } = "default";

    public bool IsKnown => Order != DatePickerFormatOrder.Unknown;
}

internal static class WinFormsDateTimePickerHelper
{
    private const int MinimumMonthClickOffsetPx = 8;
    private const double MonthClickWidthDivisor = 10.0;
    private const string MonthDayPartFormat = "00";
    private const string YearPartFormat = "0000";

    public const int DatePickerClickDelayMs = 100;
    public const int DatePickerSegmentDelayMs = 75;
    public const int DatePickerCommitDelayMs = 100;

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

    /// <summary>
    /// Detects the date format used by the given DateTimePicker element by inspecting
    /// its value, name, and help-text properties, then falling back to the current culture.
    /// </summary>
    public static DatePickerFormatInfo DetectDateFormat(AutomationElement element)
    {
        try
        {
            var candidates = new List<string>();

            try
            {
                if (element.Patterns.Value.IsSupported)
                {
                    var value = element.Patterns.Value.Pattern.Value;
                    if (!string.IsNullOrWhiteSpace(value))
                        candidates.Add(value);
                }
            }
            catch { /* ignore */ }

            try
            {
                var name = element.Name;
                if (!string.IsNullOrWhiteSpace(name))
                    candidates.Add(name);
            }
            catch { /* ignore */ }

            try
            {
                var helpText = element.Properties.HelpText.ValueOrDefault;
                if (!string.IsNullOrWhiteSpace(helpText))
                    candidates.Add(helpText);
            }
            catch { /* ignore */ }

            foreach (var candidate in candidates)
            {
                var detected = DetectDateFormatFromText(candidate);
                if (detected.IsKnown)
                    return detected;
            }

            // Fallback to current culture short date pattern.
            var culturePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            var fromCulture = DetectDateFormatFromPattern(culturePattern);

            if (fromCulture.IsKnown)
            {
                return new DatePickerFormatInfo
                {
                    Order = fromCulture.Order,
                    DisplayFormat = fromCulture.DisplayFormat,
                    Separator = fromCulture.Separator,
                    Source = $"culture:{culturePattern}"
                };
            }

            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.MonthDayYear,
                DisplayFormat = "MM/DD/YYYY",
                Separator = "/",
                Source = "default"
            };
        }
        catch
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.MonthDayYear,
                DisplayFormat = "MM/DD/YYYY",
                Separator = "/",
                Source = "fallback-exception"
            };
        }
    }

    /// <summary>
    /// Attempts to infer the date format order from a date-like string (e.g. "31/12/2026").
    /// Returns an unknown-order <see cref="DatePickerFormatInfo"/> when detection is ambiguous.
    /// </summary>
    private static DatePickerFormatInfo DetectDateFormatFromText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return new DatePickerFormatInfo();

        var normalized = text.Trim();

        // Match patterns like  31/12/2026  or  12/31/2026  or  2026-12-31
        var match = Regex.Match(
            normalized,
            @"^(?<a>\d{1,4})(?<sep>[/\-\.])(?<b>\d{1,2})\k<sep>(?<c>\d{1,4})$");

        if (!match.Success)
            return new DatePickerFormatInfo();

        var sep = match.Groups["sep"].Value;

        if (!int.TryParse(match.Groups["a"].Value, out var a) ||
            !int.TryParse(match.Groups["b"].Value, out var b) ||
            !int.TryParse(match.Groups["c"].Value, out var c))
        {
            return new DatePickerFormatInfo();
        }

        // YYYY-MM-DD: first component >= 100 means year-first
        if (a >= 100)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.YearMonthDay,
                DisplayFormat = $"YYYY{sep}MM{sep}DD",
                Separator = sep,
                Source = $"value-detect:{normalized}"
            };
        }

        // DD/MM/YYYY: first component > 12 can only be day
        if (a > 12 && c >= 100)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.DayMonthYear,
                DisplayFormat = $"DD{sep}MM{sep}YYYY",
                Separator = sep,
                Source = $"value-detect:{normalized}"
            };
        }

        // MM/DD/YYYY: second component > 12 can only be day
        if (b > 12 && c >= 100)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.MonthDayYear,
                DisplayFormat = $"MM{sep}DD{sep}YYYY",
                Separator = sep,
                Source = $"value-detect:{normalized}"
            };
        }

        // Ambiguous — cannot determine from value alone
        return new DatePickerFormatInfo();
    }

    /// <summary>
    /// Infers format order from a .NET date format pattern string (e.g. "dd/MM/yyyy").
    /// </summary>
    private static DatePickerFormatInfo DetectDateFormatFromPattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
            return new DatePickerFormatInfo();

        var lower = pattern.ToLowerInvariant();

        // Determine separator
        var sep = "/";
        foreach (var c in pattern)
        {
            if (c == '/' || c == '-' || c == '.')
            {
                sep = c.ToString();
                break;
            }
        }

        // Find position of year, month, day tokens
        var yIdx = lower.IndexOf('y');
        var mIdx = lower.IndexOf('m');
        var dIdx = lower.IndexOf('d');

        if (yIdx < 0 || mIdx < 0 || dIdx < 0)
            return new DatePickerFormatInfo();

        if (yIdx < mIdx && mIdx < dIdx)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.YearMonthDay,
                DisplayFormat = $"YYYY{sep}MM{sep}DD",
                Separator = sep,
                Source = $"pattern:{pattern}"
            };
        }

        if (dIdx < mIdx && mIdx < yIdx)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.DayMonthYear,
                DisplayFormat = $"DD{sep}MM{sep}YYYY",
                Separator = sep,
                Source = $"pattern:{pattern}"
            };
        }

        if (mIdx < dIdx && dIdx < yIdx)
        {
            return new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.MonthDayYear,
                DisplayFormat = $"MM{sep}DD{sep}YYYY",
                Separator = sep,
                Source = $"pattern:{pattern}"
            };
        }

        return new DatePickerFormatInfo();
    }

    /// <summary>
    /// Parses the user-supplied date string according to the given format order,
    /// returning the three typed segments (first/second/third) in their display order
    /// and the resolved <see cref="DateTime"/>.
    /// </summary>
    public static bool TryParseDateParts(
        string input,
        DatePickerFormatInfo format,
        out string first,
        out string second,
        out string third,
        out DateTime parsedDate)
    {
        first = string.Empty;
        second = string.Empty;
        third = string.Empty;
        parsedDate = default;

        if (string.IsNullOrWhiteSpace(input))
            return false;

        var sep = format.Separator;

        // Build candidate parse formats based on the detected order and separator.
        string[] formats;
        switch (format.Order)
        {
            case DatePickerFormatOrder.DayMonthYear:
                formats = [$"dd{sep}MM{sep}yyyy", $"d{sep}M{sep}yyyy",
                            $"dd-MM-yyyy", $"d-M-yyyy",
                            $"dd/MM/yyyy", $"d/M/yyyy"];
                break;
            case DatePickerFormatOrder.YearMonthDay:
                formats = [$"yyyy{sep}MM{sep}dd", $"yyyy{sep}M{sep}d",
                            $"yyyy-MM-dd", $"yyyy/MM/dd"];
                break;
            default: // MonthDayYear
                formats = ["MM/dd/yyyy", "M/d/yyyy", "MM-dd-yyyy", "M-d-yyyy"];
                break;
        }

        var cleaned = input.Trim();

        if (!DateTime.TryParseExact(cleaned, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate) &&
            !DateTime.TryParse(cleaned, CultureInfo.CurrentCulture, DateTimeStyles.None, out parsedDate) &&
            !DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
        {
            return false;
        }

        // Produce segments in the display order required by the format.
        var month = parsedDate.Month.ToString(MonthDayPartFormat, CultureInfo.InvariantCulture);
        var day   = parsedDate.Day.ToString(MonthDayPartFormat, CultureInfo.InvariantCulture);
        var year  = parsedDate.Year.ToString(YearPartFormat, CultureInfo.InvariantCulture);

        switch (format.Order)
        {
            case DatePickerFormatOrder.DayMonthYear:
                first  = day;
                second = month;
                third  = year;
                break;
            case DatePickerFormatOrder.YearMonthDay:
                first  = year;
                second = month;
                third  = day;
                break;
            default: // MonthDayYear
                first  = month;
                second = day;
                third  = year;
                break;
        }

        return true;
    }

    /// <summary>
    /// Legacy overload — parses in MM/DD/YYYY order.
    /// Retained for backward-compatibility with <c>UiService</c>.
    /// </summary>
    public static bool TryParseDateParts(
        string input,
        out string month,
        out string day,
        out string year)
    {
        var result = TryParseDateParts(
            input,
            new DatePickerFormatInfo
            {
                Order = DatePickerFormatOrder.MonthDayYear,
                DisplayFormat = "MM/DD/YYYY",
                Separator = "/",
                Source = "legacy"
            },
            out month,
            out day,
            out year,
            out _);
        return result;
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
