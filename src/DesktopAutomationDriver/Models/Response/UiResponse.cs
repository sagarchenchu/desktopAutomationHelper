namespace DesktopAutomationDriver.Models.Response;

using System.Collections;
using System.Reflection;

/// <summary>
/// Standard response envelope for the POST /ui endpoint.
/// </summary>
public class UiResponse
{
    /// <summary>true when the operation completed without error.</summary>
    public bool Success { get; set; }

    /// <summary>
    /// Result payload. Type depends on the operation:
    /// scalar operations return a single value, list operations return an array,
    /// check operations return a boolean, etc.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Human-readable error message. Only present when Success is false.
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    /// Path to an automatically captured failure screenshot.
    /// Only present when Success is false and a screenshot could be taken.
    /// </summary>
    public string? ScreenshotPath { get; set; }

    public string? Reason { get; set; }
    public object? Locator { get; set; }
    public object? Candidates { get; set; }
    public object? Suggestions { get; set; }

    public static UiResponse Fail(object value) =>
        new() { Success = false, Error = "UI resolution failed.", Value = value };

    /// <summary>Creates a successful response with an optional value payload.</summary>
    public static UiResponse Ok(object? value = null) =>
        new() { Success = true, Value = value };

    /// <summary>
    /// Builds the HTTP envelope from an operation result. When the payload exposes
    /// <c>success: false</c>, the envelope <see cref="Success"/> matches it.
    /// </summary>
    public static UiResponse FromOperationResult(object? value)
    {
        if (value != null && TryReadBoolProperty(value, "success", out var operationSuccess))
        {
            if (operationSuccess)
                return Ok(value);

            var message = TryReadStringProperty(value, "message");
            var reason = TryReadStringProperty(value, "reason");

            return new UiResponse
            {
                Success = false,
                Value = value,
                Error = message ?? reason ?? "Operation failed.",
                Reason = reason
            };
        }

        return Ok(value);
    }

    /// <summary>Creates a failed response with an error message.</summary>
    public static UiResponse Fail(string error) =>
        new() { Success = false, Error = error };

    /// <summary>Creates a failed response with an error message and a screenshot path.</summary>
    public static UiResponse Fail(string error, string? screenshotPath) =>
        new() { Success = false, Error = error, ScreenshotPath = screenshotPath };

    private static bool TryReadBoolProperty(object value, string propertyName, out bool result)
    {
        result = default;

        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key is string key
                    && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase)
                    && entry.Value is bool boolValue)
                {
                    result = boolValue;
                    return true;
                }
            }
        }

        var property = value.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if (property?.PropertyType == typeof(bool))
        {
            result = (bool)property.GetValue(value)!;
            return true;
        }

        return false;
    }

    private static string? TryReadStringProperty(object value, string propertyName)
    {
        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key is string key
                    && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase)
                    && entry.Value is string stringValue)
                {
                    return stringValue;
                }
            }
        }

        var property = value.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        return property?.PropertyType == typeof(string)
            ? property.GetValue(value) as string
            : null;
    }
}
