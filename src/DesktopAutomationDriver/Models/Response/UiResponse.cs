namespace DesktopAutomationDriver.Models.Response;

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

    /// <summary>Creates a successful response with an optional value payload.</summary>
    public static UiResponse Ok(object? value = null) =>
        new() { Success = true, Value = value };

    /// <summary>Creates a failed response with an error message.</summary>
    public static UiResponse Fail(string error) =>
        new() { Success = false, Error = error };

    /// <summary>Creates a failed response with an error message and a screenshot path.</summary>
    public static UiResponse Fail(string error, string? screenshotPath) =>
        new() { Success = false, Error = error, ScreenshotPath = screenshotPath };
}
