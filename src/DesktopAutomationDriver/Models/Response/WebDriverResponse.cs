namespace DesktopAutomationDriver.Models.Response;

/// <summary>
/// Standard WebDriver-compatible response envelope.
/// </summary>
/// <typeparam name="T">Type of the value payload.</typeparam>
public class WebDriverResponse<T>
{
    /// <summary>
    /// The session ID this response belongs to.
    /// </summary>
    public string? SessionId { get; set; }

    /// <summary>
    /// Status code: 0 = success, non-zero = error.
    /// </summary>
    public int Status { get; set; }

    /// <summary>
    /// The response payload.
    /// </summary>
    public T? Value { get; set; }

    public static WebDriverResponse<T> Success(T value, string? sessionId = null) =>
        new() { SessionId = sessionId, Status = 0, Value = value };

    public static WebDriverResponse<ErrorDetail> Error(int status, string message, string error = "unknown error") =>
        new WebDriverResponse<ErrorDetail>
        {
            Status = status,
            Value = new ErrorDetail { Error = error, Message = message }
        };
}

/// <summary>
/// Error detail payload returned when an operation fails.
/// </summary>
public class ErrorDetail
{
    public string Error { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? Stacktrace { get; set; }
}
