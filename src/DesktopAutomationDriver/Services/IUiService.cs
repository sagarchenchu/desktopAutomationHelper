using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Executes the full set of operations exposed by the POST /ui endpoint.
/// </summary>
public interface IUiService
{
    /// <summary>
    /// Dispatches <paramref name="request"/> to the appropriate handler and returns
    /// the operation result as a plain .NET object suitable for JSON serialisation.
    /// Throws <see cref="InvalidOperationException"/> when the operation cannot be
    /// completed (e.g. element not found, no active session) and
    /// <see cref="ArgumentException"/> for invalid request arguments.
    /// </summary>
    /// <param name="request">The unified UI request.</param>
    /// <returns>
    /// The result value – type depends on the operation.
    /// Returns null for void operations (click, type, etc.).
    /// </returns>
    object? Execute(UiRequest request);

    /// <summary>
    /// Captures a screenshot of the active application window (if a session is open)
    /// or the primary screen (if no session is active), saves it as a PNG file inside
    /// <paramref name="directory"/>, and returns the full path to the saved file.
    /// Returns <c>null</c> when the screenshot cannot be captured.
    /// </summary>
    /// <param name="directory">Directory where the screenshot file will be created.</param>
    string? TakeFailureScreenshot(string directory);
}
