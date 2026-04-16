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
}
