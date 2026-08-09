using DesktopAutomationDriver.Models.Operations;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Provides a machine-readable catalog of /ui operations recognized by <see cref="IUiService"/>.
/// </summary>
public interface IUiOperationCatalog
{
    /// <summary>Builds the GET /ui/operations response payload.</summary>
    UiOperationCatalogResponse GetCatalog();

    /// <summary>All canonical operation descriptors in alphabetical order.</summary>
    IReadOnlyList<UiOperationDescriptor> GetOperations();

    /// <summary>
    /// Every executable name (canonical + aliases), lower-invariant, for parity checks.
    /// </summary>
    IReadOnlyCollection<string> GetAllRecognizedNames();

    /// <summary>True when the name is a known canonical operation or alias (case-insensitive).</summary>
    bool IsKnownOperation(string? operation);
}
