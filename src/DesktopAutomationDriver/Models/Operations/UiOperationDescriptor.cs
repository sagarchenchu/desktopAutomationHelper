namespace DesktopAutomationDriver.Models.Operations;

/// <summary>
/// Machine-readable description of a single /ui operation (canonical name).
/// </summary>
public sealed class UiOperationDescriptor
{
    public string Name { get; init; } = string.Empty;

    public IReadOnlyList<string> Aliases { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Subset of <see cref="Aliases"/> that remain executable for compatibility
    /// but should not be used by new clients (for example misspelled legacy names).
    /// </summary>
    public IReadOnlyList<string> DeprecatedAliases { get; init; } = Array.Empty<string>();

    public string Category { get; init; } = string.Empty;

    public string OperationType { get; init; } = string.Empty;

    public bool RequiresSession { get; init; }

    /// <summary>
    /// Inputs that are always required. When <see cref="RequiredInputAlternatives"/> is
    /// non-empty, this is typically the intersection of those alternatives (common inputs).
    /// </summary>
    public IReadOnlyList<string> RequiredInputs { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Accepted alternative input combinations. Clients must satisfy at least one
    /// complete alternative. Empty when a single fixed input set applies.
    /// </summary>
    public IReadOnlyList<IReadOnlyList<string>> RequiredInputAlternatives { get; init; } =
        Array.Empty<IReadOnlyList<string>>();

    public bool Deprecated { get; init; }
}

/// <summary>
/// Response payload for GET /ui/operations.
/// </summary>
public sealed class UiOperationCatalogResponse
{
    public int SchemaVersion { get; init; } = 1;

    public string DriverVersion { get; init; } = string.Empty;

    public IReadOnlyList<UiOperationDescriptor> Operations { get; init; } = Array.Empty<UiOperationDescriptor>();
}
