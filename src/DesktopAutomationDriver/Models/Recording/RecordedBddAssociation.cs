namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Optional BDD association attached to a recorded Assistive action.
/// Omitted from JSON when no BDD is active for the action.
/// </summary>
public sealed class RecordedBddAssociation
{
    /// <summary>Deterministic session group id such as <c>bdd-0001</c>.</summary>
    public string GroupId { get; set; } = string.Empty;

    /// <summary>Exact user-entered statement (trimmed; not interpreted).</summary>
    public string Statement { get; set; } = string.Empty;
}
