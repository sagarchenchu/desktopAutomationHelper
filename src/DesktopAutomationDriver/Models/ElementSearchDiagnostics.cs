using System.Collections.Generic;

namespace DesktopAutomationDriver.Models;

/// <summary>
/// Diagnostics information for element search, used in case of failures or ambiguity.
/// </summary>
public class ElementSearchDiagnostics
{
    public string Status { get; set; } = string.Empty; // e.g. "ElementNotFound" or "ElementAmbiguous"
    public string Message { get; set; } = string.Empty;
    public List<ElementCandidate> Candidates { get; set; } = new();
    public List<string> Errors { get; set; } = new();
}
