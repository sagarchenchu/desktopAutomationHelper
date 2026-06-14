using System;
using System.Collections.Generic;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementDiagnostics
{
    public int CandidatesScanned { get; set; }
    public string Message { get; set; } = "";
    public string Status { get; set; } = "";
    public List<ElementCandidate> Candidates { get; init; } = new();
    public List<ElementCandidate> TopRejectedCandidates { get; init; } = new();
}
