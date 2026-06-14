using System.Collections.Generic;

namespace DesktopAutomationDriver.Models.Resolver;

public sealed class ResolveDiagnostics
{
    public string SearchRoot { get; set; } = string.Empty;
    public string TreeView { get; set; } = string.Empty;
    public string Backend { get; set; } = string.Empty;
    public string Strategy { get; set; } = string.Empty;
    public int CandidateCount { get; set; }
    public List<ResolvedElement> Candidates { get; } = new();
    public List<string> Errors { get; } = new();
}
