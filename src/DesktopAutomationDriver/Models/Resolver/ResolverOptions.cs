namespace DesktopAutomationDriver.Models.Resolver;

public sealed class ResolverOptions
{
    public string SearchRoot { get; set; } = "currentWindow";
    public string TreeView { get; set; } = "control";
    public string Backend { get; set; } = "uia";
    public bool ReturnCandidates { get; set; }
    public bool IncludeDiagnostics { get; set; }
    public string Ambiguity { get; set; } = "error";
    public int MaxMatches { get; set; } = 50;
}
