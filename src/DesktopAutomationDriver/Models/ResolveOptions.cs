namespace DesktopAutomationDriver.Models;

/// <summary>
/// Options for resolving an element.
/// </summary>
public class ResolveOptions
{
    public string? MatchMode { get; set; }
    public string? TreeView { get; set; }
    public string? Backend { get; set; }
    public bool? ReturnAllMatches { get; set; }
    public int? MaxMatches { get; set; }
    public bool? IncludeDiagnostics { get; set; }
    public bool? AllowBestMatch { get; set; }
    public string? BestMatch { get; set; }
    public bool? UseDesktopRoot { get; set; }
    public bool? UseActiveWindowRoot { get; set; }
    public bool? ReturnCandidates { get; set; }
    public bool? Debug { get; set; }
    public string? Ambiguity { get; set; }
    public string? Purpose { get; set; }
    public bool? Action { get; set; }
    public bool? AllowOffscreen { get; set; }
    public bool? RequireClickable { get; set; }
    public int? TimeoutMs { get; set; }
    public string? SearchRoot { get; set; }
    public int? MaxCandidates { get; set; }
}
