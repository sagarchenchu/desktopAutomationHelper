using System;
using System.Collections.Generic;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementResolveResult
{
    public AutomationElement? Element { get; init; }
    public IReadOnlyList<ElementCandidate> Candidates { get; init; } = Array.Empty<ElementCandidate>();
    public string Strategy { get; init; } = "";
    public bool Success => Element != null;
    public ElementDiagnostics? Diagnostics { get; init; }
}
