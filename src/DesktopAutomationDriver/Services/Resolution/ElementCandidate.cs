using System;
using System.Collections.Generic;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementCandidate
{
    public required AutomationElement Element { get; init; }
    public required ElementSnapshot Snapshot { get; init; }

    public int Score { get; set; }
    public List<string> MatchReasons { get; } = new();
    public List<string> RejectReasons { get; } = new();

    public bool IsAccepted => RejectReasons.Count == 0;
}
