namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Agent-side vision fallback policy. Disabled by default; never replaces working UIA locators.
/// </summary>
public sealed class VisionOptions
{
    /// <summary>Master switch. Default false.</summary>
    public bool Enabled { get; set; }

    /// <summary>Allow vision when semantic UIA reports the control inaccessible.</summary>
    public bool AllowWhenUiaInaccessible { get; set; } = true;

    /// <summary>Explicit policy permit (still requires Enabled=true).</summary>
    public bool AllowWhenPolicyPermits { get; set; }

    /// <summary>Max screenshot age before coordinate actions are rejected.</summary>
    public int MaxScreenshotAgeMilliseconds { get; set; } = 1500;

    /// <summary>Never persist sensitive screenshots by default.</summary>
    public bool PersistSensitiveScreenshots { get; set; }

    /// <summary>Provider: <c>fake</c> for CI.</summary>
    public string Provider { get; set; } = "fake";

    public double MinimumConfidence { get; set; } = 0.6;
}
