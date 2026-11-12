namespace DesktopAutomationAgent.Compilation;

/// <summary>
/// Immutable version pins for Phase 4D determinism / fingerprinting.
/// </summary>
public static class BddCompilerVersions
{
    /// <summary>Phase 4D compiler contract version.</summary>
    public const string CompilerVersion = "phase4d-1";

    /// <summary>
    /// Official Cucumber <c>Gherkin</c> NuGet package version (MIT license).
    /// Documented in <c>docs/phase4d-deterministic-bdd-compiler.md</c>.
    /// </summary>
    public const string ParserVersion = "gherkin-30.1.2";

    /// <summary>Bounded regex match timeout for step-registry patterns.</summary>
    public static readonly TimeSpan RegexMatchTimeout = TimeSpan.FromMilliseconds(250);
}
