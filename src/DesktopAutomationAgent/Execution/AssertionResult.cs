using System.Text.Json;

namespace DesktopAutomationAgent.Execution;

public sealed class AssertionResult
{
    public required string Path { get; init; }

    public required string Operator { get; init; }

    public bool Passed { get; init; }

    public string? Message { get; init; }

    public JsonElement? Expected { get; init; }

    public JsonElement? Actual { get; init; }
}
