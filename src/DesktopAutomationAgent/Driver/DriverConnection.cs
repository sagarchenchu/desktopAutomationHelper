namespace DesktopAutomationAgent.Driver;

public sealed class DriverConnection
{
    public required Uri BaseUri { get; init; }

    /// <summary>Raw bearer token value (never log).</summary>
    public required string BearerToken { get; init; }

    public required string DiscoveryMethod { get; init; }

    public string? DiscoveredUsername { get; init; }

    public string SafeBaseUrl => $"{BaseUri.Scheme}://{BaseUri.Authority}";
}
