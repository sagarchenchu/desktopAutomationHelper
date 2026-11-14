namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Publishes the bearer token over a per-user SID ACL-protected named pipe and
/// connection descriptor. Must never log the token.
/// </summary>
public interface ISecureBootstrapHost : IAsyncDisposable
{
    string UserSid { get; }

    string PipeName { get; }

    string DescriptorPath { get; }

    bool IsRunning { get; }

    /// <summary>Starts the named-pipe server and writes the ACL-protected descriptor.</summary>
    Task StartAsync(CancellationToken cancellationToken = default);
}
