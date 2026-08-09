namespace DesktopAutomationAgent.Driver;

public interface IDriverConnectionResolver
{
    Task<DriverConnection> ResolveAsync(CancellationToken cancellationToken = default);
}
