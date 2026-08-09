using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.Driver;

public interface IDriverCatalogClient
{
    Task<StatusValueDto> GetStatusAsync(DriverConnection connection, CancellationToken cancellationToken = default);

    Task<OperationsCatalogDto> GetOperationsAsync(DriverConnection connection, CancellationToken cancellationToken = default);
}
