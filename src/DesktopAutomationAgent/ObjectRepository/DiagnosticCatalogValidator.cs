using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.ObjectRepository;

internal static class DiagnosticCatalogValidator
{
    public static OperationDescriptorDto RequireDiagnosticOperation(
        OperationsCatalogDto catalog,
        string operationName)
    {
        ArgumentNullException.ThrowIfNull(catalog);

        var descriptor = catalog.Operations.FirstOrDefault(
            op => string.Equals(op.Name, operationName, StringComparison.OrdinalIgnoreCase));

        if (descriptor is null)
        {
            throw new DriverCatalogException(
                $"Operation catalog does not include required diagnostic operation '{operationName}'.");
        }

        if (descriptor.Deprecated)
        {
            throw new DriverCatalogException(
                $"Operation '{operationName}' is deprecated and cannot be used.");
        }

        if (!string.Equals(descriptor.OperationType, "diagnostic", StringComparison.OrdinalIgnoreCase))
        {
            throw new DriverCatalogException(
                $"Operation '{operationName}' must have operationType 'diagnostic'.");
        }

        return descriptor;
    }
}
