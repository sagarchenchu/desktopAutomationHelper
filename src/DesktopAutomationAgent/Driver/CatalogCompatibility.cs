using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.Driver;

public static class CatalogCompatibility
{
    public static void Validate(OperationsCatalogDto catalog, int expectedSchemaVersion)
    {
        ArgumentNullException.ThrowIfNull(catalog);

        if (catalog.SchemaVersion != expectedSchemaVersion)
        {
            throw new DriverCatalogException(
                $"Unsupported catalog schemaVersion {catalog.SchemaVersion}; expected {expectedSchemaVersion}.");
        }

        if (catalog.Operations is null || catalog.Operations.Count == 0)
            throw new DriverCatalogException("Operation catalog is empty.");

        var canonical = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var operation in catalog.Operations)
        {
            if (string.IsNullOrWhiteSpace(operation.Name))
                throw new DriverCatalogException("Operation catalog contains an entry with an empty name.");

            if (!canonical.TryAdd(operation.Name, operation.Name))
            {
                throw new DriverCatalogException(
                    $"Duplicate canonical operation name '{operation.Name}'.");
            }
        }

        var claimedAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var operation in catalog.Operations)
        {
            foreach (var alias in operation.Aliases.Concat(operation.DeprecatedAliases).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(alias))
                    continue;

                if (canonical.TryGetValue(alias, out var owner)
                    && !string.Equals(owner, operation.Name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new DriverCatalogException(
                        $"Alias '{alias}' on operation '{operation.Name}' collides with canonical name '{owner}'.");
                }

                if (claimedAliases.TryGetValue(alias, out var previous)
                    && !string.Equals(previous, operation.Name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new DriverCatalogException(
                        $"Alias '{alias}' is claimed by both '{previous}' and '{operation.Name}'.");
                }

                claimedAliases[alias] = operation.Name;
            }
        }
    }
}
