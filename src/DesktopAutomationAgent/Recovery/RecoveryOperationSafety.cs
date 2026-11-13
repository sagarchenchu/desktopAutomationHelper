using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Planning;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Determines whether a failed operation may be safely re-executed during recovery.
/// Timed-out / uncertain mutation/invoke/input must never be replayed unless AllowMutationReplay.
/// </summary>
public static class RecoveryOperationSafety
{
    public static bool IsSafeToReplay(
        string operation,
        OperationsCatalogDto? catalog,
        UiFailureClassification? failureClassification,
        string? errorMessage,
        bool allowMutationReplay)
    {
        if (allowMutationReplay)
            return true;

        if (IsReadOrQuery(operation, catalog))
            return true;

        // Mutation / invoke / input: never replay on uncertain transport outcomes.
        // Classifier must already have required structured proof (e.g. element not found)
        // before recovery enters the ladder for non-timeout failures.
        return !IsUncertainTransportFailure(failureClassification, errorMessage);
    }

    public static bool IsReadOrQuery(string operation, OperationsCatalogDto? catalog)
    {
        if (catalog is not null)
        {
            var descriptor = catalog.Operations.FirstOrDefault(
                op => string.Equals(op.Name, operation, StringComparison.OrdinalIgnoreCase));
            if (descriptor is not null)
                return PlanningRiskPolicy.IsSafeReadOrQuery(descriptor);
        }

        return LooksLikeReadOnlyName(operation);
    }

    public static bool IsUncertainTransportFailure(
        UiFailureClassification? failureClassification,
        string? errorMessage)
    {
        if (failureClassification is UiFailureClassification.ExecutionTimeout
            or UiFailureClassification.DriverUnavailable
            or UiFailureClassification.Cancelled)
        {
            return true;
        }

        var message = errorMessage ?? string.Empty;
        return message.Contains("timed out", StringComparison.OrdinalIgnoreCase)
               || message.Contains("timeout", StringComparison.OrdinalIgnoreCase)
               || message.Contains("unable to reach driver", StringComparison.OrdinalIgnoreCase)
               || message.Contains("connection reset", StringComparison.OrdinalIgnoreCase)
               || message.Contains("connection aborted", StringComparison.OrdinalIgnoreCase)
               || message.Contains("an error occurred while sending the request", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeReadOnlyName(string operation)
    {
        var name = (operation ?? string.Empty).Trim().ToLowerInvariant();
        return name is "listwindows" or "finduia" or "dumpuia" or "gettext" or "getvalue"
            or "getname" or "getcontroltype" or "screenshot" or "observe";
    }
}
