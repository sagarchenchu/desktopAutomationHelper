using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Planning;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Determines whether a failed operation may be safely re-executed during recovery.
/// Mutations require structured proof of non-execution; uncertain protocol/transport
/// outcomes after send never authorize mutation replay unless AllowMutationReplay.
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
        {
            // Reads may retry even after uncertain transport; never "upgrade" via this helper alone.
            return true;
        }

        // Mutation / invoke / input: never replay without structured proof the action did not run.
        if (IsUncertainExecutionOutcome(failureClassification, errorMessage))
            return false;

        return HasStructuredNonExecutionProof(errorMessage);
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

    /// <summary>
    /// True when the failure leaves execution state uncertain (timeout, connection loss,
    /// or protocol responses that arrive after the request may already have been applied).
    /// </summary>
    public static bool IsUncertainExecutionOutcome(
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
        if (message.Length == 0)
            return false;

        return message.Contains("timed out", StringComparison.OrdinalIgnoreCase)
               || message.Contains("timeout", StringComparison.OrdinalIgnoreCase)
               || message.Contains("unable to reach driver", StringComparison.OrdinalIgnoreCase)
               || message.Contains("connection reset", StringComparison.OrdinalIgnoreCase)
               || message.Contains("connection aborted", StringComparison.OrdinalIgnoreCase)
               || message.Contains("an error occurred while sending the request", StringComparison.OrdinalIgnoreCase)
               // Protocol responses after send — request may already have mutated state.
               || message.Contains("empty response", StringComparison.OrdinalIgnoreCase)
               || message.Contains("invalid json", StringComparison.OrdinalIgnoreCase)
               || message.Contains("exceeded maximum size", StringComparison.OrdinalIgnoreCase)
               || message.Contains("content-length", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Backward-compatible alias. </summary>
    public static bool IsUncertainTransportFailure(
        UiFailureClassification? failureClassification,
        string? errorMessage) =>
        IsUncertainExecutionOutcome(failureClassification, errorMessage);

    public static bool HasStructuredNonExecutionProof(string? errorMessage)
    {
        var message = errorMessage ?? string.Empty;
        if (message.Length == 0)
            return false;

        return ContainsAny(
            message,
            "element not found",
            "not found",
            "no match",
            "child window",
            "dialog not",
            "window not active",
            "not yet active",
            "modal not",
            "not stable",
            "stabiliz",
            "timing");
    }

    /// <summary>
    /// Equivalent transitions must not upgrade a read/uncertain failure into a mutation
    /// when AllowMutationReplay is false.
    /// </summary>
    public static bool IsEquivalentTransitionAllowed(
        string originalOperation,
        string targetOperation,
        OperationsCatalogDto? catalog,
        UiFailureClassification? failureClassification,
        string? errorMessage,
        bool allowMutationReplay)
    {
        if (allowMutationReplay)
            return true;

        var targetIsRead = IsReadOrQuery(targetOperation, catalog);
        if (targetIsRead)
        {
            // Read→read is fine; still block only if we somehow lack a target.
            return true;
        }

        // Target is a mutation: never upgrade from a read/query failure.
        if (IsReadOrQuery(originalOperation, catalog))
            return false;

        // Original was already a mutation — only allow with structured non-execution proof
        // and not on uncertain protocol/transport outcomes.
        return IsSafeToReplay(
            targetOperation,
            catalog,
            failureClassification,
            errorMessage,
            allowMutationReplay: false);
    }

    private static bool LooksLikeReadOnlyName(string operation)
    {
        var name = (operation ?? string.Empty).Trim().ToLowerInvariant();
        return name is "listwindows" or "finduia" or "dumpuia" or "gettext" or "getvalue"
            or "getname" or "getcontroltype" or "screenshot" or "observe";
    }

    private static bool ContainsAny(string message, params string[] needles)
    {
        foreach (var needle in needles)
        {
            if (message.Contains(needle, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }
}
