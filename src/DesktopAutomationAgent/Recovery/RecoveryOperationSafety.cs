using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Planning;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Determines whether a failed operation may be safely re-executed during recovery.
/// Mutations require <em>typed</em> driver non-execution proof (executionState / failureCode);
/// free-text error phrases never authorize mutation replay.
/// </summary>
public static class RecoveryOperationSafety
{
    public const string ExecutionStateNotStarted = "notStarted";
    public const string FailureCodeLocatorNotFoundBeforeInvoke = "locatorNotFoundBeforeInvoke";

    public static bool IsSafeToReplay(
        string operation,
        OperationsCatalogDto? catalog,
        UiFailureClassification? failureClassification,
        string? errorMessage,
        bool allowMutationReplay,
        string? executionState = null,
        string? failureCode = null)
    {
        if (allowMutationReplay)
            return true;

        if (IsReadOrQuery(operation, catalog))
            return true;

        // Mutation / invoke / input: never replay on uncertain outcomes, and never on free-text alone.
        if (IsUncertainExecutionOutcome(failureClassification, errorMessage))
            return false;

        return HasTypedNonExecutionProof(executionState, failureCode);
    }

    public static bool IsSafeToReplay(StepRunResult failedStep, OperationsCatalogDto? catalog, bool allowMutationReplay) =>
        IsSafeToReplay(
            failedStep.Operation,
            catalog,
            failedStep.Classification,
            Combine(failedStep.Error, failedStep.DriverReason),
            allowMutationReplay,
            failedStep.ExecutionState,
            failedStep.FailureCode);

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

    /// <summary>
    /// Typed driver fields only. Free-text phrases such as "not found" / "timing" do <b>not</b> count.
    /// </summary>
    public static bool HasTypedNonExecutionProof(string? executionState, string? failureCode)
    {
        if (!string.IsNullOrWhiteSpace(executionState)
            && string.Equals(executionState.Trim(), ExecutionStateNotStarted, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(failureCode)
            && string.Equals(
                failureCode.Trim(),
                FailureCodeLocatorNotFoundBeforeInvoke,
                StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
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
        bool allowMutationReplay,
        string? executionState = null,
        string? failureCode = null)
    {
        if (allowMutationReplay)
            return true;

        if (IsReadOrQuery(targetOperation, catalog))
            return true;

        if (IsReadOrQuery(originalOperation, catalog))
            return false;

        return IsSafeToReplay(
            targetOperation,
            catalog,
            failureClassification,
            errorMessage,
            allowMutationReplay: false,
            executionState,
            failureCode);
    }

    private static bool LooksLikeReadOnlyName(string operation)
    {
        var name = (operation ?? string.Empty).Trim().ToLowerInvariant();
        return name is "listwindows" or "finduia" or "dumpuia" or "gettext" or "getvalue"
            or "getname" or "getcontroltype" or "screenshot" or "observe";
    }

    private static string Combine(string? a, string? b)
    {
        if (string.IsNullOrWhiteSpace(a)) return b ?? string.Empty;
        if (string.IsNullOrWhiteSpace(b)) return a!;
        return a + " " + b;
    }
}
