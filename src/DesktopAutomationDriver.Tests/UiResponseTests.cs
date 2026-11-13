using DesktopAutomationDriver.Models.Response;

namespace DesktopAutomationDriver.Tests;

public class UiResponseTests
{
    [Fact]
    public void FromOperationResult_WhenPayloadSuccessFalse_EnvelopeSuccessFalse()
    {
        var payload = new
        {
            operation = "clickmenuuia",
            success = false,
            reason = "element-not-found",
            message = "Native UIA resolver could not find an element for the locator."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.False(response.Success);
        Assert.Equal(payload, response.Value);
        Assert.Equal("element-not-found", response.Reason);
        Assert.Equal(payload.message, response.Error);
        // Free-text alone must not manufacture typed proof.
        Assert.Null(response.ExecutionState);
        Assert.Null(response.FailureCode);
    }

    [Fact]
    public void FromOperationResult_WhenFoundFalse_SetsTypedNonExecutionProof()
    {
        var payload = new
        {
            success = false,
            found = false,
            message = "Element not found for clickuia."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.False(response.Success);
        Assert.Equal(UiResponse.ExecutionStateNotStarted, response.ExecutionState);
        Assert.Equal(UiResponse.FailureCodeLocatorNotFoundBeforeInvoke, response.FailureCode);
    }

    [Fact]
    public void FromOperationResult_WhenMatchCountZero_SetsTypedNonExecutionProof()
    {
        var payload = new
        {
            success = false,
            matchCount = 0,
            message = "No matches."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.Equal(UiResponse.ExecutionStateNotStarted, response.ExecutionState);
        Assert.Equal(UiResponse.FailureCodeLocatorNotFoundBeforeInvoke, response.FailureCode);
    }

    [Fact]
    public void FromOperationResult_PostActionNotFound_DoesNotSetNotStarted()
    {
        // Post-action business/visibility miss — action may already have run.
        var payload = new
        {
            success = false,
            found = true,
            message = "Target was found, but driver could not verify visibility after scrolling."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.False(response.Success);
        Assert.Null(response.ExecutionState);
        Assert.Null(response.FailureCode);
    }

    [Fact]
    public void FromOperationResult_ReceiptNotFoundFreeText_DoesNotSetNotStarted()
    {
        var payload = new
        {
            success = false,
            message = "Payment receipt not found after click."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.Null(response.ExecutionState);
        Assert.Null(response.FailureCode);
    }

    [Fact]
    public void Fail_FreeTextNotFound_DoesNotSetNotStarted()
    {
        var response = UiResponse.Fail("Element not found for clickuia.", "shot.png");

        Assert.False(response.Success);
        Assert.Equal("Element not found for clickuia.", response.Error);
        Assert.Null(response.ExecutionState);
        Assert.Null(response.FailureCode);
    }

    [Fact]
    public void Fail_WasFoundButInvocationFailed_DoesNotSetNotStarted()
    {
        var response = UiResponse.Fail(
            "Target was found but invocation failed: pattern not supported.");

        Assert.Null(response.ExecutionState);
        Assert.Null(response.FailureCode);
    }

    [Fact]
    public void FailLocatorNotFoundBeforeInvoke_SetsTypedProof()
    {
        var response = UiResponse.FailLocatorNotFoundBeforeInvoke(
            "No element matched locator.",
            reason: "element-not-found");

        Assert.False(response.Success);
        Assert.Equal(UiResponse.ExecutionStateNotStarted, response.ExecutionState);
        Assert.Equal(UiResponse.FailureCodeLocatorNotFoundBeforeInvoke, response.FailureCode);
    }

    [Fact]
    public void FromOperationResult_WhenPayloadSuccessTrue_EnvelopeSuccessTrue()
    {
        var payload = new
        {
            operation = "clickmenuuia",
            success = true,
            strategy = "invoke-pattern"
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.True(response.Success);
        Assert.Equal(payload, response.Value);
        Assert.Null(response.Error);
    }

    [Fact]
    public void FromOperationResult_WhenPayloadHasNoSuccessFlag_EnvelopeSuccessTrue()
    {
        var payload = new { selected = "Option A" };

        var response = UiResponse.FromOperationResult(payload);

        Assert.True(response.Success);
        Assert.Equal(payload, response.Value);
    }
}
