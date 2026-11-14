namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Driver-owned human / user-presence confirmation for vision approval.
/// Same-user ACL file reads are not presence — agents must not self-approve via file.
/// </summary>
public interface IOperatorPresenceService
{
    bool TryConfirmVisionApproval(OperatorPresenceRequest request, out string denyReason);
}

public sealed record OperatorPresenceRequest(
    string CandidateId,
    string Action,
    string Label);

/// <summary>
/// Interactive WinForms Yes/No confirmation owned by the driver process.
/// Not satisfiable by reading an ACL secret file.
/// </summary>
public sealed class InteractiveOperatorPresenceService : IOperatorPresenceService
{
    public bool TryConfirmVisionApproval(OperatorPresenceRequest request, out string denyReason)
    {
        ArgumentNullException.ThrowIfNull(request);
        denyReason = string.Empty;

        var message =
            "Approve vision coordinate action?\r\n\r\n" +
            $"Candidate: {request.CandidateId}\r\n" +
            $"Action: {request.Action}\r\n" +
            $"Label: {request.Label}\r\n\r\n" +
            "Only a human at this desktop should confirm. Automation cannot self-approve via ACL files.";

        var result = System.Windows.Forms.DialogResult.No;
        void Show()
        {
            result = System.Windows.Forms.MessageBox.Show(
                message,
                "Desktop Automation Driver — Vision approval",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning,
                System.Windows.Forms.MessageBoxDefaultButton.Button2);
        }

        if (System.Windows.Forms.Application.MessageLoop)
        {
            Show();
        }
        else
        {
            var thread = new Thread(() =>
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                Show();
            })
            {
                IsBackground = false,
                Name = "VisionOperatorPresence"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        if (result != System.Windows.Forms.DialogResult.Yes)
        {
            denyReason = "Operator presence denied or cancelled (interactive confirmation required).";
            return false;
        }

        return true;
    }
}

/// <summary>Test double — allow/deny without UI.</summary>
public sealed class ScriptedOperatorPresenceService : IOperatorPresenceService
{
    public bool Allow { get; set; } = true;
    public int ConfirmCalls { get; private set; }

    public bool TryConfirmVisionApproval(OperatorPresenceRequest request, out string denyReason)
    {
        ConfirmCalls++;
        if (Allow)
        {
            denyReason = string.Empty;
            return true;
        }

        denyReason = "Operator presence denied (scripted).";
        return false;
    }
}
