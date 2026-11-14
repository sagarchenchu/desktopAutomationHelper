using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Builds Windows ACLs that restrict bootstrap objects to the intended user SID
/// plus required local service identities (SYSTEM, Administrators).
/// </summary>
public static class BootstrapAcl
{
    public static FileSecurity CreateFileSecurity(SecurityIdentifier userSid)
    {
        ArgumentNullException.ThrowIfNull(userSid);

        var security = new FileSecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

        AddFileRule(security, userSid, FileSystemRights.FullControl);
        AddFileRule(security, new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null), FileSystemRights.FullControl);
        AddFileRule(security, new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null), FileSystemRights.FullControl);

        return security;
    }

    public static PipeSecurity CreatePipeSecurity(SecurityIdentifier userSid)
    {
        ArgumentNullException.ThrowIfNull(userSid);

        var security = new PipeSecurity();
        // Deny-by-default: only explicitly allowed identities can connect.
        AddPipeRule(security, userSid, PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance);
        AddPipeRule(security, new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null), PipeAccessRights.FullControl);
        AddPipeRule(security, new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null), PipeAccessRights.FullControl);

        return security;
    }

    /// <summary>
    /// Test helper: ACL that grants only <paramref name="allowedSid"/> (plus SYSTEM)
    /// so a different interactive user identity is denied by Windows.
    /// </summary>
    public static PipeSecurity CreatePipeSecurityAllowingOnly(SecurityIdentifier allowedSid)
    {
        ArgumentNullException.ThrowIfNull(allowedSid);

        var security = new PipeSecurity();
        AddPipeRule(security, allowedSid, PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance);
        AddPipeRule(security, new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null), PipeAccessRights.FullControl);
        return security;
    }

    public static FileSecurity CreateFileSecurityAllowingOnly(SecurityIdentifier allowedSid)
    {
        ArgumentNullException.ThrowIfNull(allowedSid);

        var security = new FileSecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        AddFileRule(security, allowedSid, FileSystemRights.FullControl);
        AddFileRule(security, new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null), FileSystemRights.FullControl);
        return security;
    }

    private static void AddFileRule(FileSecurity security, SecurityIdentifier sid, FileSystemRights rights)
    {
        security.AddAccessRule(new FileSystemAccessRule(
            sid,
            rights,
            InheritanceFlags.None,
            PropagationFlags.None,
            AccessControlType.Allow));
    }

    private static void AddPipeRule(PipeSecurity security, SecurityIdentifier sid, PipeAccessRights rights)
    {
        security.AddAccessRule(new PipeAccessRule(sid, rights, AccessControlType.Allow));
    }
}
