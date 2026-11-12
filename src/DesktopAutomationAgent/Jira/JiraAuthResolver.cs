using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Resolves Jira credentials from options / env / optional OS secret store.
/// Never returns secrets in exception messages.
/// </summary>
internal static class JiraAuthResolver
{
    public static AuthenticationHeaderValue Resolve(JiraOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var mode = (options.AuthMode ?? "basic").Trim().ToLowerInvariant();
        var secretFromStore = TryReadCredentialSecret(options.CredentialTarget);

        return mode switch
        {
            "basic" => BuildBasic(options, secretFromStore),
            "bearer" => BuildBearer(options.BearerToken, secretFromStore, "BearerToken"),
            "pat" => BuildBearer(
                options.PersonalAccessToken ?? options.BearerToken,
                secretFromStore,
                "PersonalAccessToken"),
            _ => throw new AgentConfigurationException("Jira:AuthMode must be one of: basic, bearer, pat.")
        };
    }

    private static AuthenticationHeaderValue BuildBasic(JiraOptions options, string? secretFromStore)
    {
        if (string.IsNullOrWhiteSpace(options.Email))
            throw new AgentConfigurationException("Jira:Email is required for AuthMode=basic.");

        var token = FirstNonEmpty(options.ApiToken, secretFromStore);
        if (string.IsNullOrWhiteSpace(token))
        {
            throw new AgentConfigurationException(
                "Jira:ApiToken (or Jira:CredentialTarget) is required for AuthMode=basic.");
        }

        var raw = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{options.Email.Trim()}:{token}"));
        return new AuthenticationHeaderValue("Basic", raw);
    }

    private static AuthenticationHeaderValue BuildBearer(string? configured, string? secretFromStore, string label)
    {
        var token = FirstNonEmpty(configured, secretFromStore);
        if (string.IsNullOrWhiteSpace(token))
        {
            throw new AgentConfigurationException(
                $"Jira:{label} (or Jira:CredentialTarget) is required for the configured AuthMode.");
        }

        return new AuthenticationHeaderValue("Bearer", token);
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value;
        }

        return null;
    }

    private static string? TryReadCredentialSecret(string? target)
    {
        if (string.IsNullOrWhiteSpace(target))
            return null;

        if (!OperatingSystem.IsWindows())
        {
            throw new AgentConfigurationException(
                "Jira:CredentialTarget is only supported on Windows.");
        }

        if (!NativeCredentialReader.TryReadPassword(target.Trim(), out var secret) || string.IsNullOrEmpty(secret))
        {
            throw new AgentConfigurationException(
                "Jira:CredentialTarget did not resolve a usable secret from the OS secret store.");
        }

        return secret;
    }
}

internal static class NativeCredentialReader
{
    private const int CredTypeGeneric = 1;

    public static bool TryReadPassword(string target, out string? password)
    {
        password = null;
        if (!CredRead(target, CredTypeGeneric, 0, out var credPtr) || credPtr == IntPtr.Zero)
            return false;

        try
        {
            var cred = Marshal.PtrToStructure<NativeCredential>(credPtr);
            if (cred.CredentialBlob == IntPtr.Zero || cred.CredentialBlobSize == 0)
                return false;

            password = Marshal.PtrToStringUni(cred.CredentialBlob, (int)cred.CredentialBlobSize / 2);
            return !string.IsNullOrEmpty(password);
        }
        finally
        {
            CredFree(credPtr);
        }
    }

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredRead(string target, int type, int flags, out IntPtr credential);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern void CredFree(IntPtr credential);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags;
        public uint Type;
        public string TargetName;
        public string Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }
}
