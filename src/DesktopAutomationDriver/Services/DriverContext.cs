using System.Security.Cryptography;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Computes a deterministic, per-Windows-user port for the automation driver
/// and generates a one-time Bearer token at process startup.
/// </summary>
public sealed class DriverContext : IDriverContext
{
    // Ports are drawn from 30000–39999 to avoid well-known service ports.
    internal const int PortBase = 30000;
    internal const int PortRange = 10000;
    internal const int ProbePortNumber = 9102;

    /// <inheritdoc/>
    public string Username { get; }

    /// <inheritdoc/>
    public int MainPort { get; }

    /// <inheritdoc/>
    public int ProbePort => ProbePortNumber;

    /// <inheritdoc/>
    public bool ProbePortActive { get; set; }

    /// <inheritdoc/>
    public string BearerToken { get; }

    public DriverContext() : this(Environment.UserName) { }

    /// <summary>
    /// Internal constructor that accepts an explicit username so tests can
    /// verify port-computation logic without relying on the running OS user.
    /// </summary>
    internal DriverContext(string username)
    {
        Username = username;
        MainPort = ComputeUserPort(username);
        BearerToken = GenerateToken();
    }

    /// <summary>
    /// Maps a Windows username to a stable port in [PortBase, PortBase+PortRange).
    /// Uses FNV-1a 32-bit over the UTF-8 bytes of the lower-cased username so
    /// the mapping is stable across .NET versions and handles non-ASCII names.
    /// </summary>
    internal static int ComputeUserPort(string username)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(username.ToLowerInvariant());
        unchecked
        {
            uint hash = 2166136261u; // FNV-1a 32-bit offset basis
            foreach (var b in bytes)
            {
                hash ^= b;
                hash *= 16777619u; // FNV prime
            }
            return PortBase + (int)(hash % (uint)PortRange);
        }
    }

    private static string GenerateToken()
    {
        var bytes = RandomNumberGenerator.GetBytes(32);
        return Convert.ToBase64String(bytes);
    }
}
