using System.Security.Cryptography;
using System.Text;

namespace DesktopAutomationAgent.Knowledge;

internal static class KnowledgeHash
{
    public static string Sha256Hex(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    public static string Sha256Hex(ReadOnlySpan<byte> bytes)
    {
        var hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public static string StableIdPrefix(string hex, int length = 16)
    {
        if (string.IsNullOrWhiteSpace(hex))
            throw new ArgumentException("Hash hex is required.", nameof(hex));
        if (length < 8 || length > hex.Length)
            length = Math.Min(16, hex.Length);
        return hex[..length];
    }
}
