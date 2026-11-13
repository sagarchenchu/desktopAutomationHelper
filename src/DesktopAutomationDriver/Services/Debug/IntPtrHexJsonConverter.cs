using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Serializes <see cref="IntPtr"/> as a hex string (e.g. <c>0x1234ABCD</c>) so System.Text.Json
/// does not 500 on /debug/inspect. Deserializes hex or decimal strings back to <see cref="IntPtr"/>.
/// </summary>
public sealed class IntPtrHexJsonConverter : JsonConverter<IntPtr>
{
    public override IntPtr Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.Null:
                return IntPtr.Zero;
            case JsonTokenType.Number:
                if (reader.TryGetInt64(out var number))
                    return new IntPtr(number);
                throw new JsonException("IntPtr number value is out of range.");
            case JsonTokenType.String:
                return Parse(reader.GetString());
            default:
                throw new JsonException($"Unexpected token {reader.TokenType} when reading IntPtr.");
        }
    }

    public override void Write(Utf8JsonWriter writer, IntPtr value, JsonSerializerOptions options)
    {
        if (value == IntPtr.Zero)
        {
            writer.WriteStringValue("0");
            return;
        }

        writer.WriteStringValue($"0x{value.ToInt64():X}");
    }

    public static string ToHex(IntPtr value) =>
        value == IntPtr.Zero ? "0" : $"0x{value.ToInt64():X}";

    public static IntPtr Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "0")
            return IntPtr.Zero;

        var trimmed = value.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            return new IntPtr(long.Parse(trimmed.AsSpan(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }

        if (long.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dec))
            return new IntPtr(dec);

        return new IntPtr(long.Parse(trimmed, NumberStyles.HexNumber, CultureInfo.InvariantCulture));
    }
}
