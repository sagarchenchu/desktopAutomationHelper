using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Tests;

public class SecretRedactorTests
{
    [Fact]
    public void RedactsBearerTokensAndJsonSecrets()
    {
        const string token = "SuperSecretTokenValue123";
        var input =
            $"Authorization: Bearer {token}; " +
            $"{{\"token\":\"{token}\",\"authorizationHeader\":\"Bearer {token}\"}}";

        var redacted = SecretRedactor.RedactKnownSecrets(input, token);

        Assert.DoesNotContain(token, redacted, StringComparison.Ordinal);
        Assert.Contains("[REDACTED]", redacted, StringComparison.Ordinal);
    }

    [Fact]
    public void DriverExceptions_AreRedacted()
    {
        var ex = new DriverConnectionException("token=SuperSecretTokenValue123 Authorization: Bearer SuperSecretTokenValue123");
        Assert.DoesNotContain("SuperSecretTokenValue123", ex.Message, StringComparison.Ordinal);
    }
}
