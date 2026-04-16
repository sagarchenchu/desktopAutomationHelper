using DesktopAutomationDriver.Services;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="DriverContext"/> port computation and token generation.
/// </summary>
public class DriverContextTests
{
    [Fact]
    public void ComputeUserPort_IsInValidRange()
    {
        var port = DriverContext.ComputeUserPort("alice");
        Assert.InRange(port, DriverContext.PortBase, DriverContext.PortBase + DriverContext.PortRange - 1);
    }

    [Fact]
    public void ComputeUserPort_IsDeterministic()
    {
        var port1 = DriverContext.ComputeUserPort("bob");
        var port2 = DriverContext.ComputeUserPort("bob");
        Assert.Equal(port1, port2);
    }

    [Fact]
    public void ComputeUserPort_IsCaseInsensitive()
    {
        var lower = DriverContext.ComputeUserPort("charlie");
        var upper = DriverContext.ComputeUserPort("CHARLIE");
        Assert.Equal(lower, upper);
    }

    [Fact]
    public void ComputeUserPort_DifferentUsersGetDifferentPorts()
    {
        // With high probability two different user names should map to
        // different ports across a 10 000-port range.
        var portA = DriverContext.ComputeUserPort("user_alpha");
        var portB = DriverContext.ComputeUserPort("user_beta");
        Assert.NotEqual(portA, portB);
    }

    [Fact]
    public void DriverContext_BearerTokenIsNotEmpty()
    {
        var ctx = new DriverContext("testuser");
        Assert.False(string.IsNullOrWhiteSpace(ctx.BearerToken));
    }

    [Fact]
    public void DriverContext_BearerTokenIsUnique()
    {
        // Each instance (each driver start-up) should get a different token.
        var ctx1 = new DriverContext("testuser");
        var ctx2 = new DriverContext("testuser");
        Assert.NotEqual(ctx1.BearerToken, ctx2.BearerToken);
    }

    [Fact]
    public void DriverContext_ProbePortAlwaysIs9102()
    {
        var ctx = new DriverContext("testuser");
        Assert.Equal(9102, ctx.ProbePort);
    }

    [Fact]
    public void DriverContext_ProbePortActiveDefaultsFalse()
    {
        var ctx = new DriverContext("testuser");
        Assert.False(ctx.ProbePortActive);
    }

    [Fact]
    public void DriverContext_UsernameIsSet()
    {
        var ctx = new DriverContext("myuser");
        Assert.Equal("myuser", ctx.Username);
    }
}
