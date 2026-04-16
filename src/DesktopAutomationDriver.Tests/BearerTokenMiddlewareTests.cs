using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Http;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="BearerTokenMiddleware"/>.
/// </summary>
public class BearerTokenMiddlewareTests
{
    private const string ValidToken = "test-valid-token-abc123";

    private static Mock<IDriverContext> BuildContext(string token = ValidToken)
    {
        var mock = new Mock<IDriverContext>();
        mock.Setup(c => c.BearerToken).Returns(token);
        return mock;
    }

    private static (BearerTokenMiddleware middleware, DefaultHttpContext httpContext, bool nextCalled)
        BuildMiddleware(string? authHeaderValue = null, string path = "/session",
                        string token = ValidToken)
    {
        bool nextCalled = false;
        RequestDelegate next = ctx => { nextCalled = true; return Task.CompletedTask; };

        var ctx = BuildContext(token);
        var middleware = new BearerTokenMiddleware(next, ctx.Object);

        var httpContext = new DefaultHttpContext();
        httpContext.Request.Path = path;
        if (authHeaderValue != null)
            httpContext.Request.Headers.Authorization = authHeaderValue;

        // Provide a writable body stream so WriteAsync doesn't throw.
        httpContext.Response.Body = new System.IO.MemoryStream();

        return (middleware, httpContext, nextCalled);
    }

    [Fact]
    public async Task MissingAuthHeader_Returns401()
    {
        var (middleware, ctx, _) = BuildMiddleware(authHeaderValue: null);
        await middleware.InvokeAsync(ctx);
        Assert.Equal(401, ctx.Response.StatusCode);
    }

    [Fact]
    public async Task WrongToken_Returns401()
    {
        var (middleware, ctx, _) = BuildMiddleware("Bearer wrong-token");
        await middleware.InvokeAsync(ctx);
        Assert.Equal(401, ctx.Response.StatusCode);
    }

    [Fact]
    public async Task CorrectToken_CallsNext()
    {
        bool nextCalled = false;
        RequestDelegate next = _ => { nextCalled = true; return Task.CompletedTask; };

        var driverCtx = BuildContext().Object;
        var middleware = new BearerTokenMiddleware(next, driverCtx);

        var httpContext = new DefaultHttpContext();
        httpContext.Request.Path = "/session";
        httpContext.Request.Headers.Authorization = $"Bearer {ValidToken}";
        httpContext.Response.Body = new System.IO.MemoryStream();

        await middleware.InvokeAsync(httpContext);

        Assert.True(nextCalled);
        Assert.Equal(200, httpContext.Response.StatusCode);
    }

    [Fact]
    public async Task VerifyPath_IsExemptFromAuth()
    {
        bool nextCalled = false;
        RequestDelegate next = _ => { nextCalled = true; return Task.CompletedTask; };

        var driverCtx = BuildContext().Object;
        var middleware = new BearerTokenMiddleware(next, driverCtx);

        var httpContext = new DefaultHttpContext();
        httpContext.Request.Path = "/verify";
        // Deliberately omit the Authorization header.
        httpContext.Response.Body = new System.IO.MemoryStream();

        await middleware.InvokeAsync(httpContext);

        Assert.True(nextCalled, "/verify should pass through without auth.");
    }

    [Fact]
    public async Task NonBearerScheme_Returns401()
    {
        var (middleware, ctx, _) = BuildMiddleware($"Basic dXNlcjpwYXNz");
        await middleware.InvokeAsync(ctx);
        Assert.Equal(401, ctx.Response.StatusCode);
    }

    [Fact]
    public async Task BearerTokenWithExtraWhitespace_IsValidated()
    {
        // Clients that add extra spaces around the token value should still work.
        bool nextCalled = false;
        RequestDelegate next = _ => { nextCalled = true; return Task.CompletedTask; };

        var driverCtx = BuildContext().Object;
        var middleware = new BearerTokenMiddleware(next, driverCtx);

        var httpContext = new DefaultHttpContext();
        httpContext.Request.Path = "/status";
        httpContext.Request.Headers.Authorization = $"Bearer   {ValidToken}   ";
        httpContext.Response.Body = new System.IO.MemoryStream();

        await middleware.InvokeAsync(httpContext);

        Assert.True(nextCalled);
    }
}
