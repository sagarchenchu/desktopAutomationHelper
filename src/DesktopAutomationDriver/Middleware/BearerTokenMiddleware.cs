using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Bootstrap;
using Microsoft.Extensions.Options;
using System.Text.Json;

namespace DesktopAutomationDriver.Middleware;

/// <summary>
/// ASP.NET Core middleware that enforces Bearer-token authentication on all
/// routes except <c>GET /verify</c> (which remains reachable but does not
/// return the token in secure mode).
///
/// <c>POST /vision/approve</c> requires a valid automation bearer; the coordinator
/// then runs operator confirmation UX (UIA-automatable dialog — not OS human-presence).
/// </summary>
public class BearerTokenMiddleware
{
    private readonly RequestDelegate _next;
    private readonly IDriverContext _driverContext;
    private readonly BootstrapOptions _bootstrap;

    private static readonly string[] PublicPaths = ["/verify"];

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public BearerTokenMiddleware(RequestDelegate next, IDriverContext driverContext)
        : this(next, driverContext, Options.Create(new BootstrapOptions()))
    {
    }

    public BearerTokenMiddleware(
        RequestDelegate next,
        IDriverContext driverContext,
        IOptions<BootstrapOptions> bootstrapOptions)
    {
        _next = next;
        _driverContext = driverContext;
        _bootstrap = bootstrapOptions?.Value ?? new BootstrapOptions();
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var path = context.Request.Path.Value ?? string.Empty;
        foreach (var publicPath in PublicPaths)
        {
            if (path.StartsWith(publicPath, StringComparison.OrdinalIgnoreCase))
            {
                await _next(context);
                return;
            }
        }

        if (!TryValidateBearer(context, out var error))
        {
            await WriteUnauthorized(context, error!);
            return;
        }

        await _next(context);
    }

    private bool TryValidateBearer(HttpContext context, out string? error)
    {
        error = null;
        var authHeader = context.Request.Headers.Authorization.FirstOrDefault();
        if (string.IsNullOrEmpty(authHeader) ||
            !authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            error = _bootstrap.IsSecure
                ? "Missing or malformed Authorization header. " +
                  "Expected: Authorization: Bearer <token>. " +
                  "Retrieve the token via the SID ACL-protected secure bootstrap channel " +
                  "(named pipe / connection descriptor) — not from HTTP /verify."
                : "Missing or malformed Authorization header. " +
                  "Expected: Authorization: Bearer <token>. " +
                  "Retrieve the token from GET /verify or from the driver startup log.";
            return false;
        }

        var providedToken = authHeader["Bearer ".Length..].Trim();
        if (!string.Equals(providedToken, _driverContext.BearerToken, StringComparison.Ordinal))
        {
            error = "Invalid Bearer token.";
            return false;
        }

        return true;
    }

    private static async Task WriteUnauthorized(HttpContext context, string message)
    {
        context.Response.StatusCode = StatusCodes.Status401Unauthorized;
        context.Response.ContentType = "application/json";
        var body = JsonSerializer.Serialize(
            new
            {
                status = 401,
                value = new { error = "unauthorized", message }
            },
            JsonOpts);
        await context.Response.WriteAsync(body);
    }
}
