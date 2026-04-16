using DesktopAutomationDriver.Services;
using System.Text.Json;

namespace DesktopAutomationDriver.Middleware;

/// <summary>
/// ASP.NET Core middleware that enforces Bearer-token authentication on all
/// routes except <c>GET /verify</c>.
///
/// Clients must supply the header:
/// <code>Authorization: Bearer &lt;token&gt;</code>
/// where <c>&lt;token&gt;</c> is the value returned by <c>GET /verify</c>
/// and logged to the console at driver startup.
/// </summary>
public class BearerTokenMiddleware
{
    private readonly RequestDelegate _next;
    private readonly IDriverContext _driverContext;

    // Paths that are exempt from authentication.
    private static readonly string[] PublicPaths = ["/verify"];

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public BearerTokenMiddleware(RequestDelegate next, IDriverContext driverContext)
    {
        _next = next;
        _driverContext = driverContext;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // Allow public paths without authentication.
        var path = context.Request.Path.Value ?? string.Empty;
        foreach (var publicPath in PublicPaths)
        {
            if (path.StartsWith(publicPath, StringComparison.OrdinalIgnoreCase))
            {
                await _next(context);
                return;
            }
        }

        // Validate Authorization header.
        var authHeader = context.Request.Headers.Authorization.FirstOrDefault();
        if (string.IsNullOrEmpty(authHeader) ||
            !authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            await WriteUnauthorized(context,
                "Missing or malformed Authorization header. " +
                "Expected: Authorization: Bearer <token>. " +
                "Retrieve the token from GET /verify or from the driver startup log.");
            return;
        }

        var providedToken = authHeader["Bearer ".Length..].Trim();
        if (!string.Equals(providedToken, _driverContext.BearerToken, StringComparison.Ordinal))
        {
            await WriteUnauthorized(context, "Invalid Bearer token.");
            return;
        }

        await _next(context);
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
