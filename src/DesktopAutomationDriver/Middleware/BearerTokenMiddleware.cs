using DesktopAutomationDriver.Services;
using System.Text.Json;

namespace DesktopAutomationDriver.Middleware;

/// <summary>
/// ASP.NET Core middleware that enforces Bearer-token authentication on all
/// routes except <c>GET /verify</c>.
///
/// <c>POST /vision/approve</c> requires a separate operator channel
/// (<c>X-Operator-Token</c>) and rejects the automation bearer for approval.
/// </summary>
public class BearerTokenMiddleware
{
    private readonly RequestDelegate _next;
    private readonly IDriverContext _driverContext;

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
        var path = context.Request.Path.Value ?? string.Empty;
        foreach (var publicPath in PublicPaths)
        {
            if (path.StartsWith(publicPath, StringComparison.OrdinalIgnoreCase))
            {
                await _next(context);
                return;
            }
        }

        // Operator approval channel — must not accept the automation bearer alone.
        if (IsVisionApprovePath(path, context.Request.Method))
        {
            var operatorToken = context.Request.Headers["X-Operator-Token"].FirstOrDefault()?.Trim();
            if (string.IsNullOrEmpty(operatorToken)
                || !string.Equals(operatorToken, _driverContext.OperatorToken, StringComparison.Ordinal))
            {
                await WriteForbidden(context,
                    "Vision approval requires X-Operator-Token (operator channel). " +
                    "The automation Bearer token cannot approve candidates.");
                return;
            }

            // Still require a valid automation bearer for non-public routes, but approval
            // additionally requires the distinct operator token above.
            if (!TryValidateBearer(context, out var bearerError))
            {
                await WriteUnauthorized(context, bearerError!);
                return;
            }

            await _next(context);
            return;
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
            error =
                "Missing or malformed Authorization header. " +
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

    private static bool IsVisionApprovePath(string path, string method) =>
        path.StartsWith("/vision/approve", StringComparison.OrdinalIgnoreCase)
        && string.Equals(method, "POST", StringComparison.OrdinalIgnoreCase);

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

    private static async Task WriteForbidden(HttpContext context, string message)
    {
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        context.Response.ContentType = "application/json";
        var body = JsonSerializer.Serialize(
            new
            {
                status = 403,
                value = new { error = "forbidden", message }
            },
            JsonOpts);
        await context.Response.WriteAsync(body);
    }
}
