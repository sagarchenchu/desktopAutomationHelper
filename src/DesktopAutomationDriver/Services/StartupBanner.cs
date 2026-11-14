using DesktopAutomationDriver.Services.Bootstrap;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Formats the driver startup banner. Bearer and operator tokens must never appear
/// in secure bootstrap mode. Operator token must never appear in any mode.
/// </summary>
public static class StartupBanner
{
    public static string Format(IDriverContext ctx, BootstrapOptions? bootstrap = null)
    {
        ArgumentNullException.ThrowIfNull(ctx);
        bootstrap ??= new BootstrapOptions();

        var probeInfo = DescribeProbe(ctx, bootstrap);
        var p = ctx.MainPort;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine();
        sb.AppendLine("╔══════════════════════════════════════════════════════════════╗");
        sb.AppendLine("║          Desktop Automation Driver — Started                 ║");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine($"║  Windows user  : {ctx.Username,-44}║");
        if (!string.IsNullOrWhiteSpace(ctx.UserSid))
            sb.AppendLine($"║  User SID      : {Truncate(ctx.UserSid, 44),-44}║");
        sb.AppendLine($"║  Main port     : {ctx.MainPort,-44}║");
        sb.AppendLine($"║  Bootstrap     : {DescribeMode(bootstrap),-44}║");
        if (bootstrap.IsSecure)
        {
            sb.AppendLine("║  Bearer token  : (not logged — secure SID bootstrap)         ║");
        }
        else
        {
            sb.AppendLine($"║  Bearer token  : {ctx.BearerToken,-44}║");
        }

        sb.AppendLine("║  Operator token: (not logged — confirmation UX dialog)       ║");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine("║  UNIFIED ENDPOINT (require: Authorization: Bearer <token>)   ║");
        sb.AppendLine("║                                                              ║");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/ui  <- all operations");
        sb.AppendLine("║                                                              ║");
        sb.AppendLine("║  Operations: launch, close, maximize, minimize,              ║");
        sb.AppendLine("║    switchwindow, refresh, screenshot, listelements,          ║");
        sb.AppendLine("║    listwindows, exists, waitfor, isenabled, isvisible,       ║");
        sb.AppendLine("║    isclickable, iseditable, ischecked, getvalue, gettext,    ║");
        sb.AppendLine("║    getname, getcontroltype, getselected,                     ║");
        sb.AppendLine("║    gettabledata, gettableheaders,                            ║");
        sb.AppendLine("║    isrightof, isleftof, isabove, isbelow, getposition,       ║");
        sb.AppendLine("║    click, clickmenu, clickmenupath, clicklogicalmenupath,    ║");
        sb.AppendLine("║    inspectlogicalmenu, doubleclick, rightclick, hover,       ║");
        sb.AppendLine("║    focus, type, typedate, clear, sendkeys, scroll, check,    ║");
        sb.AppendLine("║    uncheck, select, openheaderdropdown,                      ║");
        sb.AppendLine("║    selectheaderdropdownitem, alertok, alertcancel,           ║");
        sb.AppendLine("║    alertclose, popupok                                       ║");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine("║  OTHER ENDPOINTS (require: Authorization: Bearer <token>)    ║");
        sb.AppendLine("║                                                              ║");
        sb.AppendLine($"║  GET  http://127.0.0.1:{p}/status");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/session");
        sb.AppendLine($"║  GET  http://127.0.0.1:{p}/sessions");
        sb.AppendLine($"║  GET  http://127.0.0.1:{p}/session/{{id}}");
        sb.AppendLine($"║  DEL  http://127.0.0.1:{p}/session/{{id}}");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/session/{{id}}/element");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/record/start");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/playback");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine("║  SIMPLE ENDPOINTS (require: Authorization: Bearer <token>)   ║");
        sb.AppendLine("║                                                              ║");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/launch");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/close");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/listallwindows");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/switchwindow");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/maximize");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/click/name");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/click/aid");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/click/advanced");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/doubleclick/name");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/doubleclick/aid");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/select/combobox/name");
        sb.AppendLine($"║  POST http://127.0.0.1:{p}/select/combobox/aid");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine("║  DISCOVERY                                                   ║");
        sb.AppendLine($"║  {probeInfo,-62}║");
        sb.AppendLine("╚══════════════════════════════════════════════════════════════╝");
        if (bootstrap.IsSecure)
        {
            sb.AppendLine("  Secure bootstrap: retrieve the bearer token from the SID");
            sb.AppendLine("  ACL-protected named pipe / connection descriptor (never HTTP).");
        }
        else
        {
            sb.AppendLine("  WARNING: LegacyVerify returns the bearer token from HTTP /verify.");
            sb.AppendLine("  Unsafe on shared Citrix/RDS/VDI hosts.");
            sb.AppendLine("  Authorization header to use:");
            sb.AppendLine($"    Authorization: Bearer {ctx.BearerToken}");
        }

        sb.AppendLine("  Vision approve: POST /vision/approve (Bearer + operator");
        sb.AppendLine("  confirmation UX dialog). Dialog is UIA-automatable — not");
        sb.AppendLine("  OS-enforced human presence. ACL file self-read is not used.");
        return sb.ToString();
    }

    private static string DescribeMode(BootstrapOptions bootstrap) =>
        bootstrap.IsSecure
            ? "Secure (SID pipe/descriptor)"
            : "LegacyVerify (HTTP /verify — RISKY)";

    private static string DescribeProbe(IDriverContext ctx, BootstrapOptions bootstrap)
    {
        if (bootstrap.IsSecure)
        {
            return string.IsNullOrWhiteSpace(ctx.UserSid)
                ? "secure SID pipe/descriptor (SID unresolved at banner time)"
                : $"pipe {SecureBootstrapConstants.BuildPipeName(ctx.UserSid)}";
        }

        return ctx.ProbePortActive
            ? $"http://localhost:{ctx.ProbePort}/verify  (active — LEGACY)"
            : $"port {ctx.ProbePort} unavailable — use main port for /verify (LEGACY)";
    }

    private static string Truncate(string value, int max)
    {
        if (value.Length <= max)
            return value;
        return value[..(max - 3)] + "...";
    }
}
