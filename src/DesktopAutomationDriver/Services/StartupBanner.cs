namespace DesktopAutomationDriver.Services;

/// <summary>
/// Formats the driver startup banner. Operator token must never appear in the output.
/// </summary>
public static class StartupBanner
{
    public static string Format(IDriverContext ctx)
    {
        ArgumentNullException.ThrowIfNull(ctx);

        var probeInfo = ctx.ProbePortActive
            ? $"http://localhost:{ctx.ProbePort}/verify  (active)"
            : $"port {ctx.ProbePort} unavailable — use main port for /verify";

        var p = ctx.MainPort;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine();
        sb.AppendLine("╔══════════════════════════════════════════════════════════════╗");
        sb.AppendLine("║          Desktop Automation Driver — Started                 ║");
        sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
        sb.AppendLine($"║  Windows user  : {ctx.Username,-44}║");
        sb.AppendLine($"║  Main port     : {ctx.MainPort,-44}║");
        sb.AppendLine($"║  Bearer token  : {ctx.BearerToken,-44}║");
        sb.AppendLine("║  Operator token: (not logged — use interactive/ACL secret)   ║");
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
        sb.AppendLine("║  PROBE (no auth)                                             ║");
        sb.AppendLine($"║  {probeInfo,-62}║");
        sb.AppendLine("╚══════════════════════════════════════════════════════════════╝");
        sb.AppendLine($"  Authorization header to use:");
        sb.AppendLine($"    Authorization: Bearer {ctx.BearerToken}");
        sb.AppendLine("  Vision operator approval: POST /vision/approve with X-Operator-Token");
        sb.AppendLine("  (operator token is never written to stdout/logs; deliver via secure UI/ACL).");
        return sb.ToString();
    }
}
