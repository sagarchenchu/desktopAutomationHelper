using System.Reflection;
using DesktopAutomationDriver.Models.Operations;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Explicit immutable catalog of public /ui operations handled by <see cref="UiService.Execute"/>.
/// </summary>
public sealed class UiOperationCatalog : IUiOperationCatalog
{
    private static readonly IReadOnlyList<UiOperationDescriptor> Operations = BuildOperations();

    private static readonly HashSet<string> RecognizedNames = new(
        Operations.SelectMany(op =>
            op.Aliases.Append(op.Name).Select(n => n.ToLowerInvariant())),
        StringComparer.Ordinal);

    public UiOperationCatalogResponse GetCatalog() =>
        new()
        {
            SchemaVersion = 1,
            DriverVersion = ResolveDriverVersion(),
            Operations = GetOperations()
        };

    public IReadOnlyList<UiOperationDescriptor> GetOperations() => Operations;

    public IReadOnlyCollection<string> GetAllRecognizedNames() => RecognizedNames;

    public bool IsKnownOperation(string? operation) =>
        !string.IsNullOrWhiteSpace(operation)
        && RecognizedNames.Contains(operation.Trim().ToLowerInvariant());

    private static string ResolveDriverVersion()
    {
        var informational = Assembly
            .GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

        if (string.IsNullOrWhiteSpace(informational))
            return "0.0.0";

        var plus = informational.IndexOf('+', StringComparison.Ordinal);
        return plus >= 0 ? informational[..plus] : informational;
    }

    private static IReadOnlyList<UiOperationDescriptor> BuildOperations()
    {
        var ops = new List<UiOperationDescriptor>
        {
Op("alertcancel", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("alertclose", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("alertok", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("check", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("checkuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("clear", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("clearuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("click", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("clickgridcell", "grid-table", "action", requiresSession: true, requiredInputs: ["locator", "index", "columnIndex"], aliases: [], deprecatedAliases: []),
        Op("clickheaderdropdownitem", "grid-table", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("clicklogicalmenupath", "menu", "action", requiresSession: true, requiredInputs: ["value"], aliases: ["clickmenulogical", "menupath"], deprecatedAliases: []),
        Op("clickmenu", "menu", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("clickmenupath", "menu", "action", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("clickmenuuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("clickuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("close", "session-window", "session", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("closewindow", "session-window", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("collapsetreeitem", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("contextmenupath", "menu", "action", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("doubleclick", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("doubleclickgridcell", "grid-table", "action", requiresSession: true, requiredInputs: ["locator", "index", "columnIndex"], aliases: [], deprecatedAliases: []),
        Op("doubleclickuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("draganddrop", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["locator", "locator2"], aliases: [], deprecatedAliases: []),
        Op("dragbyoffset", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["locator", "offsetX", "offsetY"], aliases: [], deprecatedAliases: []),
        Op("dragcoordinates", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["fromX", "fromY", "toX", "toY"], aliases: [], deprecatedAliases: []),
        Op("dumpcontrols", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: [], aliases: ["printcontrolidentifiers"], deprecatedAliases: []),
        Op("dumpmenus", "menu", "diagnostic", requiresSession: true, requiredInputs: [], aliases: ["dumplogicalmenus"], deprecatedAliases: []),
        Op("dumptree", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: [], aliases: ["dump_tree", "inspecttree"], deprecatedAliases: []),
        Op("dumpuia", "native-uia", "diagnostic", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("exists", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("existsuia", "native-uia", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("expandtreeitem", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("expandtreepath", "element-action", "action", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("findall", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: [], aliases: ["findmany", "resolvemany"], deprecatedAliases: []),
        Op("findcomboboxuia", "native-uia", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("findelement", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("findelements", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("findlocator", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: ["locator"], aliases: ["inspectlocator"], deprecatedAliases: []),
        Op("finduia", "native-uia", "diagnostic", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("focus", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("focusuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getcontroltype", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getcurrentroot", "session-window", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("getgriduia", "native-uia", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getname", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getposition", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getselected", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("gettable", "grid-table", "query", requiresSession: true, requiredInputs: ["locator"], aliases: ["gettabledata"], deprecatedAliases: []),
        Op("gettableheaders", "grid-table", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("gettext", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getvalue", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("getvalueuia", "native-uia", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("hover", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("inspectcombobox", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("inspectelement", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("inspectlogicalmenu", "menu", "diagnostic", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("inspectmenupathcandidates", "menu", "diagnostic", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("isabove", "element-query", "query", requiresSession: true, requiredInputs: ["locator", "locator2"], aliases: [], deprecatedAliases: []),
        Op("isbelow", "element-query", "query", requiresSession: true, requiredInputs: ["locator", "locator2"], aliases: [], deprecatedAliases: []),
        Op("ischecked", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("isclickable", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("iseditable", "grid-table", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("isenabled", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("isfocused", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: ["hasfocus"], deprecatedAliases: []),
        Op("isleftof", "element-query", "query", requiresSession: true, requiredInputs: ["locator", "locator2"], aliases: [], deprecatedAliases: []),
        Op("isrightof", "element-query", "query", requiresSession: true, requiredInputs: ["locator", "locator2"], aliases: [], deprecatedAliases: []),
        Op("isvisible", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("iswindowactive", "element-query", "query", requiresSession: true, requiredInputs: [], aliases: ["isactive"], deprecatedAliases: []),
        Op("launch", "session-window", "session", requiresSession: false, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("listelements", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("listheaderdropdownitems", "grid-table", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("listopendropdownitems", "element-action", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("listtrackedwindows", "session-window", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("listwindows", "session-window", "query", requiresSession: false, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("maximize", "session-window", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("minimize", "session-window", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("mouse", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["action"], aliases: [], deprecatedAliases: []),
        Op("mousescroll", "mouse-keyboard", "action", requiresSession: true, requiredInputs: [], aliases: ["wheelscroll"], deprecatedAliases: []),
        Op("openheaderdropdown", "grid-table", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("popupaction", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("popupexists", "popup-alert", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("popupok", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("popuptext", "popup-alert", "query", requiresSession: true, requiredInputs: [], aliases: ["alerttext", "readpopup"], deprecatedAliases: []),
        Op("quit", "session-window", "session", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("refresh", "session-window", "session", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("resolve", "diagnostic", "diagnostic", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("rightclick", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("rightclickuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("screenshot", "session-window", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("screenshotelementuia", "native-uia", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("scroll", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("scrollintoview", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("select", "element-action", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: ["selectcomboboxitem"], deprecatedAliases: []),
        Op("selectaid", "element-action", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("selectcomboboxuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("selectdynamicmenuitem", "menu", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("selectdynamicmenupath", "menu", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("selectgridrowuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("selectheaderdropdownitem", "grid-table", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("selectopendropdownitem", "element-action", "action", requiresSession: true, requiredInputs: ["value"], aliases: ["clickopendropdownitem"], deprecatedAliases: []),
        Op("selecttabuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("selecttreeitem", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("selecttreepath", "element-action", "action", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("sendkeys", "mouse-keyboard", "action", requiresSession: true, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("sendkeysuia", "native-uia", "action", requiresSession: false, requiredInputs: ["value"], aliases: [], deprecatedAliases: []),
        Op("switchwindow", "session-window", "session", requiresSession: true, requiredInputs: ["value"], aliases: ["switch_window", "switchto", "switchwinodw"], deprecatedAliases: ["switchwinodw"]),
        Op("topwindow", "popup-alert", "query", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("type", "element-action", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("typeandselect", "element-action", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("typedate", "element-action", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("typeuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator", "value"], aliases: [], deprecatedAliases: []),
        Op("uncheck", "element-action", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("uncheckuia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("wait", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("waitfor", "element-query", "query", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        Op("waitforpopup", "popup-alert", "action", requiresSession: true, requiredInputs: [], aliases: [], deprecatedAliases: []),
        Op("waituia", "native-uia", "action", requiresSession: true, requiredInputs: ["locator"], aliases: [], deprecatedAliases: []),
        };

        return ops
            .OrderBy(o => o.Name, StringComparer.Ordinal)
            .ToArray();
    }

    private static UiOperationDescriptor Op(
        string name,
        string category,
        string operationType,
        bool requiresSession,
        string[] requiredInputs,
        string[] aliases,
        string[] deprecatedAliases,
        bool deprecated = false) =>
        new()
        {
            Name = name,
            Category = category,
            OperationType = operationType,
            RequiresSession = requiresSession,
            RequiredInputs = requiredInputs,
            Aliases = aliases,
            DeprecatedAliases = deprecatedAliases,
            Deprecated = deprecated
        };
}
