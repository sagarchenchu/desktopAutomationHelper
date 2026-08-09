namespace DesktopAutomationAgent.Cli;

public enum AgentCommandKind
{
    Help,
    Init,
    ValidateSuite,
    ValidateKeys,
    Doctor,
    ValidatePlan,
    RunPlan,
    ValidateObjectRepository,
    ResolveObject,
    CapturePage,
    VerifyObjectRepository
}

public sealed class ParsedCommand
{
    public AgentCommandKind Kind { get; init; }

    public string? SuiteFile { get; init; }

    public string? PlanFile { get; init; }

    public string? RepositoryFile { get; init; }

    public string? ObjectRef { get; init; }

    public string? PageId { get; init; }

    public string? PageName { get; init; }

    public string? View { get; init; }

    public string? Root { get; init; }

    public int? MaxDepth { get; init; }

    public int? MaxChildren { get; init; }

    public bool? IncludeOffscreen { get; init; }

    public IReadOnlyList<string> Keys { get; init; } = Array.Empty<string>();

    public bool Json { get; init; }

    public bool DryRun { get; init; }

    public string? Error { get; init; }

    public string[] ConfigurationArgs { get; init; } = Array.Empty<string>();
}

public static class CommandLine
{
    public static ParsedCommand Parse(string[] args)
    {
        if (args.Length == 0 || IsHelp(args[0]))
        {
            return new ParsedCommand { Kind = AgentCommandKind.Help };
        }

        var command = args[0].Trim().ToLowerInvariant();
        var rest = args.Skip(1).ToArray();

        return command switch
        {
            "init" => ParseInit(rest),
            "validate-suite" => ParseValidateSuite(rest),
            "validate-keys" => ParseValidateKeys(rest),
            "doctor" => ParseDoctor(rest),
            "validate-plan" => ParseValidatePlan(rest),
            "run-plan" => ParseRunPlan(rest),
            "validate-object-repository" => ParseValidateObjectRepository(rest),
            "resolve-object" => ParseResolveObject(rest),
            "capture-page" => ParseCapturePage(rest),
            "verify-object-repository" => ParseVerifyObjectRepository(rest),
            _ => new ParsedCommand
            {
                Kind = AgentCommandKind.Help,
                Json = HasFlag(rest, "--json") || HasFlag(args, "--json"),
                Error = $"Unknown command '{args[0]}'."
            }
        };
    }

    private static ParsedCommand ParseInit(string[] rest)
    {
        var (configArgs, unknown) = SplitConfigArgs(rest);
        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.Init,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = configArgs
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.Init,
            ConfigurationArgs = configArgs
        };
    }

    private static ParsedCommand ParseValidateSuite(string[] rest)
    {
        string? file = null;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (i + 1 >= rest.Length)
                {
                    return new ParsedCommand
                    {
                        Kind = AgentCommandKind.ValidateSuite,
                        Error = "--file requires a path."
                    };
                }

                file = rest[++i];
                continue;
            }

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidateSuite,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidateSuite,
                Error = "validate-suite requires --file <path>.",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateSuite,
            SuiteFile = file,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateKeys(string[] rest)
    {
        string? keysArg = null;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--keys" or "-k")
            {
                if (i + 1 >= rest.Length)
                {
                    return new ParsedCommand
                    {
                        Kind = AgentCommandKind.ValidateKeys,
                        Error = "--keys requires a comma-separated list."
                    };
                }

                keysArg = rest[++i];
                continue;
            }

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidateKeys,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(keysArg))
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidateKeys,
                Error = "validate-keys requires --keys <KEY-1,KEY-2>.",
                ConfigurationArgs = config.ToArray()
            };
        }

        var keys = keysArg
            .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .ToArray();

        if (keys.Length == 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidateKeys,
                Error = "validate-keys requires at least one Jira key in --keys.",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateKeys,
            Keys = keys,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseDoctor(string[] rest)
    {
        var json = HasFlag(rest, "--json");
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--json")
                continue;

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.Doctor,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.Doctor,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidatePlan(string[] rest)
    {
        var json = HasFlag(rest, "--json");
        string? file = null;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (i + 1 >= rest.Length)
                {
                    return new ParsedCommand
                    {
                        Kind = AgentCommandKind.ValidatePlan,
                        Json = json,
                        Error = "--file requires a path.",
                        ConfigurationArgs = config.ToArray()
                    };
                }

                file = rest[++i];
                continue;
            }

            if (arg is "--json")
                continue;

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidatePlan,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.ValidatePlan,
                Json = json,
                Error = "validate-plan requires --file <path>.",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidatePlan,
            PlanFile = file,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseRunPlan(string[] rest)
    {
        var json = HasFlag(rest, "--json");
        var dryRun = HasFlag(rest, "--dry-run");
        string? file = null;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (i + 1 >= rest.Length)
                {
                    return new ParsedCommand
                    {
                        Kind = AgentCommandKind.RunPlan,
                        Json = json,
                        DryRun = dryRun,
                        Error = "--file requires a path.",
                        ConfigurationArgs = config.ToArray()
                    };
                }

                file = rest[++i];
                continue;
            }

            if (arg is "--json" or "--dry-run")
                continue;

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.RunPlan,
                Json = json,
                DryRun = dryRun,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
        {
            return new ParsedCommand
            {
                Kind = AgentCommandKind.RunPlan,
                Json = json,
                DryRun = dryRun,
                Error = "run-plan requires --file <path>.",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.RunPlan,
            PlanFile = file,
            Json = json,
            DryRun = dryRun,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateObjectRepository(string[] rest) =>
        ParseRepositoryCommand(AgentCommandKind.ValidateObjectRepository, rest, requireRef: false, requirePage: false);

    private static ParsedCommand ParseResolveObject(string[] rest) =>
        ParseRepositoryCommand(AgentCommandKind.ResolveObject, rest, requireRef: true, requirePage: false);

    private static ParsedCommand ParseCapturePage(string[] rest) =>
        ParseRepositoryCommand(AgentCommandKind.CapturePage, rest, requireRef: false, requirePage: true, requireName: true);

    private static ParsedCommand ParseVerifyObjectRepository(string[] rest) =>
        ParseRepositoryCommand(AgentCommandKind.VerifyObjectRepository, rest, requireRef: false, requirePage: false);

    private static ParsedCommand ParseRepositoryCommand(
        AgentCommandKind kind,
        string[] rest,
        bool requireRef,
        bool requirePage,
        bool requireName = false)
    {
        var json = HasFlag(rest, "--json");
        string? file = null;
        string? objectRef = null;
        string? pageId = null;
        string? pageName = null;
        string? view = null;
        string? root = null;
        int? maxDepth = null;
        int? maxChildren = null;
        bool? includeOffscreen = null;
        var config = new List<string>();
        var unknown = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            switch (arg)
            {
                case "--file" or "-f":
                    if (!TryReadFlagValue(rest, ref i, out var fileValue))
                        return Error(kind, json, "--file requires a path.", config);
                    if (!seen.Add("file"))
                        return Error(kind, json, "--file may not be repeated.", config);
                    file = fileValue;
                    continue;
                case "--ref":
                    if (!TryReadFlagValue(rest, ref i, out var refValue))
                        return Error(kind, json, "--ref requires a page.element reference.", config);
                    if (!seen.Add("ref"))
                        return Error(kind, json, "--ref may not be repeated.", config);
                    objectRef = refValue;
                    continue;
                case "--page":
                    if (!TryReadFlagValue(rest, ref i, out var pageValue))
                        return Error(kind, json, "--page requires a page id.", config);
                    if (!seen.Add("page"))
                        return Error(kind, json, "--page may not be repeated.", config);
                    pageId = pageValue;
                    continue;
                case "--name":
                    if (!TryReadFlagValue(rest, ref i, out var nameValue))
                        return Error(kind, json, "--name requires a page name.", config);
                    if (!seen.Add("name"))
                        return Error(kind, json, "--name may not be repeated.", config);
                    pageName = nameValue;
                    continue;
                case "--view":
                    if (!TryReadFlagValue(rest, ref i, out var viewValue))
                        return Error(kind, json, "--view requires a value.", config);
                    if (!seen.Add("view"))
                        return Error(kind, json, "--view may not be repeated.", config);
                    view = viewValue;
                    continue;
                case "--root":
                    if (!TryReadFlagValue(rest, ref i, out var rootValue))
                        return Error(kind, json, "--root requires a value.", config);
                    if (!seen.Add("root"))
                        return Error(kind, json, "--root may not be repeated.", config);
                    root = rootValue;
                    continue;
                case "--max-depth":
                    if (!TryReadFlagValue(rest, ref i, out var depthRaw))
                        return Error(kind, json, "--max-depth requires a value.", config);
                    if (!seen.Add("max-depth"))
                        return Error(kind, json, "--max-depth may not be repeated.", config);
                    if (!int.TryParse(depthRaw, out var depth))
                        return Error(kind, json, "--max-depth requires an integer.", config);
                    maxDepth = depth;
                    continue;
                case "--max-children":
                    if (!TryReadFlagValue(rest, ref i, out var childrenRaw))
                        return Error(kind, json, "--max-children requires a value.", config);
                    if (!seen.Add("max-children"))
                        return Error(kind, json, "--max-children may not be repeated.", config);
                    if (!int.TryParse(childrenRaw, out var children))
                        return Error(kind, json, "--max-children requires an integer.", config);
                    maxChildren = children;
                    continue;
                case "--include-offscreen":
                    if (!seen.Add("include-offscreen"))
                        return Error(kind, json, "--include-offscreen may not be repeated.", config);
                    includeOffscreen = true;
                    continue;
                case "--json":
                    continue;
            }

            if (TryTakeConfigArg(rest, ref i, config))
                continue;

            unknown.Add(arg);
        }

        if (unknown.Count > 0)
        {
            return new ParsedCommand
            {
                Kind = kind,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
        {
            return Error(kind, json, $"{CommandName(kind)} requires --file <path>.", config);
        }

        if (requireRef && string.IsNullOrWhiteSpace(objectRef))
        {
            return Error(kind, json, $"{CommandName(kind)} requires --ref <page.element>.", config);
        }

        if (requirePage && string.IsNullOrWhiteSpace(pageId))
        {
            return Error(kind, json, $"{CommandName(kind)} requires --page <page-id>.", config);
        }

        if (requireName && string.IsNullOrWhiteSpace(pageName))
        {
            return Error(kind, json, $"{CommandName(kind)} requires --name <page-name>.", config);
        }

        if (kind == AgentCommandKind.VerifyObjectRepository
            && !string.IsNullOrWhiteSpace(pageId)
            && !string.IsNullOrWhiteSpace(objectRef))
        {
            return Error(kind, json, "--page and --ref are mutually exclusive.", config);
        }

        var captureOnlyFlags = new List<string>();
        if (kind != AgentCommandKind.CapturePage)
        {
            if (!string.IsNullOrWhiteSpace(pageName))
                captureOnlyFlags.Add("--name");
            if (maxChildren is not null)
                captureOnlyFlags.Add("--max-children");
        }

        if (kind is AgentCommandKind.ValidateObjectRepository or AgentCommandKind.ResolveObject)
        {
            if (!string.IsNullOrWhiteSpace(view))
                captureOnlyFlags.Add("--view");
            if (!string.IsNullOrWhiteSpace(root))
                captureOnlyFlags.Add("--root");
            if (maxDepth is not null)
                captureOnlyFlags.Add("--max-depth");
            if (includeOffscreen is not null)
                captureOnlyFlags.Add("--include-offscreen");
            if (kind == AgentCommandKind.ValidateObjectRepository && !string.IsNullOrWhiteSpace(pageId))
                captureOnlyFlags.Add("--page");
            if (kind == AgentCommandKind.ValidateObjectRepository && !string.IsNullOrWhiteSpace(objectRef))
                captureOnlyFlags.Add("--ref");
            if (kind == AgentCommandKind.ResolveObject && !string.IsNullOrWhiteSpace(pageId))
                captureOnlyFlags.Add("--page");
        }

        if (captureOnlyFlags.Count > 0)
        {
            return Error(
                kind,
                json,
                $"Unexpected argument(s) for {CommandName(kind)}: {string.Join(' ', captureOnlyFlags.Distinct())}.",
                config);
        }

        if (kind is AgentCommandKind.CapturePage or AgentCommandKind.VerifyObjectRepository)
        {
            if (!string.IsNullOrWhiteSpace(view)
                && view is not ("control" or "content" or "raw"))
            {
                return Error(kind, json, "--view must be one of: control, content, raw.", config);
            }

            if (!string.IsNullOrWhiteSpace(root)
                && root is not ("activeWindow" or "processWindows" or "desktopChildren"))
            {
                return Error(
                    kind,
                    json,
                    "--root must be one of: activeWindow, processWindows, desktopChildren.",
                    config);
            }

            if (maxDepth is < 0 or > 20)
                return Error(kind, json, "--max-depth must be between 0 and 20.", config);
        }

        if (kind == AgentCommandKind.CapturePage && maxChildren is < 1 or > 1000)
            return Error(kind, json, "--max-children must be between 1 and 1000.", config);

        return new ParsedCommand
        {
            Kind = kind,
            RepositoryFile = file,
            ObjectRef = objectRef,
            PageId = pageId,
            PageName = pageName,
            View = view,
            Root = root,
            MaxDepth = maxDepth,
            MaxChildren = maxChildren,
            IncludeOffscreen = includeOffscreen,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand Error(
        AgentCommandKind kind,
        bool json,
        string message,
        List<string> config) =>
        new()
        {
            Kind = kind,
            Json = json,
            Error = message,
            ConfigurationArgs = config.ToArray()
        };

    private static string CommandName(AgentCommandKind kind) =>
        kind switch
        {
            AgentCommandKind.ValidateObjectRepository => "validate-object-repository",
            AgentCommandKind.ResolveObject => "resolve-object",
            AgentCommandKind.CapturePage => "capture-page",
            AgentCommandKind.VerifyObjectRepository => "verify-object-repository",
            _ => kind.ToString()
        };

    private static bool HasFlag(string[] args, string flag) =>
        args.Any(arg => string.Equals(arg, flag, StringComparison.Ordinal));

    private static bool TryReadFlagValue(string[] args, ref int index, out string value)
    {
        if (index + 1 >= args.Length)
        {
            value = string.Empty;
            return false;
        }

        var candidate = args[index + 1];
        if (candidate.StartsWith("--", StringComparison.Ordinal))
        {
            value = string.Empty;
            return false;
        }

        index++;
        value = candidate;
        return true;
    }

    private static (string[] ConfigArgs, List<string> Unknown) SplitConfigArgs(string[] rest)
    {
        var config = new List<string>();
        var unknown = new List<string>();
        for (var i = 0; i < rest.Length; i++)
        {
            if (TryTakeConfigArg(rest, ref i, config))
                continue;
            unknown.Add(rest[i]);
        }

        return (config.ToArray(), unknown);
    }

    private static bool TryTakeConfigArg(string[] args, ref int index, List<string> config)
    {
        var arg = args[index];
        if (arg.StartsWith("--Driver:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("--Workspace:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("--Suites:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("--Runner:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("--ObjectRepository:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Driver__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Workspace__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Suites__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Runner__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("ObjectRepository__", StringComparison.OrdinalIgnoreCase))
        {
            config.Add(arg);
            if (!arg.Contains('=', StringComparison.Ordinal) && index + 1 < args.Length)
            {
                config.Add(args[++index]);
            }

            return true;
        }

        return false;
    }

    private static bool IsHelp(string value) =>
        value is "-h" or "--help" or "help" or "/?";

    public static string HelpText =>
        """
        Desktop Automation Agent (Phase 3)

        Commands:
          init
              Create the automation workspace templates (idempotent).

          validate-suite --file <path>
              Validate one suite manifest. Does not require the driver.

          validate-keys --keys <KEY-1,KEY-2>
              Validate ad-hoc Jira keys. Does not call Jira.

          validate-plan --file <path> [--json]
              Offline plan validation, including optional object-repository expansion.
              Makes no HTTP calls and writes no run artifacts.

          run-plan --file <path> [--dry-run] [--json]
              Preflight against the live driver catalog, then execute the plan through POST /ui.
              --dry-run validates only (GET /status and /ui/operations); never calls POST /ui.

          validate-object-repository --file <repository.json> [--json]
              Offline object repository validation. Makes no HTTP calls.

          resolve-object --file <repository.json> --ref <page.element> [--json]
              Offline object reference resolution. Makes no HTTP calls.

          capture-page --file <repository.json> --page <page-id> --name <page-name>
              [--view control|content|raw] [--root activeWindow|processWindows|desktopChildren]
              [--max-depth N] [--max-children N] [--include-offscreen] [--json]
              Capture UI nodes via dumpuia and write capture/candidate artifacts.

          verify-object-repository --file <repository.json>
              [--page <page-id> | --ref <page.element>]
              [--view control|content|raw] [--root activeWindow|processWindows|desktopChildren]
              [--max-depth <0-20>] [--include-offscreen] [--json]
              Verify active repository objects via finduia using the same root/depth options as capture.

          doctor [--json]
              Run configuration, workspace, driver discovery, status, and catalog checks.

        Configuration precedence:
          appsettings.json -> automation/config/agentsettings.local.json -> DA_AGENT__* env -> CLI

        Exit codes:
          0 success or successful dry run
          2 usage or configuration error
          3 driver unavailable or unsafe discovery
          4 authentication or catalog incompatibility
          5 suite, plan, workspace, object repository or artifact validation failure
          6 UI operation, timeout, capture/verify or assertion failure
          7 execution cancelled

        Phase 3 adds object-repository capture, verify, resolve, and plan $objectRef expansion.
        """;
}
