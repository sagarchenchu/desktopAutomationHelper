namespace DesktopAutomationAgent.Cli;

public enum AgentCommandKind
{
    Help,
    Init,
    ValidateSuite,
    ValidateKeys,
    Doctor,
    ValidatePlan,
    RunPlan
}

public sealed class ParsedCommand
{
    public AgentCommandKind Kind { get; init; }

    public string? SuiteFile { get; init; }

    public string? PlanFile { get; init; }

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
        // Scan flags before detailed parsing so early error returns still preserve --json.
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
        // Scan flags before detailed parsing so early error returns still preserve --json.
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

    private static bool HasFlag(string[] args, string flag) =>
        args.Any(arg => string.Equals(arg, flag, StringComparison.Ordinal));

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
            || arg.StartsWith("Driver__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Workspace__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Suites__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Runner__", StringComparison.OrdinalIgnoreCase))
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
        Desktop Automation Agent (Phase 2)

        Commands:
          init
              Create the automation workspace templates (idempotent).

          validate-suite --file <path>
              Validate one suite manifest. Does not require the driver.

          validate-keys --keys <KEY-1,KEY-2>
              Validate ad-hoc Jira keys. Does not call Jira.

          validate-plan --file <path> [--json]
              Offline plan validation. Makes no HTTP calls and writes no run artifacts.

          run-plan --file <path> [--dry-run] [--json]
              Preflight against the live driver catalog, then execute the plan through POST /ui.
              --dry-run validates only (GET /status and /ui/operations); never calls POST /ui.

          doctor [--json]
              Run configuration, workspace, driver discovery, status, and catalog checks.

        Configuration precedence:
          appsettings.json -> automation/config/agentsettings.local.json -> DA_AGENT__* env -> CLI

        Exit codes:
          0 success or successful dry run
          2 usage or configuration error
          3 driver unavailable or unsafe discovery
          4 authentication or catalog incompatibility
          5 suite, plan, workspace or artifact validation failure
          6 UI operation, timeout or assertion failure
          7 execution cancelled

        Phase 2 performs no Jira, BDD, AI, object-repository, database, scheduling or suite orchestration work.
        """;
}
