namespace DesktopAutomationAgent.Cli;

public enum AgentCommandKind
{
    Help,
    Init,
    ValidateSuite,
    ValidateKeys,
    Doctor
}

public sealed class ParsedCommand
{
    public AgentCommandKind Kind { get; init; }

    public string? SuiteFile { get; init; }

    public IReadOnlyList<string> Keys { get; init; } = Array.Empty<string>();

    public bool Json { get; init; }

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
            _ => new ParsedCommand
            {
                Kind = AgentCommandKind.Help,
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

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateKeys,
            Keys = keys,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseDoctor(string[] rest)
    {
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--json")
            {
                json = true;
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
                Kind = AgentCommandKind.Doctor,
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
            || arg.StartsWith("Driver__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Workspace__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Suites__", StringComparison.OrdinalIgnoreCase))
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
        Desktop Automation Agent (Phase 1)

        Commands:
          init
              Create the automation workspace templates (idempotent).

          validate-suite --file <path>
              Validate one suite manifest. Does not require the driver.

          validate-keys --keys <KEY-1,KEY-2>
              Validate ad-hoc Jira keys. Does not call Jira.

          doctor [--json]
              Run configuration, workspace, driver discovery, status, and catalog checks.

        Configuration precedence:
          appsettings.json -> automation/config/agentsettings.local.json -> DA_AGENT__* env -> CLI

        Exit codes:
          0 success
          2 usage or configuration error
          3 driver unavailable or unsafe discovery
          4 authentication or catalog incompatibility
          5 suite or workspace validation failure
        """;
}
