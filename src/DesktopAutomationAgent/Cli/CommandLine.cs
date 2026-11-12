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
    VerifyObjectRepository,
    ObserveApp,
    ExploreApp,
    ValidateKnowledge,
    BuildKnowledge,
    CompareKnowledge,
    SearchKnowledge,
    ApproveKnowledge,
    ValidateTestCase,
    ValidateStepRegistry,
    CompileKey,
    ValidateCompiledKey,
    SyncKey,
    RunKey,
    RunSuite
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

    public int? MaxNodes { get; init; }

    public bool IncludeBounds { get; init; }

    public bool IncludeHelpText { get; init; }

    public bool? IncludeFocusState { get; init; }

    public bool? ActionableOnly { get; init; }

    public string? OutputPath { get; init; }

    public string? PolicyFile { get; init; }

    public string? KnowledgePath { get; init; }

    public string? KnowledgeLeftPath { get; init; }

    public string? KnowledgeRightPath { get; init; }

    public string? KnowledgeQuery { get; init; }

    public string? ApplicationId { get; init; }

    public string? Reviewer { get; init; }

    public string? ReviewedInputPath { get; init; }

    public string? KnowledgeSourcePath { get; init; }

    public IReadOnlyList<string> ExplorationRuns { get; init; } = Array.Empty<string>();

    public bool SuggestPageObjects { get; init; }

    public IReadOnlyList<string> Keys { get; init; } = Array.Empty<string>();

    public string? TestCaseFile { get; init; }

    public string? StepRegistryFile { get; init; }

    public string? CompileJiraKey { get; init; }

    public string? CatalogFile { get; init; }

    public string? CompileFingerprint { get; init; }

    public bool Explain { get; init; }

    public bool Json { get; init; }

    public bool DryRun { get; init; }

    public bool FromJira { get; init; }

    public bool ContinueOnFailure { get; init; }

    public int? MaxFailures { get; init; }

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
            "observe-app" => ParseObserveApp(rest),
            "explore-app" => ParseExploreApp(rest),
            "validate-knowledge" => ParseValidateKnowledge(rest),
            "build-knowledge" => ParseBuildKnowledge(rest),
            "compare-knowledge" => ParseCompareKnowledge(rest),
            "search-knowledge" => ParseSearchKnowledge(rest),
            "approve-knowledge" => ParseApproveKnowledge(rest),
            "validate-test-case" => ParseValidateTestCase(rest),
            "validate-step-registry" => ParseValidateStepRegistry(rest),
            "compile-key" => ParseCompileKey(rest),
            "validate-compiled-key" => ParseValidateCompiledKey(rest),
            "sync-key" => ParseSyncKey(rest),
            "run-key" => ParseRunKey(rest),
            "run-suite" => ParseRunSuite(rest),
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

    private static ParsedCommand ParseObserveApp(string[] rest)
    {
        var json = HasFlag(rest, "--json");
        string? view = null;
        string? root = null;
        int? maxDepth = null;
        int? maxChildren = null;
        int? maxNodes = null;
        bool? includeOffscreen = null;
        var includeBounds = false;
        var includeHelpText = false;
        bool? includeFocusState = null;
        bool? actionableOnly = null;
        string? outputPath = null;
        var config = new List<string>();
        var unknown = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            switch (arg)
            {
                case "--view":
                    if (!TryReadFlagValue(rest, ref i, out var viewValue))
                        return Error(AgentCommandKind.ObserveApp, json, "--view requires a value.", config);
                    if (!seen.Add("view"))
                        return Error(AgentCommandKind.ObserveApp, json, "--view may not be repeated.", config);
                    view = viewValue;
                    continue;
                case "--root":
                    if (!TryReadFlagValue(rest, ref i, out var rootValue))
                        return Error(AgentCommandKind.ObserveApp, json, "--root requires a value.", config);
                    if (!seen.Add("root"))
                        return Error(AgentCommandKind.ObserveApp, json, "--root may not be repeated.", config);
                    root = rootValue;
                    continue;
                case "--max-depth":
                    if (!TryReadFlagValue(rest, ref i, out var depthRaw))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-depth requires a value.", config);
                    if (!seen.Add("max-depth"))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-depth may not be repeated.", config);
                    if (!int.TryParse(depthRaw, out var depth))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-depth requires an integer.", config);
                    maxDepth = depth;
                    continue;
                case "--max-children":
                    if (!TryReadFlagValue(rest, ref i, out var childrenRaw))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-children requires a value.", config);
                    if (!seen.Add("max-children"))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-children may not be repeated.", config);
                    if (!int.TryParse(childrenRaw, out var children))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-children requires an integer.", config);
                    maxChildren = children;
                    continue;
                case "--max-nodes":
                    if (!TryReadFlagValue(rest, ref i, out var nodesRaw))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-nodes requires a value.", config);
                    if (!seen.Add("max-nodes"))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-nodes may not be repeated.", config);
                    if (!int.TryParse(nodesRaw, out var nodes))
                        return Error(AgentCommandKind.ObserveApp, json, "--max-nodes requires an integer.", config);
                    maxNodes = nodes;
                    continue;
                case "--include-offscreen":
                    if (!seen.Add("include-offscreen"))
                        return Error(AgentCommandKind.ObserveApp, json, "--include-offscreen may not be repeated.", config);
                    includeOffscreen = true;
                    continue;
                case "--include-bounds":
                    if (!seen.Add("include-bounds"))
                        return Error(AgentCommandKind.ObserveApp, json, "--include-bounds may not be repeated.", config);
                    includeBounds = true;
                    continue;
                case "--include-help-text":
                    if (!seen.Add("include-help-text"))
                        return Error(AgentCommandKind.ObserveApp, json, "--include-help-text may not be repeated.", config);
                    includeHelpText = true;
                    continue;
                case "--no-focus-state":
                    if (!seen.Add("focus-state"))
                        return Error(AgentCommandKind.ObserveApp, json, "focus-state flags may not be repeated.", config);
                    includeFocusState = false;
                    continue;
                case "--include-structural":
                    if (!seen.Add("actionable-only"))
                        return Error(AgentCommandKind.ObserveApp, json, "actionable-only flags may not be repeated.", config);
                    actionableOnly = false;
                    continue;
                case "--output":
                    if (!TryReadFlagValue(rest, ref i, out var outputValue))
                        return Error(AgentCommandKind.ObserveApp, json, "--output requires a path.", config);
                    if (!seen.Add("output"))
                        return Error(AgentCommandKind.ObserveApp, json, "--output may not be repeated.", config);
                    outputPath = outputValue;
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
                Kind = AgentCommandKind.ObserveApp,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (!string.IsNullOrWhiteSpace(view) && view is not ("control" or "content" or "raw"))
            return Error(AgentCommandKind.ObserveApp, json, "--view must be one of: control, content, raw.", config);

        if (!string.IsNullOrWhiteSpace(root)
            && root is not ("activeWindow" or "processWindows" or "desktopChildren"))
        {
            return Error(
                AgentCommandKind.ObserveApp,
                json,
                "--root must be one of: activeWindow, processWindows, desktopChildren.",
                config);
        }

        if (maxDepth is < 0 or > 20)
            return Error(AgentCommandKind.ObserveApp, json, "--max-depth must be between 0 and 20.", config);

        if (maxChildren is < 1 or > 1000)
            return Error(AgentCommandKind.ObserveApp, json, "--max-children must be between 1 and 1000.", config);

        if (maxNodes is < 1 or > 2000)
            return Error(AgentCommandKind.ObserveApp, json, "--max-nodes must be between 1 and 2000.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ObserveApp,
            View = view,
            Root = root,
            MaxDepth = maxDepth,
            MaxChildren = maxChildren,
            MaxNodes = maxNodes,
            IncludeOffscreen = includeOffscreen,
            IncludeBounds = includeBounds,
            IncludeHelpText = includeHelpText,
            IncludeFocusState = includeFocusState,
            ActionableOnly = actionableOnly,
            OutputPath = outputPath,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateKnowledge(string[] rest)
    {
        string? path = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--path" or "--file")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ValidateKnowledge, json, "--path requires a path.", config);
                if (path is not null)
                    return Error(AgentCommandKind.ValidateKnowledge, json, "--path may not be repeated.", config);
                path = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.ValidateKnowledge,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(path))
            return Error(AgentCommandKind.ValidateKnowledge, json, "validate-knowledge requires --path <knowledge-dir>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateKnowledge,
            KnowledgePath = path,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseBuildKnowledge(string[] rest)
    {
        string? applicationId = null;
        string? output = null;
        var runs = new List<string>();
        var suggest = false;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg == "--application-id")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.BuildKnowledge, json, "--application-id requires a value.", config);
                if (applicationId is not null)
                    return Error(AgentCommandKind.BuildKnowledge, json, "--application-id may not be repeated.", config);
                applicationId = value;
                continue;
            }

            if (arg == "--exploration-run")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.BuildKnowledge, json, "--exploration-run requires a value.", config);
                runs.Add(value);
                continue;
            }

            if (arg == "--output")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.BuildKnowledge, json, "--output requires a path.", config);
                if (output is not null)
                    return Error(AgentCommandKind.BuildKnowledge, json, "--output may not be repeated.", config);
                output = value;
                continue;
            }

            if (arg == "--suggest-page-objects")
            {
                suggest = true;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.BuildKnowledge,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(applicationId))
            return Error(AgentCommandKind.BuildKnowledge, json, "build-knowledge requires --application-id <id>.", config);
        if (runs.Count == 0)
            return Error(AgentCommandKind.BuildKnowledge, json, "build-knowledge requires at least one --exploration-run.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.BuildKnowledge,
            ApplicationId = applicationId,
            ExplorationRuns = runs,
            OutputPath = output,
            SuggestPageObjects = suggest,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseCompareKnowledge(string[] rest)
    {
        string? left = null;
        string? right = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg == "--left")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompareKnowledge, json, "--left requires a path.", config);
                if (left is not null)
                    return Error(AgentCommandKind.CompareKnowledge, json, "--left may not be repeated.", config);
                left = value;
                continue;
            }

            if (arg == "--right")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompareKnowledge, json, "--right requires a path.", config);
                if (right is not null)
                    return Error(AgentCommandKind.CompareKnowledge, json, "--right may not be repeated.", config);
                right = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.CompareKnowledge,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return Error(AgentCommandKind.CompareKnowledge, json, "compare-knowledge requires --left and --right.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.CompareKnowledge,
            KnowledgeLeftPath = left,
            KnowledgeRightPath = right,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseSearchKnowledge(string[] rest)
    {
        string? path = null;
        string? query = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--path" or "--file")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.SearchKnowledge, json, "--path requires a path.", config);
                if (path is not null)
                    return Error(AgentCommandKind.SearchKnowledge, json, "--path may not be repeated.", config);
                path = value;
                continue;
            }

            if (arg is "--query" or "-q")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.SearchKnowledge, json, "--query requires a value.", config);
                if (query is not null)
                    return Error(AgentCommandKind.SearchKnowledge, json, "--query may not be repeated.", config);
                query = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.SearchKnowledge,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(path))
            return Error(AgentCommandKind.SearchKnowledge, json, "search-knowledge requires --path <knowledge-dir>.", config);
        if (string.IsNullOrWhiteSpace(query))
            return Error(AgentCommandKind.SearchKnowledge, json, "search-knowledge requires --query <text>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.SearchKnowledge,
            KnowledgePath = path,
            KnowledgeQuery = query,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseApproveKnowledge(string[] rest)
    {
        string? source = null;
        string? reviewer = null;
        string? reviewedInput = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg == "--source")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--source requires a path.", config);
                if (source is not null)
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--source may not be repeated.", config);
                source = value;
                continue;
            }

            if (arg == "--reviewer")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--reviewer requires a value.", config);
                if (reviewer is not null)
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--reviewer may not be repeated.", config);
                reviewer = value;
                continue;
            }

            if (arg == "--reviewed-input")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--reviewed-input requires a path.", config);
                if (reviewedInput is not null)
                    return Error(AgentCommandKind.ApproveKnowledge, json, "--reviewed-input may not be repeated.", config);
                reviewedInput = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.ApproveKnowledge,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(source))
            return Error(AgentCommandKind.ApproveKnowledge, json, "approve-knowledge requires --source <candidate-dir>.", config);
        if (string.IsNullOrWhiteSpace(reviewer))
            return Error(AgentCommandKind.ApproveKnowledge, json, "approve-knowledge requires --reviewer <name>.", config);
        if (string.IsNullOrWhiteSpace(reviewedInput))
            return Error(AgentCommandKind.ApproveKnowledge, json, "approve-knowledge requires --reviewed-input <file>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ApproveKnowledge,
            KnowledgeSourcePath = source,
            Reviewer = reviewer,
            ReviewedInputPath = reviewedInput,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateTestCase(string[] rest)
    {
        string? file = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ValidateTestCase, json, "--file requires a path.", config);
                if (file is not null)
                    return Error(AgentCommandKind.ValidateTestCase, json, "--file may not be repeated.", config);
                file = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.ValidateTestCase,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
            return Error(AgentCommandKind.ValidateTestCase, json, "validate-test-case requires --file <path>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateTestCase,
            TestCaseFile = file,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateStepRegistry(string[] rest)
    {
        string? file = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ValidateStepRegistry, json, "--file requires a path.", config);
                if (file is not null)
                    return Error(AgentCommandKind.ValidateStepRegistry, json, "--file may not be repeated.", config);
                file = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.ValidateStepRegistry,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateStepRegistry,
            StepRegistryFile = file,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseCompileKey(string[] rest)
    {
        string? key = null;
        string? file = null;
        string? registry = null;
        string? objectRepo = null;
        string? knowledge = null;
        string? catalog = null;
        var explain = false;
        var fromJira = false;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--key" or "-k")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--key requires a Jira key.", config);
                if (key is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--key may not be repeated.", config);
                key = value;
                continue;
            }

            if (arg is "--file" or "-f")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--file requires a path.", config);
                if (file is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--file may not be repeated.", config);
                file = value;
                continue;
            }

            if (arg == "--registry")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--registry requires a path.", config);
                if (registry is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--registry may not be repeated.", config);
                registry = value;
                continue;
            }

            if (arg is "--object-repository" or "--repository")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--object-repository requires a path.", config);
                if (objectRepo is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--object-repository may not be repeated.", config);
                objectRepo = value;
                continue;
            }

            if (arg == "--knowledge")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--knowledge requires a path.", config);
                if (knowledge is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--knowledge may not be repeated.", config);
                knowledge = value;
                continue;
            }

            if (arg == "--catalog")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.CompileKey, json, "--catalog requires a path.", config);
                if (catalog is not null)
                    return Error(AgentCommandKind.CompileKey, json, "--catalog may not be repeated.", config);
                catalog = value;
                continue;
            }

            if (arg == "--from-jira")
            {
                fromJira = true;
                continue;
            }

            if (arg == "--explain")
            {
                explain = true;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.CompileKey,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(key))
            return Error(AgentCommandKind.CompileKey, json, "compile-key requires --key <JIRA-KEY>.", config);

        if (fromJira && !string.IsNullOrWhiteSpace(file))
            return Error(AgentCommandKind.CompileKey, json, "compile-key --from-jira cannot be combined with --file.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.CompileKey,
            CompileJiraKey = key,
            TestCaseFile = file,
            StepRegistryFile = registry,
            RepositoryFile = objectRepo,
            KnowledgePath = knowledge,
            CatalogFile = catalog,
            Explain = explain,
            FromJira = fromJira,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseValidateCompiledKey(string[] rest)
    {
        string? key = null;
        string? fingerprint = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--key" or "-k")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ValidateCompiledKey, json, "--key requires a Jira key.", config);
                if (key is not null)
                    return Error(AgentCommandKind.ValidateCompiledKey, json, "--key may not be repeated.", config);
                key = value;
                continue;
            }

            if (arg == "--fingerprint")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.ValidateCompiledKey, json, "--fingerprint requires a value.", config);
                if (fingerprint is not null)
                    return Error(AgentCommandKind.ValidateCompiledKey, json, "--fingerprint may not be repeated.", config);
                fingerprint = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.ValidateCompiledKey,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(key))
            return Error(AgentCommandKind.ValidateCompiledKey, json, "validate-compiled-key requires --key <JIRA-KEY>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ValidateCompiledKey,
            CompileJiraKey = key,
            CompileFingerprint = fingerprint,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseSyncKey(string[] rest)
    {
        string? key = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--key" or "-k")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.SyncKey, json, "--key requires a Jira key.", config);
                if (key is not null)
                    return Error(AgentCommandKind.SyncKey, json, "--key may not be repeated.", config);
                key = value;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.SyncKey,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(key))
            return Error(AgentCommandKind.SyncKey, json, "sync-key requires --key <JIRA-KEY>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.SyncKey,
            CompileJiraKey = key,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseRunKey(string[] rest)
    {
        string? key = null;
        string? registry = null;
        string? objectRepo = null;
        string? knowledge = null;
        string? catalog = null;
        var dryRun = false;
        var fromJira = false;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--key" or "-k")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunKey, json, "--key requires a Jira key.", config);
                if (key is not null)
                    return Error(AgentCommandKind.RunKey, json, "--key may not be repeated.", config);
                key = value;
                continue;
            }

            if (arg == "--registry")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunKey, json, "--registry requires a path.", config);
                if (registry is not null)
                    return Error(AgentCommandKind.RunKey, json, "--registry may not be repeated.", config);
                registry = value;
                continue;
            }

            if (arg is "--object-repository" or "--repository")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunKey, json, "--object-repository requires a path.", config);
                if (objectRepo is not null)
                    return Error(AgentCommandKind.RunKey, json, "--object-repository may not be repeated.", config);
                objectRepo = value;
                continue;
            }

            if (arg == "--knowledge")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunKey, json, "--knowledge requires a path.", config);
                if (knowledge is not null)
                    return Error(AgentCommandKind.RunKey, json, "--knowledge may not be repeated.", config);
                knowledge = value;
                continue;
            }

            if (arg == "--catalog")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunKey, json, "--catalog requires a path.", config);
                if (catalog is not null)
                    return Error(AgentCommandKind.RunKey, json, "--catalog may not be repeated.", config);
                catalog = value;
                continue;
            }

            if (arg == "--from-jira")
            {
                fromJira = true;
                continue;
            }

            if (arg == "--dry-run")
            {
                dryRun = true;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.RunKey,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(key))
            return Error(AgentCommandKind.RunKey, json, "run-key requires --key <JIRA-KEY>.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.RunKey,
            CompileJiraKey = key,
            StepRegistryFile = registry,
            RepositoryFile = objectRepo,
            KnowledgePath = knowledge,
            CatalogFile = catalog,
            FromJira = fromJira,
            DryRun = dryRun,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseRunSuite(string[] rest)
    {
        string? file = null;
        string? registry = null;
        string? objectRepo = null;
        string? knowledge = null;
        string? catalog = null;
        var dryRun = false;
        var fromJira = false;
        var continueOnFailure = false;
        int? maxFailures = null;
        var json = false;
        var config = new List<string>();
        var unknown = new List<string>();

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            if (arg is "--file" or "-f")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--file requires a path.", config);
                if (file is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--file may not be repeated.", config);
                file = value;
                continue;
            }

            if (arg == "--registry")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--registry requires a path.", config);
                if (registry is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--registry may not be repeated.", config);
                registry = value;
                continue;
            }

            if (arg is "--object-repository" or "--repository")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--object-repository requires a path.", config);
                if (objectRepo is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--object-repository may not be repeated.", config);
                objectRepo = value;
                continue;
            }

            if (arg == "--knowledge")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--knowledge requires a path.", config);
                if (knowledge is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--knowledge may not be repeated.", config);
                knowledge = value;
                continue;
            }

            if (arg == "--catalog")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--catalog requires a path.", config);
                if (catalog is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--catalog may not be repeated.", config);
                catalog = value;
                continue;
            }

            if (arg == "--from-jira")
            {
                fromJira = true;
                continue;
            }

            if (arg == "--continue-on-failure")
            {
                continueOnFailure = true;
                continue;
            }

            if (arg == "--max-failures")
            {
                if (!TryReadFlagValue(rest, ref i, out var value))
                    return Error(AgentCommandKind.RunSuite, json, "--max-failures requires an integer.", config);
                if (maxFailures is not null)
                    return Error(AgentCommandKind.RunSuite, json, "--max-failures may not be repeated.", config);
                if (!int.TryParse(value, out var n) || n <= 0)
                    return Error(AgentCommandKind.RunSuite, json, "--max-failures must be a positive integer.", config);
                maxFailures = n;
                continue;
            }

            if (arg == "--dry-run")
            {
                dryRun = true;
                continue;
            }

            if (arg == "--json")
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
                Kind = AgentCommandKind.RunSuite,
                Json = json,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(file))
            return Error(AgentCommandKind.RunSuite, json, "run-suite requires --file <suite.json>.", config);

        if (maxFailures is not null && !continueOnFailure)
            return Error(AgentCommandKind.RunSuite, json, "--max-failures requires --continue-on-failure.", config);

        return new ParsedCommand
        {
            Kind = AgentCommandKind.RunSuite,
            SuiteFile = file,
            StepRegistryFile = registry,
            RepositoryFile = objectRepo,
            KnowledgePath = knowledge,
            CatalogFile = catalog,
            FromJira = fromJira,
            ContinueOnFailure = continueOnFailure,
            MaxFailures = maxFailures,
            DryRun = dryRun,
            Json = json,
            ConfigurationArgs = config.ToArray()
        };
    }

    private static ParsedCommand ParseExploreApp(string[] rest)
    {
        var json = HasFlag(rest, "--json");
        var dryRun = HasFlag(rest, "--dry-run");
        string? policy = null;
        var config = new List<string>();
        var unknown = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        for (var i = 0; i < rest.Length; i++)
        {
            var arg = rest[i];
            switch (arg)
            {
                case "--policy":
                    if (!TryReadFlagValue(rest, ref i, out var policyValue))
                        return Error(AgentCommandKind.ExploreApp, json, "--policy requires a path.", config);
                    if (!seen.Add("policy"))
                        return Error(AgentCommandKind.ExploreApp, json, "--policy may not be repeated.", config);
                    policy = policyValue;
                    continue;
                case "--json":
                case "--dry-run":
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
                Kind = AgentCommandKind.ExploreApp,
                Json = json,
                DryRun = dryRun,
                Error = $"Unexpected argument(s): {string.Join(' ', unknown)}",
                ConfigurationArgs = config.ToArray()
            };
        }

        if (string.IsNullOrWhiteSpace(policy))
        {
            return Error(AgentCommandKind.ExploreApp, json, "explore-app requires --policy <file>.", config);
        }

        return new ParsedCommand
        {
            Kind = AgentCommandKind.ExploreApp,
            PolicyFile = policy,
            Json = json,
            DryRun = dryRun,
            ConfigurationArgs = config.ToArray()
        };
    }

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

        var disallowedFlags = new List<string>();
        if (kind != AgentCommandKind.CapturePage)
        {
            if (!string.IsNullOrWhiteSpace(pageName))
                disallowedFlags.Add("--name");
        }

        if (kind is not (AgentCommandKind.CapturePage or AgentCommandKind.VerifyObjectRepository)
            && maxChildren is not null)
        {
            disallowedFlags.Add("--max-children");
        }

        if (kind == AgentCommandKind.CapturePage && !string.IsNullOrWhiteSpace(objectRef))
            disallowedFlags.Add("--ref");

        if (kind is AgentCommandKind.ValidateObjectRepository or AgentCommandKind.ResolveObject)
        {
            if (!string.IsNullOrWhiteSpace(view))
                disallowedFlags.Add("--view");
            if (!string.IsNullOrWhiteSpace(root))
                disallowedFlags.Add("--root");
            if (maxDepth is not null)
                disallowedFlags.Add("--max-depth");
            if (includeOffscreen is not null)
                disallowedFlags.Add("--include-offscreen");
            if (kind == AgentCommandKind.ValidateObjectRepository && !string.IsNullOrWhiteSpace(pageId))
                disallowedFlags.Add("--page");
            if (kind == AgentCommandKind.ValidateObjectRepository && !string.IsNullOrWhiteSpace(objectRef))
                disallowedFlags.Add("--ref");
            if (kind == AgentCommandKind.ResolveObject && !string.IsNullOrWhiteSpace(pageId))
                disallowedFlags.Add("--page");
        }

        if (disallowedFlags.Count > 0)
        {
            return Error(
                kind,
                json,
                $"Unexpected argument(s) for {CommandName(kind)}: {string.Join(' ', disallowedFlags.Distinct())}.",
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

            if (maxChildren is < 1 or > 1000)
                return Error(kind, json, "--max-children must be between 1 and 1000.", config);
        }

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
            AgentCommandKind.ObserveApp => "observe-app",
            AgentCommandKind.ExploreApp => "explore-app",
            AgentCommandKind.ValidateKnowledge => "validate-knowledge",
            AgentCommandKind.BuildKnowledge => "build-knowledge",
            AgentCommandKind.CompareKnowledge => "compare-knowledge",
            AgentCommandKind.SearchKnowledge => "search-knowledge",
            AgentCommandKind.ApproveKnowledge => "approve-knowledge",
            AgentCommandKind.ValidateTestCase => "validate-test-case",
            AgentCommandKind.ValidateStepRegistry => "validate-step-registry",
            AgentCommandKind.CompileKey => "compile-key",
            AgentCommandKind.ValidateCompiledKey => "validate-compiled-key",
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
            || arg.StartsWith("--Knowledge:", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Driver__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Workspace__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Suites__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Runner__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("ObjectRepository__", StringComparison.OrdinalIgnoreCase)
            || arg.StartsWith("Knowledge__", StringComparison.OrdinalIgnoreCase))
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
        Desktop Automation Agent (Phase 4E)

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
              [--max-depth <0-20>] [--max-children <1-1000>] [--include-offscreen] [--json]
              Verify active repository objects via finduia using the same root/depth/children options as capture.

          observe-app [--json]
              [--view control|content|raw] [--root activeWindow|processWindows|desktopChildren]
              [--max-depth N] [--max-children N] [--max-nodes N] [--include-offscreen]
              [--include-bounds] [--include-help-text] [--no-focus-state] [--include-structural]
              [--output <workspace-relative-or-safe-path>]
              Observe-only semantic snapshot via dumpuia. Never clicks, types, crawls, or plans.

          explore-app --policy <file> [--dry-run] [--json]
              Policy-bounded observe/navigate explorer. Deny-by-default. No AI, typing, or
              transactional actions. --dry-run lists permitted candidates without UI mutation.

          validate-knowledge --path <knowledge-dir> [--json]
              Offline knowledge validation. No driver or network access.

          build-knowledge --application-id <id> --exploration-run <id-or-path>
              [--exploration-run ...] [--output <dir>] [--suggest-page-objects] [--json]
              Merge exploration runs into candidate application knowledge. Offline only.
              --suggest-page-objects writes only under object-repository/candidates/.

          compare-knowledge --left <path> --right <path> [--json]
              Compare two knowledge documents. Incompatible fingerprints are reported, not merged.

          search-knowledge --path <knowledge-dir> --query <text> [--json]
              Deterministic lexical search only (no embeddings / vector DB).

          approve-knowledge --source <candidate-dir> --reviewer <name>
              --reviewed-input <approval.json> [--json]
              Promote reviewed candidate knowledge. Requires source hash + reviewer.
              Never promotes Page Objects.

          validate-test-case --file <path> [--json]
              Offline Jira test-case export validation. No Jira network access.

          validate-step-registry [--file <path>] [--json]
              Offline step-registry validation (default: step-registry/registry.json).

          compile-key --key <JIRA-KEY> [--file <test-case>] [--from-jira] [--registry <path>]
              [--object-repository <path>] [--knowledge <approved-dir>] [--catalog <path>]
              [--explain] [--json]
              Compile BDD into a deterministic Phase 2 plan cache entry.
              Without --from-jira: offline only (never calls Jira).
              With --from-jira: read-only sync into jira-cache, then compile. Never POST /ui.

          validate-compiled-key --key <JIRA-KEY> [--fingerprint <sha>] [--json]
              Validate a compiled plan.json + compilation.json cache entry offline.

          sync-key --key <JIRA-KEY> [--json]
              Read-only fetch of one Jira key into automation/jira-cache (normalized Phase 4D only).
              Rejects invalid keys before HTTP. Never writes back to Jira.

          run-key --key <JIRA-KEY> [--from-jira] [--dry-run] [--json]
              Compile (or reuse cache) then execute via the Phase 2 runner.
              --dry-run never calls POST /ui.

          run-suite --file <suite.json> [--from-jira] [--dry-run]
              [--continue-on-failure] [--max-failures N] [--json]
              Sequential suite orchestration. Fail-fast by default. Never runs disabled keys.
              Writes automation/runs/<suiteRunId>/suite-run.json (does not overwrite Phase 2 run dirs).

          doctor [--json]
              Run configuration, workspace, driver discovery, status, and catalog checks.

        Configuration precedence:
          appsettings.json -> automation/config/agentsettings.local.json -> DA_AGENT__* env -> CLI

        Exit codes:
          0 success or successful dry run
          2 usage or configuration error
          3 driver unavailable or unsafe discovery
          4 authentication or catalog incompatibility
          5 suite, plan, workspace, object repository, knowledge, BDD compile or artifact validation failure
          6 UI operation, timeout, capture/verify/observe or assertion failure
          7 execution cancelled

        Phase 4E adds a provider-neutral read-only Jira test-source and sequential suite execution.
        AI is not implemented. Scheduling, notifications, parallel runs, and Jira writeback are not implemented.
        """;
}
