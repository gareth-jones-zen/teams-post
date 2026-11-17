using System.Reflection;
using M365Graph;
using WinBackground;

namespace TeamsPost;

/// <summary>Everything the command line touches outside itself, so tests can give it fakes.</summary>
public sealed class CliHost
{
    public TextWriter Out { get; init; } = Console.Out;
    public TextWriter Err { get; init; } = Console.Error;
    public Func<string, string?> Env { get; init; } = Config.Environment;

    /// <summary>The Graph client for these settings.</summary>
    public Func<Config, GraphClient> Graph { get; init; } = Cli.GraphFor;

    /// <summary>Sign in, interactively.</summary>
    public Action<Config> SignIn { get; init; } = Cli.SignIn;

    public Func<DateTimeOffset> Now { get; init; } = () => DateTimeOffset.Now;
    public Action<TimeSpan> Wait { get; init; } = wait => Thread.Sleep(wait);

    /// <summary>The run lock: wait up to the seconds given; null if it wasn't free by then.</summary>
    public Func<string, double, IDisposable?> RunLock { get; init; } = (name, seconds) =>
    {
        var held = new NamedLock(name);
        return held.Acquire(seconds) ? held : null;
    };
}

/// <summary>teams-post: post a markdown file to a Teams chat, read it back, and edit it later. Exit codes: 0 done;
/// 1 refused or failed; 2 settings, arguments or sign-in.</summary>
public static class Cli
{
    public const string Prog = "teams-post";

    private const string Usage =
        "usage: teams-post [-h] [--version] [--config CONFIG] {config,sign-in,post,edit,show} ...";

    // Explicit "\n" throughout (a raw literal would take the .cs file's line endings).
    public const string Help =
        Usage + "\n\n" +
        "teams-post: post a markdown file to a Microsoft Teams chat, read it back, and edit it later.\n\n" +
        "    teams-post config                           create the settings file (if missing) and say where it is\n" +
        "    teams-post sign-in                          sign in to Microsoft (once; opens a browser)\n" +
        "    teams-post post <target> <file.md>          show what would be posted (sends nothing)\n" +
        "    teams-post post <target> <file.md> --yes    post it, then read it back\n" +
        "    teams-post edit <file.md>                   show what would change in the message it was posted as\n" +
        "    teams-post edit <file.md> --yes             edit that message, then read it back\n" +
        "    teams-post show <target> [<message id>]     a message, or the newest ones (--last N, default 5)\n\n" +
        "        --again                  post: post a file that has already been posted to that target\n" +
        "        --target T               edit: which target, when the file went to more than one\n" +
        "        --id MESSAGE_ID          edit: a message by its id (with --target), when there's no record\n\n" +
        "options:\n" +
        "  -h, --help       show this help message and exit\n" +
        "  --version        show program's version number and exit\n" +
        "  --config CONFIG  settings file (default %APPDATA%\\teams-post\\config.toml)\n";

    private static readonly string[] Commands = ["config", "sign-in", "post", "edit", "show"];

    public static string Version =>
        typeof(Cli).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";

    internal sealed class Options
    {
        public string? Config;
        public string Command = "";
        public bool Yes, Again;
        public string? Target, File, Id;
        public int Last = 5;
    }

    private sealed class ArgumentsException(string message) : Exception(message);

    public static GraphClient GraphFor(Config config) => new(new GraphClientOptions
    {
        CachePath = config.MsalCache ?? M365Graph.Graph.DefaultCachePath,
        Tenant = config.Tenant,
        LoginHint = config.LoginHint,
        Scopes = TeamsChat.SignInScopes,
        SignInHint = "At a console, run: teams-post sign-in",
        Log = Logger.Create(config.LogFile),
    });

    public static void SignIn(Config config) =>
        GraphFor(config).TokenAsync(TeamsChat.SignInScopes, true).GetAwaiter().GetResult();

    public static int Run(IReadOnlyList<string> argv, CliHost? host = null)
    {
        host ??= new CliHost();
        Options options;
        try
        {
            var parsed = Parse(argv, host);
            if (parsed is null)
                return 0;   // --help or --version
            options = parsed;
        }
        catch (ArgumentsException error)
        {
            host.Err.WriteLine(Usage);
            host.Err.WriteLine($"{Prog}: error: {error.Message}");
            return 2;
        }

        try
        {
            if (options.Command == "config")
            {
                host.Out.WriteLine(TeamsPost.Config.CreateDefault(options.Config ?? host.Env(TeamsPost.Config.EnvConfig), host.Env));
                return 0;
            }
            var config = TeamsPost.Config.Load(options.Config, host.Env);
            var log = Logger.Create(config.LogFile);
            if (options.Command == "sign-in")
            {
                host.SignIn(config);
                host.Out.WriteLine("Signed in.");
                return 0;
            }
            GraphClient? graph = null;
            var session = new Session
            {
                Config = config,
                Chat = () => new TeamsChat(graph ??= host.Graph(config)),
                Out = host.Out,
                Log = log,
                Now = host.Now,
                Wait = host.Wait,
            };
            switch (options.Command)
            {
                case "show":
                    return TeamsPost.Commands.Show(session, options.Target!, options.Id, options.Last);
                case "post" when !options.Yes:
                    return TeamsPost.Commands.Post(session, options.Target!, options.File!, false, options.Again);
                case "edit" when !options.Yes:
                    return TeamsPost.Commands.Edit(session, options.File!, false, options.Target, options.Id);
            }
            // A post or an edit for real: one at a time, so two runs can't both decide a file hasn't been posted yet.
            using (host.RunLock($"teams-post-run-{Locks.Key(config.StateFile)}", 120)
                   ?? throw new TimeoutException("Another teams-post run didn't finish within 2 minutes."))
            {
                return options.Command == "post"
                    ? TeamsPost.Commands.Post(session, options.Target!, options.File!, true, options.Again)
                    : TeamsPost.Commands.Edit(session, options.File!, true, options.Target, options.Id);
            }
        }
        catch (RefusedException error)
        {
            host.Err.WriteLine(error.Message);
            return 1;
        }
        catch (Exception error) when (error is UsageException or SignInNeededException)
        {
            host.Err.WriteLine(error.Message);
            return 2;
        }
        catch (Exception error)
        {
            host.Err.WriteLine($"{error.GetType().Name}: {error.Message}");
            return 1;
        }
    }

    private static Options? Parse(IReadOnlyList<string> argv, CliHost host)
    {
        var options = new Options();
        var i = 0;
        for (; i < argv.Count && argv[i].StartsWith('-'); i++)
        {
            var arg = argv[i];
            if (arg is "-h" or "--help")
            {
                host.Out.Write(Help);
                return null;
            }
            if (arg == "--version")
            {
                host.Out.WriteLine(Version);
                return null;
            }
            if (arg == "--config")
            {
                if (++i >= argv.Count)
                    throw new ArgumentsException("argument --config: expected one argument");
                options.Config = argv[i];
            }
            else if (arg.StartsWith("--config=", StringComparison.Ordinal))
                options.Config = arg["--config=".Length..];
            else
                throw new ArgumentsException($"unrecognized arguments: {arg}");
        }
        if (i >= argv.Count)
            throw new ArgumentsException("the following arguments are required: command");
        options.Command = argv[i++];
        if (!Commands.Contains(options.Command))
            throw new ArgumentsException($"argument command: invalid choice: '{options.Command}' (choose from {string.Join(", ", Commands.Select(c => $"'{c}'"))})");

        string Value(string name, ref int at, string arg)
        {
            if (arg.StartsWith(name + "=", StringComparison.Ordinal))
                return arg[(name.Length + 1)..];
            if (++at >= argv.Count)
                throw new ArgumentsException($"argument {name}: expected one argument");
            return argv[at];
        }

        var positional = new List<string>();
        for (; i < argv.Count; i++)
        {
            var arg = argv[i];
            var name = arg.Split('=', 2)[0];
            switch (options.Command, name)
            {
                case (_, "-h" or "--help"):
                    host.Out.Write(Help);
                    return null;
                case ("post" or "edit", "--yes"):
                    options.Yes = true;
                    break;
                case ("post", "--again"):
                    options.Again = true;
                    break;
                case ("edit", "--target"):
                    options.Target = Value(name, ref i, arg);
                    break;
                case ("edit", "--id"):
                    options.Id = Value(name, ref i, arg);
                    break;
                case ("show", "--last"):
                    var text = Value(name, ref i, arg);
                    if (!int.TryParse(text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var last)
                        || last is < 1 or > 50)
                        throw new ArgumentsException($"argument --last: expected a number from 1 to 50, not '{text}'");
                    options.Last = last;
                    break;
                default:
                    if (arg.StartsWith('-') && arg != "-")
                        throw new ArgumentsException($"unrecognized arguments: {arg}");
                    positional.Add(arg);
                    break;
            }
        }

        string Take(string what)
        {
            if (positional.Count == 0)
                throw new ArgumentsException($"the following arguments are required: {what}");
            var value = positional[0];
            positional.RemoveAt(0);
            return value;
        }

        switch (options.Command)
        {
            case "post":
                options.Target = Take("target, file");
                options.File = Take("file");
                break;
            case "edit":
                options.File = Take("file");
                break;
            case "show":
                options.Target = Take("target");
                if (positional.Count > 0)
                    options.Id = Take("id");
                break;
        }
        if (positional.Count > 0)
            throw new ArgumentsException($"unrecognized arguments: {string.Join(" ", positional)}");
        return options;
    }
}
