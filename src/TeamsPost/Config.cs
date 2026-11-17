using System.Text;
using System.Text.RegularExpressions;
using PyCompat;
using Tomlyn;
using Tomlyn.Model;

namespace TeamsPost;

/// <summary>Settings that can't be used, or a request that can't be met: the command line reports the message and
/// exits 2.</summary>
public class UsageException(string message) : Exception(message);

/// <summary>Where a target posts. Only a group chat for now; a channel target is recognised so the settings can say
/// so clearly, and so adding channels later doesn't change the shape of the file.</summary>
public abstract record Destination;

/// <summary>A Teams chat (group or one-to-one), by its id ("19:...@thread.v2").</summary>
public sealed record ChatDestination(string ChatId) : Destination;

/// <summary>A channel in a team (planned; needs ChannelMessage.Send).</summary>
public sealed record ChannelDestination(string TeamId, string ChannelId) : Destination;

public enum Mention
{
    None,
    Everyone,
}

public sealed record Target(string Name, Destination Destination, Mention Mention)
{
    /// <summary>The chat id. Throws UsageException for a channel target, which isn't supported yet.</summary>
    public string ChatId => Destination is ChatDestination chat
        ? chat.ChatId
        : throw new UsageException($"[target.{Name}] is a channel: posting to channels isn't supported yet (only chats).");
}

/// <summary>The settings file.
///
///     login_hint = "you@example.com"
///     tenant = "organizations"            # or a tenant id
///     msal_cache = "..."                  # optional: share another tool's sign-in cache
///     state_file = "..."                  # optional; default %LOCALAPPDATA%\teams-post\posts.json
///     log_file = "..."                    # optional; default %LOCALAPPDATA%\teams-post\teams-post.log
///     draft_marker = "[DRAFT]"            # a file containing this is refused ("" for no check)
///     strip_frontmatter = true            # a leading --- YAML --- block is ignored
///
///     [target.marketing]
///     chat = "19:...@thread.v2"
///     [target.general]
///     chat = "19:...@thread.v2"
///     mention = "everyone"                # optional: the post starts with @Everyone
///
/// The file is read strictly: an unknown key, a key given twice or a value of the wrong type is an error that names
/// the setting. Paths may use ~ and %VARIABLES%; a relative path is relative to the settings file.</summary>
public sealed partial class Config
{
    public const string EnvConfig = "TEAMS_POST_CONFIG";

    // Explicit "\n" (a raw literal would take the .cs file's line endings), written with CRLF.
    public const string Template =
        "# teams-post settings. See the README.\n" +
        "# base = \"base-config.toml\"        # settings shared with other tools\n" +
        "# login_hint = \"you@example.com\"\n" +
        "# tenant = \"organizations\"          # or your organisation's tenant id\n" +
        "# msal_cache = \"...\"                # another tool's token cache, to share its sign-in\n" +
        "# state_file = \"...\"                # default %LOCALAPPDATA%\\teams-post\\posts.json\n" +
        "# log_file = \"...\"                  # default %LOCALAPPDATA%\\teams-post\\teams-post.log\n" +
        "# draft_marker = \"[DRAFT]\"          # a file containing this is refused (\"\" for no check)\n" +
        "# strip_frontmatter = true          # a leading --- YAML --- block is ignored\n" +
        "\n" +
        "# One table per place you post to. The chat id is in the chat's link in Teams (19:...@thread.v2).\n" +
        "# [target.team]\n" +
        "# chat = \"19:0123456789abcdef0123456789abcdef@thread.v2\"\n" +
        "# mention = \"everyone\"              # optional: start the post with @Everyone\n";

    /// <summary>Every top-level setting: the strict check's list, and what a base file may supply.</summary>
    public static readonly string[] KnownKeys =
        ["login_hint", "tenant", "msal_cache", "state_file", "log_file", "draft_marker", "strip_frontmatter", "target"];

    private static readonly string[] TopKeys = KnownKeys;

    private static readonly string[] TargetKeys = ["chat", "channel", "mention"];

    public required string Path { get; init; }
    public string? LoginHint { get; init; }
    public string Tenant { get; init; } = "organizations";
    public string? MsalCache { get; init; }
    public required string StateFile { get; init; }
    public required string LogFile { get; init; }
    public string DraftMarker { get; init; } = "[DRAFT]";
    public bool StripFrontmatter { get; init; } = true;
    public IReadOnlyDictionary<string, Target> Targets { get; init; } = new Dictionary<string, Target>();

    /// <summary>The target by name. Throws UsageException naming the ones there are.</summary>
    public Target TargetNamed(string name) =>
        Targets.TryGetValue(name, out var target)
            ? target
            : throw new UsageException($"No target called '{name}' in {Path} (there {(Targets.Count == 0 ? "are none" : "are: " + string.Join(", ", Targets.Keys))}).");

    /// <summary>The real environment; an empty name is no variable.</summary>
    public static string? Environment(string name) =>
        string.IsNullOrEmpty(name) || name.Contains('\0') ? null : System.Environment.GetEnvironmentVariable(name);

    /// <summary>%LOCALAPPDATA%\teams-post: the default state file and log.</summary>
    public static string Local(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var local = env("LOCALAPPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(local) ? PyPath.Home(env) : local, "teams-post"));
    }

    /// <summary>%APPDATA%\teams-post\config.toml.</summary>
    public static string DefaultPath(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var appData = env("APPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(appData) ? PyPath.Home(env) : appData, "teams-post", "config.toml"));
    }

    [GeneratedRegex(@"^[A-Za-z0-9_-]+\z", RegexOptions.CultureInvariant)]
    private static partial Regex TargetName();

    /// <summary>A chat or channel id goes into a URL path as it is, so it can't hold anything that would change the
    /// path.</summary>
    [GeneratedRegex(@"^[^\s/?#%\\]+\z", RegexOptions.CultureInvariant)]
    private static partial Regex PathSafe();

    /// <summary>The settings. Throws UsageException (a missing file included).</summary>
    public static Config Load(string? path = null, Func<string, string?>? env = null)
    {
        env ??= Environment;
        var given = !string.IsNullOrEmpty(path) ? path : env(EnvConfig) is { Length: > 0 } fromEnv ? fromEnv : DefaultPath(env);
        var full = PyPath.Resolve(given);
        if (!File.Exists(full))
            throw new UsageException($"No settings file at {full}. Run `teams-post config` to create one.");
        TomlTable raw;
        try
        {
            var text = PyStr.ReadText(full);
            if (text.StartsWith('﻿'))
                text = text[1..];
            var read = BaseSettings.Read(text, full, env, KnownKeys);
            raw = read.Raw;
            env = read.Env;   // the base's [paths], then the environment
        }
        catch (TomlException error)
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }
        catch (FormatException error)   // the base file: the message names it
        {
            throw new UsageException(error.Message.TrimEnd());
        }
        catch (Exception error) when (error is not (IOException or UnauthorizedAccessException or UsageException))
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }

        foreach (var key in raw.Keys)
            if (!TopKeys.Contains(key))
                throw new UsageException($"{full}: unknown setting '{key}' (known: {string.Join(", ", TopKeys)}).");

        var folder = System.IO.Path.GetDirectoryName(full)!;
        var local = Local(env);
        var targets = new Dictionary<string, Target>(StringComparer.Ordinal);
        if (raw.TryGetValue("target", out var t))
        {
            if (t is not TomlTable table)
                throw new UsageException($"{full}: 'target' must be tables, [target.<name>].");
            foreach (var (name, value) in table)
                targets[name] = ReadTarget(full, name, value);
        }

        return new Config
        {
            Path = full,
            LoginHint = OptionalString(raw, "login_hint", full),
            Tenant = OptionalString(raw, "tenant", full) is { Length: > 0 } tenant ? tenant : "organizations",
            MsalCache = OptionalString(raw, "msal_cache", full) is { Length: > 0 } cache ? PathSetting(cache, folder, env) : null,
            StateFile = OptionalString(raw, "state_file", full) is { Length: > 0 } state
                ? PathSetting(state, folder, env) : PyPath.Pure(System.IO.Path.Combine(local, "posts.json")),
            LogFile = OptionalString(raw, "log_file", full) is { Length: > 0 } log
                ? PathSetting(log, folder, env) : PyPath.Pure(System.IO.Path.Combine(local, "teams-post.log")),
            DraftMarker = OptionalString(raw, "draft_marker", full) ?? "[DRAFT]",
            StripFrontmatter = !raw.TryGetValue("strip_frontmatter", out var strip) || strip switch
            {
                bool b => b,
                _ => throw new UsageException($"{full}: strip_frontmatter must be true or false."),
            },
            Targets = targets,
        };
    }

    private static Target ReadTarget(string file, string name, object? value)
    {
        var where = $"{file}: [target.{name}]";
        if (!TargetName().IsMatch(name))
            throw new UsageException($"{where}: a target's name can only use letters, digits, '-' and '_'.");
        if (value is not TomlTable table)
            throw new UsageException($"{where} must be a table (chat = \"...\").");
        foreach (var key in table.Keys)
            if (!TargetKeys.Contains(key))
                throw new UsageException($"{where}: unknown setting '{key}' (known: chat, mention; channel is planned).");

        var mention = Mention.None;
        if (table.TryGetValue("mention", out var m))
        {
            mention = m switch
            {
                "everyone" => Mention.Everyone,
                "none" or "" => Mention.None,
                _ => throw new UsageException($"{where}: mention must be \"everyone\" (or left out), not {Describe(m)}."),
            };
        }

        Destination destination;
        if (table.TryGetValue("channel", out var channel))
        {
            if (table.ContainsKey("chat"))
                throw new UsageException($"{where} has both chat and channel: give one.");
            if (channel is not TomlTable c || c.Keys.Except(["team", "channel"]).Any()
                || c.TryGetValue("team", out var team) is false || team is not string { Length: > 0 } teamId
                || c.TryGetValue("channel", out var ch) is false || ch is not string { Length: > 0 } channelId)
                throw new UsageException($"{where}: channel must be {{ team = \"<team id>\", channel = \"<channel id>\" }}.");
            destination = new ChannelDestination(teamId, channelId);
        }
        else if (table.TryGetValue("chat", out var chat))
        {
            if (chat is not string { Length: > 0 } chatId)
                throw new UsageException($"{where}: chat must be the chat's id, \"19:...@thread.v2\".");
            if (!PathSafe().IsMatch(chatId))
                throw new UsageException($"{where}: chat {Describe(chatId)} doesn't look like a chat id (\"19:...@thread.v2\").");
            destination = new ChatDestination(chatId);
        }
        else
            throw new UsageException($"{where} needs a chat (chat = \"19:...@thread.v2\").");

        return new Target(name, destination, mention);
    }

    private static string? OptionalString(TomlTable table, string key, string file) =>
        !table.TryGetValue(key, out var value) ? null
        : value is string s ? s
        : throw new UsageException($"{file}: {key} must be a string, not {Describe(value)}.");

    private static string Describe(object? value) => value switch
    {
        string s => $"\"{s}\"",
        bool b => b ? "true" : "false",
        TomlTable => "a table",
        TomlArray => "an array",
        null => "nothing",
        _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "?",
    };

    private static string PathSetting(string value, string folder, Func<string, string?> env)
    {
        var expanded = PyPath.ExpandVars(PyPath.ExpandUser(value, env), env);
        return PyPath.IsAbsolute(expanded) ? PyPath.Pure(expanded) : PyPath.ResolveAgainst(folder, expanded);
    }

    /// <summary>Create the settings file if it's missing (at path, or the default place), and return where it is.</summary>
    public static string CreateDefault(string? path = null, Func<string, string?>? env = null)
    {
        var target = !string.IsNullOrEmpty(path) ? PyPath.Resolve(path) : DefaultPath(env);
        if (!File.Exists(target))
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(target)!);
            File.WriteAllText(target, PyStr.WindowsNewlines(Template), new UTF8Encoding(false));
        }
        return target;
    }
}
