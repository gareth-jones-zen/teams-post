namespace TeamsPost.Tests;

public class SettingsTests
{
    private static Config Load(string toml)
    {
        using var env = new Env(toml);
        return Config.Load(env.ConfigPath, env.Variable);
    }

    private static string Refused(string toml) => Assert.Throws<UsageException>(() => Load(toml)).Message;

    [Fact]
    public void The_example_loads()
    {
        using var env = new Env();
        var config = Config.Load(env.ConfigPath, env.Variable);
        Assert.Equal("me@example.com", config.LoginHint);
        Assert.Equal("organizations", config.Tenant);
        Assert.Equal(env.Combine("posts.json"), config.StateFile);
        Assert.Equal("[DRAFT]", config.DraftMarker);
        Assert.True(config.StripFrontmatter);
        Assert.Equal(Env.Marketing, config.TargetNamed("marketing").ChatId);
        Assert.Equal(Mention.None, config.TargetNamed("marketing").Mention);
        Assert.Equal(Mention.Everyone, config.TargetNamed("general").Mention);
    }

    [Fact]
    public void Defaults_go_under_local_app_data()
    {
        using var env = new Env("[target.t]\nchat = \"19:x@thread.v2\"\n");
        var config = Config.Load(env.ConfigPath, env.Variable);
        Assert.Equal(env.Combine("local", "teams-post", "posts.json"), config.StateFile);
        Assert.Equal(env.Combine("local", "teams-post", "teams-post.log"), config.LogFile);
        Assert.Null(config.MsalCache);
    }

    [Fact]
    public void An_unknown_top_level_key_is_named() =>
        Assert.Contains("unknown setting 'draftmarker'", Refused("draftmarker = \"x\"\n"));

    [Fact]
    public void An_unknown_key_in_a_target_is_named() =>
        Assert.Contains("[target.t]: unknown setting 'chatid'", Refused("[target.t]\nchatid = \"19:x@thread.v2\"\n"));

    [Fact]
    public void A_target_needs_a_chat() =>
        Assert.Contains("[target.t] needs a chat", Refused("[target.t]\nmention = \"everyone\"\n"));

    [Fact]
    public void A_chat_must_be_a_string_that_fits_a_url()
    {
        Assert.Contains("chat must be", Refused("[target.t]\nchat = 5\n"));
        Assert.Contains("doesn't look like a chat id", Refused("[target.t]\nchat = \"19:x/../me@thread.v2\"\n"));
    }

    [Fact]
    public void A_bad_mention_is_named() =>
        Assert.Contains("mention must be \"everyone\"", Refused("[target.t]\nchat = \"19:x@thread.v2\"\nmention = \"all\"\n"));

    [Fact]
    public void A_wrong_type_is_named()
    {
        Assert.Contains("strip_frontmatter must be true or false", Refused("strip_frontmatter = \"yes\"\n"));
        Assert.Contains("login_hint must be a string", Refused("login_hint = 3\n"));
    }

    [Fact]
    public void A_key_given_twice_is_refused() =>
        Assert.Throws<UsageException>(() => Load("tenant = \"a\"\ntenant = \"b\"\n"));

    [Fact]
    public void A_channel_target_loads_but_isnt_supported_yet()
    {
        var config = Load("[target.news]\nchannel = { team = \"t-id\", channel = \"19:c@thread.tacv2\" }\n");
        var target = config.TargetNamed("news");
        Assert.Equal(new ChannelDestination("t-id", "19:c@thread.tacv2"), target.Destination);
        Assert.Contains("isn't supported yet", Assert.Throws<UsageException>(() => target.ChatId).Message);
    }

    [Fact]
    public void A_channel_target_needs_both_ids_and_not_a_chat_as_well()
    {
        Assert.Contains("channel must be", Refused("[target.n]\nchannel = { team = \"t\" }\n"));
        Assert.Contains("both chat and channel", Refused("[target.n]\nchat = \"19:x@thread.v2\"\nchannel = { team = \"t\", channel = \"c\" }\n"));
    }

    [Fact]
    public void An_unknown_target_lists_the_known_ones()
    {
        using var env = new Env();
        var config = Config.Load(env.ConfigPath, env.Variable);
        Assert.Contains("are: marketing, general", Assert.Throws<UsageException>(() => config.TargetNamed("exco")).Message);
    }

    [Fact]
    public void Missing_settings_say_how_to_make_them()
    {
        using var env = new Env();
        Assert.Contains("teams-post config", Assert.Throws<UsageException>(() => Config.Load(env.Combine("none.toml"), env.Variable)).Message);
    }

    [Fact]
    public void Config_writes_a_commented_template_that_loads()
    {
        using var env = new Env();
        var path = Config.CreateDefault(null, env.Variable);
        Assert.Equal(env.Combine("appdata", "teams-post", "config.toml"), path);
        Assert.Contains("[target.team]", File.ReadAllText(path));
        Assert.Empty(Config.Load(path, env.Variable).Targets);
    }
}
