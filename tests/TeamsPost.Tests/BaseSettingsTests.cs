namespace TeamsPost.Tests;

/// <summary>Settings that name a shared base file (base = "base-config.toml").</summary>
public sealed class BaseSettingsTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "teams-post-base-" + Guid.NewGuid().ToString("N"));

    private const string Base = """
        login_hint = "me@example.com"
        tenant = "contoso"
        msal_cache = "%EA_LOCAL%/msal-cache.json"
        time_zone = "GMT Standard Time"

        [paths]
        EA_LOCAL = "C:/Users/me/AppData/Local/claude-ea"

        [target.base]
        chat = "19:base@thread.v2"

        [git]
        push = true
        """;

    public BaseSettingsTests()
    {
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "base-config.toml"), Base);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(dir, true);
        }
        catch (IOException)
        {
        }
    }

    private static string? Env(string name) => name == "LOCALAPPDATA" ? @"C:\Users\me\AppData\Local" : null;

    private Config Load(string text)
    {
        var path = Path.Combine(dir, "teams-post.toml");
        File.WriteAllText(path, text);
        return Config.Load(path, Env);
    }

    [Fact]
    public void Shared_sign_in_and_paths_come_from_the_base()
    {
        var config = Load("base = \"base-config.toml\"\nstate_file = \"%EA_LOCAL%/teams-post.json\"\n");
        Assert.Equal("me@example.com", config.LoginHint);
        Assert.Equal("contoso", config.Tenant);
        Assert.Equal(@"C:\Users\me\AppData\Local\claude-ea\msal-cache.json", config.MsalCache);
        Assert.Equal(@"C:\Users\me\AppData\Local\claude-ea\teams-post.json", config.StateFile);
        Assert.Equal(@"C:\Users\me\AppData\Local\teams-post\teams-post.log", config.LogFile);   // time_zone, [git]: not ours, left out
    }

    [Fact]
    public void The_tools_own_settings_win()
    {
        var config = Load("base = \"base-config.toml\"\nlogin_hint = \"own@example.com\"\ntenant = \"own\"\n");
        Assert.Equal("own@example.com", config.LoginHint);
        Assert.Equal("own", config.Tenant);
    }

    [Fact]
    public void Targets_apply_only_where_the_tool_has_the_section()
    {
        Assert.Empty(Load("base = \"base-config.toml\"\n").Targets);
        var config = Load("base = \"base-config.toml\"\n[target.mine]\nchat = \"19:mine@thread.v2\"\n");
        Assert.Equal("19:mine@thread.v2", config.TargetNamed("mine").ChatId);
        Assert.Equal("19:base@thread.v2", config.TargetNamed("base").ChatId);
    }

    [Fact]
    public void A_missing_base_is_a_settings_error()
    {
        var error = Assert.Throws<UsageException>(() => Load("base = \"nope.toml\"\n"));
        Assert.Contains("doesn't exist", error.Message);
        Assert.Contains("nope.toml", error.Message);
    }
}
