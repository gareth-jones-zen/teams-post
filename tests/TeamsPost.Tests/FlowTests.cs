using System.Net;
using System.Text.Json.Nodes;

namespace TeamsPost.Tests;

public class PostTests
{
    private const string Expected =
        "<p>From 16/11:</p><ul><li><strong>Offer</strong> - PCW promo to 30/11<ul><li>alt-net too</li></ul></li><li>Strapi go-live 24/11</li></ul>";

    [Fact]
    public void A_preview_sends_nothing()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile));
        Assert.Empty(env.Teams.Requests);
        Assert.Empty(env.Records);
        var output = env.Out.ToString();
        Assert.Contains($"To marketing (chat {Env.Marketing}):", output);
        Assert.Contains("  - Offer - PCW promo to 30/11\n    - alt-net too\n  - Strapi go-live 24/11".Replace("\n", Environment.NewLine), output);
        Assert.Contains("Nothing has been sent", output);
        Assert.DoesNotContain("left-out", output);
    }

    [Fact]
    public void Yes_posts_once_reads_back_and_records()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        var post = Assert.Single(env.Teams.Requests, r => r.Method == "POST");
        Assert.Equal($"/v1.0/chats/{Env.Marketing}/messages", post.Url);
        var body = JsonNode.Parse(post.Body!)!;
        Assert.Equal("html", (string?)body["body"]!["contentType"]);
        Assert.Equal(Expected, (string?)body["body"]!["content"]);
        Assert.Null(body["mentions"]);
        var message = Assert.Single(env.Teams.Messages);
        Assert.Contains(env.Teams.Requests, r => r.Method == "GET" && r.Url.EndsWith("/messages/" + message.Id, StringComparison.Ordinal));
        Assert.Equal(["ChatMessage.Send"], env.Teams.Scopes[0]);
        Assert.Equal(["Chat.Read"], env.Teams.Scopes[^1]);

        var record = Assert.Single(env.Records);
        Assert.Equal(env.MessageFile, record.File);
        Assert.Equal("marketing", record.Target);
        Assert.Equal(Env.Marketing, record.Chat);
        Assert.Equal(message.Id, record.Id);
        Assert.Equal("posted", record.Status);
        Assert.Equal(PostStore.Hash(Expected), record.Sha256);
        Assert.Equal("2026-11-17T10:30:00+00:00", record.PostedAt);
        Assert.Contains($"Posted {message.Id}", env.Out.ToString());
        Assert.Contains("Read back OK.", env.Out.ToString());
    }

    [Fact]
    public void Everyone_is_mentioned_and_checked()
    {
        using var env = new Env();
        File.WriteAllText(env.MessageFile, "Can you RSVP please\n- [the form](https://forms.example/x)\n");
        Assert.Equal(0, env.Run("post", "general", env.MessageFile, "--yes"));
        var body = JsonNode.Parse(Assert.Single(env.Teams.Requests, r => r.Method == "POST").Body!)!;
        Assert.StartsWith("<p><at id=\"0\">Everyone</at> Can you RSVP please</p>", (string?)body["body"]!["content"]);
        var mention = body["mentions"]![0]!;
        Assert.Equal(0, (int)mention["id"]!);
        Assert.Equal("Everyone", (string?)mention["mentionText"]);
        Assert.Equal(Env.General, (string?)mention["mentioned"]!["conversation"]!["id"]);
        Assert.Equal("Everyone", (string?)mention["mentioned"]!["conversation"]!["displayName"]);
        Assert.Equal("chat", (string?)mention["mentioned"]!["conversation"]!["conversationIdentityType"]);
        Assert.Contains("Read back OK.", env.Out.ToString());
    }

    [Fact]
    public void A_missing_mention_on_read_back_exits_non_zero()
    {
        using var env = new Env();
        env.Teams.DropMentions = true;
        Assert.Equal(1, env.Run("post", "general", env.MessageFile, "--yes"));
        Assert.Contains("Read back DIFFERENT: the @Everyone mention isn't there", env.Out.ToString());
        Assert.Single(env.Records);   // it did land, so it's recorded: a second post is still refused
    }

    [Fact]
    public void A_read_back_that_differs_exits_non_zero()
    {
        using var env = new Env();
        env.Teams.Store = html => html.Replace("<li>Strapi go-live 24/11</li>", "", StringComparison.Ordinal);
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Contains("Read back DIFFERENT: 3 list items sent, 2 in Teams", env.Out.ToString());
        Assert.Equal(1, env.Teams.Writes);
    }

    [Fact]
    public void A_message_not_found_at_first_is_waited_for()
    {
        using var env = new Env();
        env.Teams.NotFoundFirst = 2;
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal([TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(2)], env.Waits);
    }

    [Fact]
    public void A_lost_answer_is_not_resent_and_is_recorded_as_unknown()
    {
        using var env = new Env();
        env.Teams.LoseNextWrite = true;
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Single(env.Teams.Messages);   // it did land
        var record = Assert.Single(env.Records);
        Assert.Equal("unknown", record.Status);
        Assert.Null(record.Id);
        Assert.Contains("wasn't sent again", env.Out.ToString());

        // and the next post is refused until it's checked
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Contains("no answer came back", env.Out.ToString());
    }

    [Theory]
    [InlineData(HttpStatusCode.BadGateway)]
    [InlineData(HttpStatusCode.GatewayTimeout)]
    [InlineData(HttpStatusCode.InternalServerError)]
    public void A_5xx_is_not_resent(HttpStatusCode status)
    {
        using var env = new Env();
        env.Teams.AppliedButAnswered = status;
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Equal("unknown", Assert.Single(env.Records).Status);
        Assert.Contains("may have been posted, so it wasn't sent again", env.Out.ToString());

        // and the next run doesn't send it again either
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Contains("no answer came back", env.Out.ToString());
    }

    [Fact]
    public void A_timeout_after_sending_is_not_resent_and_is_recorded_as_unknown()
    {
        using var env = new Env();
        env.Teams.TimeOutNextWrite = true;
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Single(env.Teams.Messages);   // it did land
        Assert.Equal("unknown", Assert.Single(env.Records).Status);
        Assert.Contains("wasn't sent again", env.Out.ToString());
        Assert.Contains("teams-post show marketing", env.Out.ToString());
    }

    [Fact]
    public void A_5xx_on_the_read_back_is_retried()
    {
        using var env = new Env();
        env.Teams.FailNextRead = HttpStatusCode.ServiceUnavailable;
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Equal(2, env.Teams.Requests.Count(r => r.Method == "GET"));
        Assert.Contains("Read back OK.", env.Out.ToString());
    }

    [Fact]
    public void A_429_is_retried_because_it_wasnt_looked_at()
    {
        using var env = new Env();
        env.Teams.FailNextWrite = HttpStatusCode.TooManyRequests;
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(2, env.Teams.Writes);
        Assert.Single(env.Teams.Messages);
    }

    [Fact]
    public void A_refusal_records_nothing()
    {
        using var env = new Env();
        env.Teams.FailNextWrite = HttpStatusCode.Forbidden;
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Contains("Not posted: Teams refused POST", env.Out.ToString());
        Assert.Empty(env.Records);
        Assert.Empty(env.Teams.Messages);
    }

    [Fact]
    public void A_second_post_of_the_same_file_is_refused_unless_again()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Contains("was already posted to marketing", env.Out.ToString());
        Assert.Contains("teams-post edit", env.Out.ToString());
        Assert.Equal(1, env.Teams.Writes);

        // the preview says so too
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile));

        // another target is a different post
        Assert.Equal(0, env.Run("post", "general", env.MessageFile, "--yes"));
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes", "--again"));
        Assert.Equal(3, env.Teams.Writes);
        Assert.Equal(3, env.Records.Count);
    }

    [Fact]
    public void The_same_file_by_another_path_is_the_same_post()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        var other = Path.Combine(env.Folder, ".", Path.GetFileName(env.MessageFile)).ToUpperInvariant();
        Assert.Equal(1, env.Run("post", "marketing", other, "--yes"));
    }

    [Fact]
    public void A_draft_is_refused_before_anything_is_sent()
    {
        using var env = new Env();
        File.WriteAllText(env.MessageFile, "# [DRAFT] marketing\n- x\n");
        Assert.Equal(1, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Contains("still marked [DRAFT]", env.Err.ToString());
        Assert.Empty(env.Teams.Requests);
    }

    [Fact]
    public void A_missing_file_is_refused()
    {
        using var env = new Env();
        Assert.Equal(1, env.Run("post", "marketing", env.Combine("nope.md")));
        Assert.Contains("No file at", env.Err.ToString());
    }

    [Fact]
    public void An_unknown_target_is_a_usage_error()
    {
        using var env = new Env();
        Assert.Equal(2, env.Run("post", "exco", env.MessageFile));
        Assert.Contains("No target called 'exco'", env.Err.ToString());
    }

    [Fact]
    public void A_channel_target_is_refused_clearly()
    {
        using var env = new Env(Env.Settings + "\n[target.news]\nchannel = { team = \"t\", channel = \"19:c@thread.tacv2\" }\n");
        Assert.Equal(2, env.Run("post", "news", env.MessageFile, "--yes"));
        Assert.Contains("isn't supported yet", env.Err.ToString());
        Assert.Empty(env.Teams.Requests);
    }
}

public class EditTests
{
    private static Env Posted()
    {
        var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        File.WriteAllText(env.MessageFile, "From 16/11:\n- **Offer** - PCW promo to 07/12\n- Strapi go-live 24/11\n");
        return env;
    }

    [Fact]
    public void The_preview_finds_the_message_and_shows_the_change()
    {
        using var env = Posted();
        var writes = env.Teams.Writes;
        Assert.Equal(0, env.Run("edit", env.MessageFile));
        Assert.Equal(writes, env.Teams.Writes);
        var output = env.Out.ToString();
        Assert.Contains("-- Offer - PCW promo to 30/11", output);
        Assert.Contains("-  - alt-net too", output);
        Assert.Contains("+- Offer - PCW promo to 07/12", output);
        Assert.Contains("Nothing has been sent", output);
    }

    [Fact]
    public void Yes_patches_once_reads_back_and_updates_the_record()
    {
        using var env = Posted();
        var id = env.Records[0].Id!;
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        var patch = Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        Assert.Equal($"/v1.0/chats/{Env.Marketing}/messages/{id}", patch.Url);
        Assert.Equal("<p>From 16/11:</p><ul><li><strong>Offer</strong> - PCW promo to 07/12</li><li>Strapi go-live 24/11</li></ul>",
            (string?)JsonNode.Parse(patch.Body!)!["body"]!["content"]);
        Assert.Contains(["Chat.ReadWrite"], env.Teams.Scopes);
        var record = Assert.Single(env.Records);
        Assert.Equal(id, record.Id);
        Assert.Equal(PostStore.Hash((string)JsonNode.Parse(patch.Body!)!["body"]!["content"]!), record.Sha256);
        Assert.Equal("2026-11-17T10:30:00+00:00", record.EditedAt);
        Assert.Contains("Read back OK.", env.Out.ToString());

        // nothing left to change
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("nothing to change", env.Out.ToString());
        Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
    }

    [Fact]
    public void Edit_keeps_the_mention_for_an_everyone_target()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "general", env.MessageFile, "--yes"));
        File.WriteAllText(env.MessageFile, "Changed\n");
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        var body = JsonNode.Parse(Assert.Single(env.Teams.Requests, r => r.Method == "PATCH").Body!)!;
        Assert.Equal("<p><at id=\"0\">Everyone</at> Changed</p>", (string?)body["body"]!["content"]);
        Assert.Equal(Env.General, (string?)body["mentions"]![0]!["mentioned"]!["conversation"]!["id"]);
    }

    [Fact]
    public void A_lost_answer_is_not_resent()
    {
        using var env = Posted();
        env.Teams.LoseNextWrite = true;
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        Assert.Contains("wasn't sent again", env.Out.ToString());
        Assert.Contains($"teams-post show marketing {env.Records[0].Id}", env.Out.ToString());
        Assert.Null(env.Records[0].EditedAt);
    }

    [Fact]
    public void A_5xx_is_not_resent()
    {
        using var env = Posted();
        env.Teams.AppliedButAnswered = HttpStatusCode.BadGateway;
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        Assert.Contains("may have been applied, so it wasn't sent again", env.Out.ToString());
        Assert.Null(env.Records[0].EditedAt);
    }

    [Fact]
    public void A_timeout_after_sending_is_not_resent_and_the_next_run_finds_it_landed()
    {
        using var env = Posted();
        env.Teams.TimeOutNextWrite = true;
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        Assert.Contains("wasn't sent again", env.Out.ToString());

        var record = Assert.Single(env.Records);
        Assert.NotNull(record.EditUnknown);
        Assert.Null(record.EditedAt);

        // the next run reads the message first: the edit landed, so it's recorded and nothing is sent
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("landed. Nothing sent.", env.Out.ToString());
        Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        record = Assert.Single(env.Records);
        Assert.Equal(record.Sha256, PostStore.Hash(env.Teams.Messages[0].Html));
        Assert.Null(record.EditUnknown);
        Assert.NotNull(record.EditedAt);
    }

    [Fact]
    public void An_unconfirmed_edit_that_didnt_land_can_be_sent_again()
    {
        using var env = Posted();
        env.Teams.FailNextWrite = HttpStatusCode.ServiceUnavailable;   // a 5xx, and nothing applied
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("wasn't sent again", env.Out.ToString());
        Assert.NotNull(env.Records[0].EditUnknown);

        // the message still reads as before, so the preview shows the change, and --yes sends it
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Equal(2, env.Teams.Requests.Count(r => r.Method == "PATCH"));
        Assert.Null(env.Records[0].EditUnknown);
        Assert.Contains("Read back OK.", env.Out.ToString());
    }

    [Fact]
    public void A_refusal_says_what_editing_needs()
    {
        using var env = Posted();
        env.Teams.FailNextWrite = HttpStatusCode.Forbidden;
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("Chat.ReadWrite", env.Out.ToString());
    }

    [Fact]
    public void A_file_never_posted_needs_an_id()
    {
        using var env = new Env();
        Assert.Equal(1, env.Run("edit", env.MessageFile));
        Assert.Contains("No record of", env.Out.ToString());
        Assert.Empty(env.Teams.Requests);
    }

    [Fact]
    public void An_id_needs_a_target()
    {
        using var env = new Env();
        Assert.Equal(2, env.Run("edit", env.MessageFile, "--id", "123"));
        Assert.Contains("needs --target", env.Err.ToString());
    }

    [Fact]
    public void An_id_must_be_digits()
    {
        using var env = new Env();
        Assert.Equal(2, env.Run("edit", env.MessageFile, "--target", "marketing", "--id", "12/3"));
    }

    [Fact]
    public void An_id_and_target_edit_a_message_with_no_record_and_record_it()
    {
        using var env = new Env();
        var message = env.Teams.Add(Env.Marketing, "<p>old text</p>");
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--target", "marketing", "--id", message.Id, "--yes"));
        Assert.Contains("<li>Strapi go-live 24/11</li>", message.Html);
        var record = Assert.Single(env.Records);
        Assert.Equal(message.Id, record.Id);
        Assert.Equal(env.MessageFile, record.File);

        // now the file finds it by itself
        File.AppendAllText(env.MessageFile, "- one more\n");
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("<li>one more</li>", message.Html);
    }

    [Fact]
    public void A_file_posted_to_two_targets_needs_target()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("post", "marketing", env.MessageFile, "--yes"));
        Assert.Equal(0, env.Run("post", "general", env.MessageFile, "--yes"));
        Assert.Equal(2, env.Run("edit", env.MessageFile));
        Assert.Contains("say which with --target", env.Err.ToString());
        File.WriteAllText(env.MessageFile, "new\n");
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--target", "general", "--yes"));
        var patch = Assert.Single(env.Teams.Requests, r => r.Method == "PATCH");
        Assert.Contains(Env.General, patch.Url);
    }

    [Fact]
    public void An_unknown_earlier_post_isnt_guessed_at()
    {
        using var env = new Env();
        env.Teams.LoseNextWrite = true;
        env.Run("post", "marketing", env.MessageFile, "--yes");
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("no answer came back", env.Out.ToString());
        Assert.Equal(1, env.Teams.Writes);
    }

    [Fact]
    public void A_deleted_message_isnt_edited()
    {
        using var env = Posted();
        env.Teams.Messages[0].Deleted = true;
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("has been deleted", env.Out.ToString());
        Assert.DoesNotContain(env.Teams.Requests, r => r.Method == "PATCH");
    }
}

public class ShowTests
{
    [Fact]
    public void Last_n_oldest_first()
    {
        using var env = new Env();
        env.Teams.Add(Env.Marketing, "<p>first</p>");
        env.Teams.Add(Env.Marketing, "<p>second</p><ul><li>x</li></ul>");
        env.Teams.Add(Env.Marketing, "<p>third</p>");
        env.Teams.Add(Env.General, "<p>elsewhere</p>");
        Assert.Equal(0, env.Run("show", "marketing", "--last", "2"));
        var output = env.Out.ToString();
        Assert.DoesNotContain("first", output);
        Assert.DoesNotContain("elsewhere", output);
        Assert.True(output.IndexOf("second", StringComparison.Ordinal) < output.IndexOf("third", StringComparison.Ordinal));
        Assert.Contains("  - x", output);
        Assert.Contains("Test User", output);
        var list = Assert.Single(env.Teams.Requests);
        Assert.Contains("$top=2", list.Url);
        Assert.Equal(["Chat.Read"], env.Teams.Scopes.Single());
        Assert.DoesNotContain(env.Teams.Requests, r => r.Method != "GET");
    }

    [Fact]
    public void Default_is_five()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("show", "marketing"));
        Assert.Contains("$top=5", Assert.Single(env.Teams.Requests).Url);
        Assert.Contains("No messages", env.Out.ToString());
    }

    [Fact]
    public void One_message_by_id()
    {
        using var env = new Env();
        var message = env.Teams.Add(Env.General, "<p><at id=\"0\">Everyone</at> hi</p>");
        Assert.Equal(0, env.Run("show", "general", message.Id));
        Assert.Contains($"{message.Id}  ", env.Out.ToString());
        Assert.Contains("  @Everyone hi", env.Out.ToString());
    }

    [Fact]
    public void A_5xx_on_a_read_is_retried()
    {
        using var env = new Env();
        env.Teams.Add(Env.Marketing, "<p>first</p>");
        env.Teams.FailNextRead = HttpStatusCode.BadGateway;
        Assert.Equal(0, env.Run("show", "marketing"));
        Assert.Equal(2, env.Teams.Requests.Count(r => r.Method == "GET"));
        Assert.Contains("first", env.Out.ToString());
    }

    [Fact]
    public void A_missing_message_exits_non_zero()
    {
        using var env = new Env();
        Assert.Equal(1, env.Run("show", "general", "99"));
    }
}

public class CliTests
{
    [Fact]
    public void Version_and_help()
    {
        using var env = new Env();
        Assert.Equal(0, env.Run("--version"));
        Assert.Equal(Cli.Version + Environment.NewLine, env.Out.ToString());
        Assert.Equal(0, env.Run("-h"));
        Assert.Contains("teams-post post <target> <file.md>", env.Out.ToString());
        Assert.Equal(0, env.Run("post", "--help"));
    }

    [Theory]
    [InlineData(new string[0], "the following arguments are required: command")]
    [InlineData(new[] { "send" }, "invalid choice: 'send'")]
    [InlineData(new[] { "post", "marketing" }, "the following arguments are required: file")]
    [InlineData(new[] { "post" }, "the following arguments are required: target, file")]
    [InlineData(new[] { "post", "m", "f", "--force" }, "unrecognized arguments: --force")]
    [InlineData(new[] { "post", "m", "f", "extra" }, "unrecognized arguments: extra")]
    [InlineData(new[] { "edit" }, "the following arguments are required: file")]
    [InlineData(new[] { "edit", "f", "--again" }, "unrecognized arguments: --again")]
    [InlineData(new[] { "edit", "f", "--target" }, "argument --target: expected one argument")]
    [InlineData(new[] { "show" }, "the following arguments are required: target")]
    [InlineData(new[] { "show", "t", "--last", "0" }, "expected a number from 1 to 50")]
    [InlineData(new[] { "show", "t", "--last=51" }, "expected a number from 1 to 50")]
    [InlineData(new[] { "show", "t", "--yes" }, "unrecognized arguments: --yes")]
    public void Bad_arguments_exit_2(string[] args, string message)
    {
        using var env = new Env();
        Assert.Equal(2, env.Run(args));
        Assert.Contains(message, env.Err.ToString());
        Assert.Empty(env.Teams.Requests);
    }

    [Fact]
    public void Options_take_equals_too()
    {
        using var env = new Env();
        var message = env.Teams.Add(Env.Marketing, "<p>old</p>");
        Assert.Equal(0, env.Run("edit", env.MessageFile, $"--target=marketing", $"--id={message.Id}"));
        Assert.Contains("Nothing has been sent", env.Out.ToString());
    }

    [Fact]
    public void Bad_settings_exit_2()
    {
        using var env = new Env("[target.t]\nmention = \"everyone\"\n");
        Assert.Equal(2, env.Run("show", "t"));
        Assert.Contains("needs a chat", env.Err.ToString());
    }

    [Fact]
    public void Config_creates_the_file_given()
    {
        using var env = new Env();
        var path = env.Combine("new", "settings.toml");
        Assert.Equal(0, Cli.Run(["--config", path, "config"], new CliHost { Out = new StringWriter(), Env = env.Variable }));
        Assert.True(File.Exists(path));
    }

    [Fact]
    public void Sign_in_calls_the_sign_in()
    {
        using var env = new Env();
        var signedIn = false;
        var output = new StringWriter();
        Assert.Equal(0, Cli.Run(["--config", env.ConfigPath, "sign-in"], new CliHost { Out = output, Env = env.Variable, SignIn = _ => signedIn = true }));
        Assert.True(signedIn);
        Assert.Contains("Signed in.", output.ToString());
    }

    [Fact]
    public void A_sign_in_needed_exits_2()
    {
        using var env = new Env();
        var error = new StringWriter();
        var host = new CliHost
        {
            Out = new StringWriter(),
            Err = error,
            Env = env.Variable,
            Graph = _ => throw new M365Graph.SignInNeededException("At a console, run: teams-post sign-in"),
        };
        Assert.Equal(2, Cli.Run(["--config", env.ConfigPath, "show", "marketing"], host));
        Assert.Contains("teams-post sign-in", error.ToString());
    }
}
