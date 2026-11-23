using System.Net;
using System.Text.Json.Nodes;

namespace TeamsPost.Tests;

/// <summary>Made-up chats for unanswered. Env's clock is Tue 17 Nov 2026 10:30 UTC, so the default 48 hours runs
/// from Sun 15 Nov 10:30.</summary>
internal static class People
{
    public static readonly (string, string) Me = (FakeTeams.MeId, FakeTeams.MeName);
    public static readonly (string, string) Alex = ("user-alex", "Alex Example");
    public static readonly (string, string) Sam = ("user-sam", "Sam Sample");
    public static readonly (string, string) Pat = ("user-pat", "Pat Person");
    public static readonly (string, string) Bot = ("app-bot", "Workflows");

    public const string AlexChat = "19:alex_me@unq.gbl.spaces";
    public const string Project = "19:project000000000000000000000000@thread.v2";
    public const string BotChat = "19:bot00000000000000000000000000000@thread.v2";
    public const string Design = "19:design000000000000000000000000000@thread.v2";
    public const string PatChat = "19:pat_me@unq.gbl.spaces";
    public const string Sync = "19:meeting_sync000000000000000000@thread.v2";
    public const string Standup = "19:meeting_standup000000000000000@thread.v2";
    public const string Events = "19:events0000000000000000000000000@thread.v2";
    public const string Gone = "19:gone00000000000000000000000000000@thread.v2";
}

public class UnansweredTests
{
    private static JsonNode Json(Env env) => JsonNode.Parse(env.Out.ToString())!;

    private static List<string> Listed(Env env) =>
        ((JsonArray)Json(env)["chats"]!).Select(c => (string)c!["chat"]!).ToList();

    private static JsonNode Chat(Env env, string id) =>
        ((JsonArray)Json(env)["chats"]!).Single(c => (string?)c!["chat"] == id)!;

    /// <summary>Every case at once: who spoke last, bots, system events, deletions, the window, meetings.</summary>
    private static void Scene(FakeTeams teams)
    {
        // 1:1, Alex asked last: listed, direct and a question
        teams.AddChat(People.AlexChat, "oneOnOne", null, People.Alex);
        teams.Say(People.AlexChat, People.Me, "<p>morning</p>", "2026-11-16T08:00:00Z");
        teams.Say(People.AlexChat, People.Alex, "<p>Have you seen the design?</p>", "2026-11-17T08:55:00Z");
        teams.Say(People.AlexChat, People.Alex, "<p>Can you <a href=\"https://x.example/doc\">look at this</a> today?</p>", "2026-11-17T09:00:00Z");

        // group, he answered last: not listed, and not even read
        teams.AddChat(People.Project, "group", "Project chat", People.Alex, People.Sam);
        teams.Say(People.Project, People.Alex, "<p>Shall we?</p>", "2026-11-17T08:00:00Z");
        teams.Say(People.Project, People.Me, "<p>Yes</p>", "2026-11-17T08:30:00Z");

        // a bot spoke after him: not listed
        teams.AddChat(People.BotChat, "group", "Bot chat", People.Sam);
        teams.Say(People.BotChat, People.Sam, "<p>Is this right?</p>", "2026-11-17T08:00:00Z");
        teams.Say(People.BotChat, People.Me, "<p>It is</p>", "2026-11-17T08:10:00Z");
        teams.Say(People.BotChat, People.Bot, "<p>Approval needed?</p>", "2026-11-17T09:30:00Z").App = true;

        // a group where Sam @mentions him: listed, mentionsYou, not a question
        teams.AddChat(People.Design, "group", "Design review", People.Sam, People.Alex);
        teams.Say(People.Design, People.Sam, $"<p><at id=\"0\">{FakeTeams.MeName}</at> your thoughts please</p>", "2026-11-16T12:00:00Z",
            FakeTeams.MentionOf(FakeTeams.MeId, FakeTeams.MeName));

        // outside the window: not listed
        teams.AddChat(People.PatChat, "oneOnOne", null, People.Pat);
        teams.Say(People.PatChat, People.Pat, "<p>Any news?</p>", "2026-11-14T09:00:00Z");

        // meeting chats: only with a mention of him
        teams.AddChat(People.Sync, "meeting", "Weekly sync", People.Sam);
        teams.Say(People.Sync, People.Sam, "<p>Slides are up?</p>", "2026-11-17T07:00:00Z");
        teams.AddChat(People.Standup, "meeting", "Standup", People.Alex);
        teams.Say(People.Standup, People.Alex, $"<p><at id=\"0\">{FakeTeams.MeName}</at> can you cover Friday</p>", "2026-11-17T07:30:00Z",
            FakeTeams.MentionOf(FakeTeams.MeId, FakeTeams.MeName));

        // a system event last: it doesn't count, Alex's message before it does
        teams.AddChat(People.Events, "group", "Events", People.Alex);
        teams.Say(People.Events, People.Me, "<p>done</p>", "2026-11-16T09:00:00Z");
        teams.Say(People.Events, People.Alex, "<p>thanks</p>", "2026-11-17T06:00:00Z");
        teams.Say(People.Events, People.Alex, "", "2026-11-17T10:00:00Z").MessageType = "systemEventMessage";

        // someone else's message last, but deleted: his is the last that counts
        teams.AddChat(People.Gone, "group", "Gone", People.Pat);
        teams.Say(People.Gone, People.Me, "<p>see above</p>", "2026-11-17T06:00:00Z");
        teams.Say(People.Gone, People.Pat, "<p>oops?</p>", "2026-11-17T06:05:00Z").Deleted = true;
    }

    [Fact]
    public void Lists_the_chats_where_someone_else_spoke_last_newest_first_with_flags()
    {
        using var env = new Env();
        Scene(env.Teams);
        Assert.Equal(0, env.Run("unanswered", "--json"));
        Assert.Equal([People.AlexChat, People.Standup, People.Events, People.Design], Listed(env));

        var alex = Chat(env, People.AlexChat);
        Assert.Equal("Alex Example", (string?)alex["title"]);
        Assert.Equal("oneOnOne", (string?)alex["chatType"]);
        Assert.True((bool)alex["flags"]!["direct"]!);
        Assert.True((bool)alex["flags"]!["question"]!);
        Assert.False((bool)alex["flags"]!["mentionsYou"]!);
        var messages = (JsonArray)alex["messages"]!;
        Assert.Equal(2, messages.Count);   // both since his last, oldest first
        Assert.Equal("Have you seen the design?", (string?)messages[0]!["text"]);
        Assert.Equal("Can you look at this <https://x.example/doc> today?", (string?)messages[1]!["text"]);
        Assert.Equal((string?)messages[1]!["id"], (string?)alex["replyTo"]);
        Assert.Equal("Alex Example", (string?)messages[1]!["from"]);

        var design = Chat(env, People.Design);
        Assert.Equal("Design review", (string?)design["title"]);
        Assert.True((bool)design["flags"]!["mentionsYou"]!);
        Assert.False((bool)design["flags"]!["question"]!);
        Assert.False((bool)design["flags"]!["direct"]!);

        Assert.True((bool)Chat(env, People.Standup)["flags"]!["mentionsYou"]!);

        // nothing was written, and the chat he'd answered wasn't read
        Assert.Equal(0, env.Teams.Writes);
        Assert.DoesNotContain(env.Teams.Requests, r => r.Url.Contains(People.Project, StringComparison.Ordinal) && r.Url.Contains("/messages", StringComparison.Ordinal));
        Assert.Contains(["Chat.Read"], env.Teams.Scopes);
        Assert.All(env.Teams.Scopes, s => Assert.DoesNotContain("ChatMessage.Send", s));
    }

    [Fact]
    public void The_text_output_lists_the_same()
    {
        using var env = new Env();
        Scene(env.Teams);
        Assert.Equal(0, env.Run("unanswered"));
        var output = env.Out.ToString();
        Assert.Contains("(48 hours): 4 chats where someone else spoke last.", output);
        Assert.Contains("Alex Example  (one-to-one; direct, question)", output);
        Assert.Contains("Design review  (group chat; mentionsYou)", output);
        Assert.Contains($"  chat {People.AlexChat}", output);
        Assert.Contains("  reply --to ", output);
        Assert.DoesNotContain("Project chat", output);
        Assert.DoesNotContain("Weekly sync", output);
        Assert.DoesNotContain("Pat Person", output);
    }

    [Fact]
    public void Hours_on_the_command_line_and_working_days_in_the_settings_set_the_window()
    {
        using var env = new Env(Env.Settings + "\n[unanswered]\nworking_days = 2\n");
        Scene(env.Teams);
        // Tue 10:30 back two working days is Fri 13 Nov 10:30: Pat's Saturday message is in
        Assert.Equal(0, env.Run("unanswered", "--json"));
        Assert.Contains(People.PatChat, Listed(env));
        Assert.Equal("2 working days", (string?)Json(env)["window"]);

        Assert.Equal(0, env.Run("unanswered", "--json", "--hours", "3"));
        Assert.Equal([People.AlexChat, People.Standup], Listed(env));
    }

    [Fact]
    public void Working_days_skip_the_weekend()
    {
        var monday = new DateTimeOffset(2026, 11, 23, 9, 0, 0, TimeSpan.Zero);
        Assert.Equal(new DateTimeOffset(2026, 11, 19, 9, 0, 0, TimeSpan.Zero), UnansweredSettings.Since(monday, null, 2));
        Assert.Equal(monday.AddHours(-48), UnansweredSettings.Since(monday, 48, null));
        Assert.Equal(monday.AddHours(-48), UnansweredSettings.Since(monday, null, null));
    }

    [Fact]
    public void Skipped_chats_and_senders_are_left_out()
    {
        using var env = new Env(Env.Settings + $"\n[unanswered]\nskip_chats = [\"{People.AlexChat}\", \"^design\"]\nskip_senders = [\"alex example\"]\n");
        Scene(env.Teams);
        Assert.Equal(0, env.Run("unanswered", "--json"));
        // Alex's 1:1 by id, Design review by name, and Alex's messages wherever they are
        Assert.Empty(Listed(env));
    }

    [Fact]
    public void Only_the_newest_five_are_shown_trimmed_and_a_link_with_a_question_mark_isnt_a_question()
    {
        using var env = new Env();
        env.Teams.AddChat(People.Project, "group", "Project chat", People.Sam);
        for (var i = 0; i < 7; i++)
            env.Teams.Say(People.Project, People.Sam, $"<p>note {i} https://x.example/a?b=c</p>", $"2026-11-17T0{i}:00:00Z");
        env.Teams.Say(People.Project, People.Sam, "<p>" + new string('x', 900) + "</p>", "2026-11-17T08:00:00Z");
        Assert.Equal(0, env.Run("unanswered", "--json"));
        var chat = Chat(env, People.Project);
        Assert.False((bool)chat["flags"]!["question"]!);
        Assert.Equal(8, (int)chat["since"]!);
        var messages = (JsonArray)chat["messages"]!;
        Assert.Equal(5, messages.Count);
        Assert.StartsWith("note 3", (string?)messages[0]!["text"]);
        Assert.Equal(500, ((string)messages[4]!["text"]!).Length);
    }

    [Fact]
    public void It_reads_every_page_of_chats_back_to_the_window()
    {
        using var env = new Env();
        Scene(env.Teams);
        env.Teams.ChatPageSize = 2;
        Assert.Equal(0, env.Run("unanswered", "--json"));
        Assert.Equal([People.AlexChat, People.Standup, People.Events, People.Design], Listed(env));
    }

    [Theory]
    [InlineData("hours = 48\nworking_days = 2", "both hours and working_days")]
    [InlineData("days = 2", "unknown setting 'days'")]
    [InlineData("hours = 0", "hours must be a whole number")]
    [InlineData("skip_chats = \"x\"", "skip_chats must be a list of strings")]
    [InlineData("skip_chats = [\"(\"]", "isn't a chat id (19:...) or a regular expression")]
    public void Bad_settings_are_refused(string settings, string message)
    {
        using var env = new Env(Env.Settings + "\n[unanswered]\n" + settings + "\n");
        Assert.Equal(2, env.Run("unanswered"));
        Assert.Contains(message, env.Err.ToString());
    }
}

public class ReplyTests
{
    private static Env WithAlex(out FakeTeams.Message asked)
    {
        var env = new Env();
        env.Teams.AddChat(People.AlexChat, "oneOnOne", null, People.Alex);
        env.Teams.Say(People.AlexChat, People.Me, "<p>morning</p>", "2026-11-16T08:00:00Z");
        asked = env.Teams.Say(People.AlexChat, People.Alex, "<p>Can you look at the design today?</p>", "2026-11-17T09:00:00Z");
        File.WriteAllText(env.MessageFile, "---\nnote: for me\n---\nYes, I'll look this afternoon.\n- **first** the flows\n");
        return env;
    }

    private const string ReplyHtml = "<p>Yes, I'll look this afternoon.</p><ul><li><strong>first</strong> the flows</li></ul>";

    [Fact]
    public void A_preview_shows_the_chat_the_message_answered_and_the_reply_and_sends_nothing()
    {
        using var env = WithAlex(out var asked);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id));
        var output = env.Out.ToString();
        Assert.Contains($"To Alex Example (one-to-one, chat {People.AlexChat}):", output);
        Assert.Contains("Replying to Alex Example, ", output);
        Assert.Contains("\"Can you look at the design today?\"", output);
        Assert.Contains("  Yes, I'll look this afternoon.", output);
        Assert.Contains("  - first the flows", output);
        Assert.DoesNotContain("note: for me", output);
        Assert.Contains("Nothing has been sent", output);
        Assert.Equal(0, env.Teams.Writes);
        Assert.Empty(env.Records);
    }

    [Fact]
    public void Yes_sends_it_quoting_the_message_as_Teams_does_reads_back_and_records()
    {
        using var env = WithAlex(out var asked);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        var post = Assert.Single(env.Teams.Requests, r => r.Method == "POST");
        Assert.Equal($"/v1.0/chats/{People.AlexChat}/messages", post.Url);
        var body = JsonNode.Parse(post.Body!)!;
        Assert.Equal($"<attachment id=\"{asked.Id}\"></attachment>" + ReplyHtml, (string?)body["body"]!["content"]);
        var attachment = Assert.Single((JsonArray)body["attachments"]!)!;
        Assert.Equal(asked.Id, (string?)attachment["id"]);
        Assert.Equal("messageReference", (string?)attachment["contentType"]);
        var content = JsonNode.Parse((string)attachment["content"]!)!;
        Assert.Equal(asked.Id, (string?)content["messageId"]);
        Assert.Equal("Can you look at the design today?", (string?)content["messagePreview"]);
        var sender = content["messageSender"]!["user"]!;
        Assert.Equal("user-alex", (string?)sender["id"]);
        Assert.Equal("Alex Example", (string?)sender["displayName"]);
        Assert.Equal("aadUser", (string?)sender["userIdentityType"]);
        Assert.Null(sender["@odata.type"]);
        Assert.Null(body["mentions"]);

        var output = env.Out.ToString();
        Assert.Contains("Read back OK.", output);
        var record = Assert.Single(env.Records);
        Assert.Equal(Commands.ReplyTarget, record.Target);
        Assert.Equal(People.AlexChat, record.Chat);
        Assert.Equal("posted", record.Status);
        Assert.Equal(asked.Id, record.ReplyTo);
        Assert.Equal(env.Teams.Messages[^1].Id, record.Id);
        Assert.Equal(["ChatMessage.Send"], env.Teams.Scopes.Single(s => s.Contains("ChatMessage.Send")));
    }

    [Fact]
    public void Without_to_it_is_a_plain_message()
    {
        using var env = WithAlex(out _);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        var body = JsonNode.Parse(Assert.Single(env.Teams.Requests, r => r.Method == "POST").Body!)!;
        Assert.Equal(ReplyHtml, (string?)body["body"]!["content"]);
        Assert.Null(body["attachments"]);
        Assert.Null(Assert.Single(env.Records).ReplyTo);
    }

    [Fact]
    public void A_quote_missing_on_read_back_exits_non_zero()
    {
        using var env = WithAlex(out var asked);
        env.Teams.DropAttachments = true;
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Contains($"Read back DIFFERENT: the quote of message {asked.Id} isn't there", env.Out.ToString());
        Assert.Single(env.Records);   // it landed, so a second send is still refused
    }

    [Fact]
    public void A_second_reply_of_the_same_file_to_the_same_chat_is_refused_unless_again()
    {
        using var env = WithAlex(out var asked);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        Assert.Contains("was already sent to this chat", env.Out.ToString());
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile));   // the preview says so too
        Assert.Equal(1, env.Teams.Writes);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--yes", "--again"));
        Assert.Equal(2, env.Teams.Writes);
    }

    [Fact]
    public void A_lost_answer_is_not_resent_and_the_next_try_is_refused()
    {
        using var env = WithAlex(out var asked);
        env.Teams.LoseNextWrite = true;
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        var record = Assert.Single(env.Records);
        Assert.Equal("unknown", record.Status);
        Assert.Contains($"teams-post show {People.AlexChat}", env.Out.ToString());

        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Contains("no answer came back", env.Out.ToString());
        Assert.Equal(1, env.Teams.Writes);
    }

    [Fact]
    public void A_5xx_is_not_resent()
    {
        using var env = WithAlex(out _);
        env.Teams.AppliedButAnswered = HttpStatusCode.BadGateway;
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        Assert.Equal(1, env.Teams.Writes);
        Assert.Equal("unknown", Assert.Single(env.Records).Status);
    }

    [Theory]
    [InlineData("[DRAFT] Yes, I'll look.\n")]
    [InlineData("---\nstatus: draft [DRAFT]\n---\nYes.\n")]
    public void A_draft_is_refused(string text)
    {
        using var env = WithAlex(out _);
        File.WriteAllText(env.MessageFile, text);
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        Assert.Contains("still marked [DRAFT]", env.Err.ToString());
        Assert.Equal(0, env.Teams.Writes);
    }

    [Theory]
    [InlineData("19:abcdef@thread.tacv2", 2, "a channel's id")]
    [InlineData("19:abcdef@thread.skype", 2, "a channel's id")]
    [InlineData("not-a-chat", 2, "isn't a chat id")]
    [InlineData("19:unknown@thread.v2", 1, "isn't a chat you can read")]
    public void Channels_and_chats_he_isnt_in_are_refused(string chat, int code, string message)
    {
        using var env = WithAlex(out _);
        Assert.Equal(code, env.Run("reply", chat, env.MessageFile, "--yes"));
        Assert.Contains(message, env.Out.ToString() + env.Err.ToString());
        Assert.Equal(0, env.Teams.Writes);
    }

    [Fact]
    public void A_chat_without_him_in_its_members_is_refused()
    {
        using var env = WithAlex(out _);
        env.Teams.Chats[People.AlexChat].Members.RemoveAll(m => m.Id == FakeTeams.MeId);
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        Assert.Contains("you aren't a member", env.Out.ToString());
        Assert.Equal(0, env.Teams.Writes);
    }

    [Fact]
    public void A_quote_of_a_missing_deleted_or_system_message_is_refused()
    {
        using var env = WithAlex(out var asked);
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", "1700000000000", "--yes"));
        Assert.Contains("there's no message 1700000000000", env.Out.ToString());
        asked.Deleted = true;
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Contains("is deleted", env.Out.ToString());
        asked.Deleted = false;
        asked.MessageType = "systemEventMessage";
        Assert.Equal(1, env.Run("reply", People.AlexChat, env.MessageFile, "--to", asked.Id, "--yes"));
        Assert.Contains("a system event", env.Out.ToString());
        Assert.Equal(2, env.Run("reply", People.AlexChat, env.MessageFile, "--to", "abc"));
        Assert.Equal(0, env.Teams.Writes);
    }

    [Fact]
    public void A_reply_cant_be_edited_and_show_takes_a_chat_id()
    {
        using var env = WithAlex(out _);
        Assert.Equal(0, env.Run("reply", People.AlexChat, env.MessageFile, "--yes"));
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("sent as a reply", env.Out.ToString());
        Assert.Equal(0, env.Run("show", People.AlexChat, "--last", "1"));
        Assert.Contains("Me Example", env.Out.ToString());
        Assert.Contains("Yes, I'll look this afternoon.", env.Out.ToString());
        Assert.Equal(1, env.Teams.Writes);
    }

    [Fact]
    public void The_command_line_needs_a_chat_and_a_file()
    {
        using var env = new Env();
        Assert.Equal(2, env.Run("reply", People.AlexChat));
        Assert.Contains("the following arguments are required: file", env.Err.ToString());
        Assert.Equal(2, env.Run("unanswered", "--hours", "0"));
        Assert.Equal(2, env.Run("unanswered", "--to", "1"));
    }
}
