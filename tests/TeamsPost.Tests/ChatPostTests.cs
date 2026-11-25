using System.Text.Json.Nodes;

namespace TeamsPost.Tests;

/// <summary>post --chat: a file posted into any chat by its id, and --attach: a file uploaded to the sender's OneDrive,
/// shared with the chat's members by name, and attached. Everything against the fake Teams and OneDrive.</summary>
public class ChatPostTests
{
    private const string Chat = "19:meeting_review0000000000000000000@thread.v2";

    private static Env Scene()
    {
        var env = new Env();
        env.Teams.AddChat(Chat, "meeting", "Design review", People.Alex, People.Sam);
        File.WriteAllText(env.MessageFile, "Transcript of Design review, 17 Nov 2026: speech-to-text, so names and terms may be wrong\n");
        File.WriteAllText(env.Combine("transcript.md"), "**Alex Example** 0:01\nShall we start?\n");
        return env;
    }

    private static JsonNode Posted(Env env) => JsonNode.Parse(env.Teams.Requests.Last(r => r.Method == "POST" && r.Url.Contains("/messages")).Body!)!;

    [Fact]
    public void A_preview_sends_nothing()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile));
        Assert.Equal(0, env.Teams.Writes);
        Assert.Empty(env.Records);
        Assert.Contains($"To Design review (meeting chat, chat {Chat}):", env.Out.ToString());
        Assert.Contains("Nothing has been sent", env.Out.ToString());
    }

    [Fact]
    public void Yes_posts_without_quoting_once_and_reads_back()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--yes"));
        var body = Posted(env);
        Assert.Equal("<p>Transcript of Design review, 17 Nov 2026: speech-to-text, so names and terms may be wrong</p>", (string?)body["body"]!["content"]);
        Assert.Null(body["attachments"]);
        Assert.Contains("Read back OK.", env.Out.ToString());
        var record = Assert.Single(env.Records);
        Assert.Equal((Commands.ChatTarget, Chat, "posted", (string?)null), (record.Target, record.Chat, record.Status, record.Attachment));

        // never twice: the same file to the same chat is refused, whether by post or reply, unless --again
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--yes"));
        Assert.Contains("was already sent to this chat", env.Out.ToString());
        Assert.Equal(1, env.Run("reply", Chat, env.MessageFile, "--yes"));
        Assert.Single(env.Teams.Messages);
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--yes", "--again"));
        Assert.Equal(2, env.Teams.Messages.Count);
    }

    [Fact]
    public void A_post_by_chat_id_can_be_edited_from_its_file()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--yes"));
        File.WriteAllText(env.MessageFile, "Corrected line\n");
        Assert.Equal(0, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Equal("<p>Corrected line</p>", Assert.Single(env.Teams.Messages).Html);
    }

    [Fact]
    public void Attach_previews_who_it_would_be_shared_with_and_uploads_nothing()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--name", "Design review 2026-11-17 transcript.md"));
        Assert.Equal(0, env.Teams.Writes);
        Assert.Empty(env.Teams.Files);
        Assert.DoesNotContain(env.Teams.Scopes, s => s.Contains("Files.ReadWrite"));
        Assert.Contains("Attached: Design review 2026-11-17 transcript.md (", env.Out.ToString());
        Assert.Contains("read-only, with 2 by name: Alex Example, Sam Sample.", env.Out.ToString());
    }

    [Fact]
    public void Attach_uploads_shares_with_the_members_by_name_and_posts_a_reference()
    {
        using var env = Scene();
        var name = "Design review 2026-11-17 transcript.md";
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--name", name, "--yes"));

        // the order: upload, read the item, share, then post
        var writes = env.Teams.Requests.Where(r => r.Method is "PUT" or "POST").Select(r => (r.Method, r.Url)).ToList();
        Assert.Equal([
            ("PUT", $"/v1.0/me/drive/root:/Microsoft Teams Chat Files/{name}:/content"),
            ("POST", "/v1.0/me/drive/items/item-1/invite"),
            ("POST", $"/v1.0/chats/{Chat}/messages"),
        ], writes);
        Assert.Equal("**Alex Example** 0:01\nShall we start?\n", System.Text.Encoding.UTF8.GetString(env.Teams.Files["item-1"].Content));

        // shared with each member but the sender, by name, read-only, sign-in required, no email, no link
        var (item, invite) = Assert.Single(env.Teams.Invites);
        Assert.Equal("item-1", item);
        Assert.Equal(["alex.example@example.com", "sam.sample@example.com"],
            invite["recipients"]!.AsArray().Select(r => (string?)r!["email"]).ToList());
        Assert.Equal((true, false, "read"), ((bool)invite["requireSignIn"]!, (bool)invite["sendInvitation"]!, (string?)invite["roles"]![0]));
        Assert.Null(invite["type"]);
        Assert.Null(invite["scope"]);
        Assert.DoesNotContain(env.Teams.Requests, r => r.Url.Contains("createLink", StringComparison.OrdinalIgnoreCase));

        // the message: the attachment tag at the end of the body, and a reference attachment from the item
        var body = Posted(env);
        var guid = FakeTeams.GuidOf("item-1").ToLowerInvariant();
        Assert.EndsWith($"</p><attachment id=\"{guid}\"></attachment>", (string?)body["body"]!["content"]);
        var attachment = Assert.Single(body["attachments"]!.AsArray())!;
        Assert.Equal(guid, (string?)attachment["id"]);
        Assert.Equal("reference", (string?)attachment["contentType"]);
        Assert.Equal($"https://example-my.sharepoint.com/personal/me_example_com/Documents/Microsoft Teams Chat Files/{name}", (string?)attachment["contentUrl"]);
        Assert.Equal(name, (string?)attachment["name"]);

        Assert.Contains("Read back OK.", env.Out.ToString());
        var record = Assert.Single(env.Records);
        Assert.Equal(("item-1", name), (record.Attachment, record.AttachmentName));
        Assert.Contains(env.Teams.Scopes, s => s.SequenceEqual(["Files.ReadWrite"]));
    }

    [Fact]
    public void Without_Files_ReadWrite_nothing_is_uploaded_or_sent_and_the_exit_code_says_so()
    {
        using var env = Scene();
        env.Teams.RefuseFiles = true;
        Assert.Equal(3, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--yes"));
        Assert.Equal(0, env.Teams.Writes);
        Assert.Empty(env.Records);
        Assert.Contains("Files.ReadWrite", env.Err.ToString());
        Assert.Contains("teams-post sign-in --files", env.Err.ToString());
    }

    [Fact]
    public void Someone_outside_the_organisation_needs_external()
    {
        using var env = Scene();
        env.Teams.Outsiders.Add("user-sam");
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md")));
        Assert.Contains("Sam Sample is outside your organisation", env.Out.ToString());
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--yes"));
        Assert.Equal(0, env.Teams.Writes);

        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--external", "--yes"));
        Assert.Equal(2, Assert.Single(env.Teams.Invites).Body["recipients"]!.AsArray().Count);
    }

    [Fact]
    public void A_share_that_fails_for_anyone_posts_nothing()
    {
        using var env = Scene();
        env.Teams.ShareFails.Add("sam.sample@example.com");
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--yes"));
        Assert.Contains("couldn't be shared with Sam Sample", env.Out.ToString());
        Assert.Empty(env.Teams.Messages);
        Assert.Empty(env.Records);
    }

    [Fact]
    public void A_read_back_without_the_attachment_is_different()
    {
        using var env = Scene();
        env.Teams.DropAttachments = true;
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--yes"));
        Assert.Contains("the attached file isn't there", env.Out.ToString());
    }

    [Fact]
    public void A_post_with_a_file_is_not_edited()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--yes"));
        Assert.Equal(1, env.Run("edit", env.MessageFile, "--yes"));
        Assert.Contains("with a file attached", env.Out.ToString());
    }

    [Fact]
    public void Sign_in_files_asks_for_Files_ReadWrite()
    {
        using var env = Scene();
        Assert.Equal(0, env.Run("sign-in", "--files"));
        Assert.Equal(0, env.Run("sign-in"));
        Assert.Equal(["sign-in --files", "sign-in"], env.SignIns);
    }

    [Fact]
    public void Attach_options_need_chat_and_attach()
    {
        using var env = Scene();
        Assert.Equal(2, env.Run("post", "marketing", env.MessageFile, "--attach", "x.md"));
        Assert.Equal(2, env.Run("post", "--chat", Chat, env.MessageFile, "--external"));
        Assert.Equal(2, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("transcript.md"), "--name", "a/b.md"));
        Assert.Equal(1, env.Run("post", "--chat", Chat, env.MessageFile, "--attach", env.Combine("missing.md")));
    }
}
