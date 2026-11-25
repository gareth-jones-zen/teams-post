using System.Text.Json.Nodes;

namespace TeamsPost;

/// <summary>A file to attach to a post (post --chat --attach): the file, the name it's uploaded as (null: its own
/// name), and whether people outside the organisation may be given it.</summary>
public sealed record AttachRequest(string File, string? Name, bool External);

/// <summary>reply, and post --chat: a file sent into any chat the signed-in person is in, by its id, once per chat,
/// read back; a reply can quote the message it answers, a post can carry a file.</summary>
public static partial class Commands
{
    /// <summary>The state file's target for a post --chat (not a valid target name, so it can't clash with one).</summary>
    public const string ChatTarget = "(chat)";

    /// <summary>The largest file --attach uploads (a single PUT).</summary>
    public const long MaxAttachBytes = 4 * 1024 * 1024;

    /// <summary>reply: preview, or (yes) post the file into a chat the signed-in person is in, once, optionally
    /// quoting the message it answers, and read it back.</summary>
    public static int Reply(Session session, string chatId, string file, string? toId, bool yes, bool again) =>
        SendInChat(session, chatId, file, toId, null, yes, again, ReplyTarget);

    /// <summary>post --chat: preview, or (yes) post the file into a chat the signed-in person is in, by its id, once,
    /// with a file attached if asked (uploaded to their OneDrive, shared with the chat's members by name), and read
    /// it back.</summary>
    public static int PostToChat(Session session, string chatId, string file, AttachRequest? attach, bool yes, bool again) =>
        SendInChat(session, chatId, file, null, attach, yes, again, ChatTarget);

    private static readonly char[] BadNameCharacters = ['/', '\\', ':', '*', '?', '"', '<', '>', '|'];

    private static int SendInChat(Session session, string chatId, string file, string? toId, AttachRequest? attach, bool yes, bool again,
        string recordTarget)
    {
        var verb = recordTarget == ReplyTarget ? "reply" : "post";
        CheckChatId(chatId);
        if (toId is not null)
            CheckMessageId(toId);
        var config = session.Config;
        string html;
        try
        {
            html = TeamsMarkdown.ToHtml(TeamsMarkdown.MessageText(File.ReadAllText(file), file, config.DraftMarker, config.StripFrontmatter));
        }
        catch (Exception error) when (error is FileNotFoundException or DirectoryNotFoundException)
        {
            throw new RefusedException($"No file at {PostStore.Normalise(file)}.");
        }
        var full = PostStore.Normalise(file);
        byte[]? content = null;
        string? attachName = null;
        if (attach is not null)
        {
            try
            {
                content = File.ReadAllBytes(attach.File);
            }
            catch (Exception error) when (error is FileNotFoundException or DirectoryNotFoundException)
            {
                throw new RefusedException($"No file to attach at {PostStore.Normalise(attach.File)}.");
            }
            if (content.Length > MaxAttachBytes)
                throw new RefusedException($"{PostStore.Normalise(attach.File)} is {content.Length / 1024} KB: --attach takes files up to {MaxAttachBytes / 1024 / 1024} MB.");
            attachName = attach.Name ?? Path.GetFileName(attach.File);
            if (attachName.Trim().Length == 0 || attachName.IndexOfAny(BadNameCharacters) >= 0 || attachName.EndsWith('.'))
                throw new UsageException($"'{attachName}' can't be a file's name in OneDrive (no / \\ : * ? \" < > |, and not ending in a dot).");
        }
        var output = session.Out;

        var chat = session.Chat();
        ChatSummary info;
        try
        {
            info = chat.Chat(chatId);
        }
        catch (GraphRefusedException error) when (error.Status is 400 or 403 or 404)
        {
            output.WriteLine($"Not sent: {chatId} isn't a chat you can read ({error.Status}). {verb} only sends to chats you're in, not channels.");
            return 1;
        }
        var (me, _) = chat.Me();
        var members = chat.Members(chatId);
        var mine = members.FirstOrDefault(m => m.UserId == me);
        if (mine is null)
        {
            output.WriteLine($"Not sent: you aren't a member of {chatId}.");
            return 1;
        }
        var others = members.Where(m => m.UserId != me).ToList();
        var names = others.Select(m => m.Name ?? "?").ToList();
        var title = info.Topic is { Length: > 0 } topic ? topic
            : names.Count == 0 ? "(just you)"
            : string.Join(", ", names.Take(4)) + (names.Count > 4 ? $" and {names.Count - 4} more" : "");

        ChatMessage? quoted = null;
        if (toId is not null)
        {
            try
            {
                quoted = chat.Get(chatId, toId);
            }
            catch (GraphRefusedException error) when (error.Status is 400 or 403 or 404)
            {
                // Graph answers 403, not 404, for a message id that isn't in the chat
                output.WriteLine($"Not sent: there's no message {toId} in that chat (Teams said {error.Status}).");
                return 1;
            }
            if (quoted.Deleted || quoted.MessageType != "message")
            {
                output.WriteLine($"Not sent: message {toId} is {(quoted.Deleted ? "deleted" : "a system event")}, so it can't be quoted.");
                return 1;
            }
        }

        output.WriteLine($"To {title} ({ChatTypeName(info.ChatType)}, chat {chatId}):");
        if (quoted is not null)
            output.WriteLine($"Replying to {quoted.From ?? "?"}, {When(quoted.CreatedAt)}: \"{Trim(TextOf(quoted.Html).Replace('\n', ' '), QuotePreviewLimit)}\"");
        output.WriteLine();
        foreach (var line in PlainText.Lines(html))
            output.WriteLine("  " + line);
        output.WriteLine();

        if (attach is not null)
        {
            // Shared by name with each member, so only someone outside the organisation needs saying yes to.
            var outside = others.Where(m => m.UserId is null || (m.TenantId is { Length: > 0 } tenant && mine.TenantId is { Length: > 0 } own
                                                                   && !string.Equals(tenant, own, StringComparison.OrdinalIgnoreCase))).ToList();
            var size = content!.Length < 1024 ? $"{content.Length} bytes" : $"{(content.Length + 1023) / 1024} KB";
            output.WriteLine($"Attached: {attachName} ({size}), uploaded to your OneDrive ({TeamsChat.ChatFilesFolder}) and shared, "
                             + (others.Count == 0 ? "with no one (you're the only member)." : $"read-only, with {others.Count} by name: {string.Join(", ", names)}."));
            if (outside.Count > 0)
            {
                var who = string.Join(", ", outside.Select(m => m.Name ?? m.Email ?? "(someone unnamed)"));
                if (!attach.External)
                {
                    output.WriteLine($"Not sent: {who} {(outside.Count == 1 ? "is" : "are")} outside your organisation, and the file would be shared with them. "
                                     + "Add --external to share it with them anyway.");
                    return 1;
                }
                output.WriteLine($"Outside your organisation (--external): {who}.");
            }
            if (others.FirstOrDefault(m => m.Email is null && m.UserId is null) is { } nobody)
            {
                output.WriteLine($"Not sent: {nobody.Name ?? "a member"} has neither an email address nor a user id, so the file can't be shared with them.");
                return 1;
            }
            output.WriteLine();
        }

        var earlier = session.Store.Load().Where(r => r.Chat == chatId && PostStore.SameFile(r.File, full)).ToList();
        if (earlier.Count > 0 && !again)
        {
            var last = earlier[^1];
            output.WriteLine(last.Landed
                ? $"Not sent: {full} was already sent to this chat on {last.PostedAt} (message {last.Id}). To send it again anyway: add --again."
                : $"Not sent: {full} was sent to this chat on {last.PostedAt}, but no answer came back, so it may be there. "
                  + $"Check with: teams-post show {chatId}. To send it anyway: add --again.");
            return 1;
        }
        if (!yes)
        {
            output.WriteLine($"Nothing has been sent. Run again with --yes to send this {verb}.");
            return 0;
        }

        DriveFile? uploaded = null;
        if (attach is not null)
        {
            chat.RequireFiles();   // FilesConsentException (exit 3) before anything is uploaded
            try
            {
                uploaded = chat.Upload(attachName!, content!);
            }
            catch (Exception error) when (error is GraphRefusedException or DeliveryUnknownException)
            {
                session.Log($"{verb} {chatId}: upload of {attachName} failed ({error.Message})");
                output.WriteLine($"Not sent: the upload of {attachName} failed: {error.Message}");
                output.WriteLine("Nothing was posted. Running it again uploads the file again (replacing it).");
                return 1;
            }
            if (others.Count > 0)
            {
                IReadOnlyList<string> unshared;
                try
                {
                    unshared = chat.Share(uploaded.ItemId, others);
                }
                catch (Exception error) when (error is GraphRefusedException or DeliveryUnknownException)
                {
                    session.Log($"{verb} {chatId}: sharing {attachName} failed ({error.Message})");
                    output.WriteLine($"Not sent: sharing {attachName} with the chat's members failed: {error.Message}");
                    output.WriteLine("Nothing was posted. The file is in your OneDrive; running it again uploads and shares it again.");
                    return 1;
                }
                if (unshared.Count > 0)
                {
                    session.Log($"{verb} {chatId}: {attachName} not shared with {unshared.Count} member(s)");
                    output.WriteLine($"Not sent: {attachName} couldn't be shared with {string.Join(", ", unshared)}, so they couldn't open it. Nothing was posted.");
                    return 1;
                }
            }
            output.WriteLine($"Uploaded {uploaded.Name} and shared it with {others.Count} {(others.Count == 1 ? "person" : "people")}.");
        }

        var payload = ReplyPayload(html, quoted);
        if (uploaded is not null)
            Attach(payload, uploaded);
        var sent = (string)payload["body"]!["content"]!;
        var sha = PostStore.Hash(sent);
        var postedAt = Stamp(session.Now());
        PostRecord Record(string? id, string? created, string status) =>
            new(full, recordTarget, chatId, id, created, sha, status, postedAt)
            {
                ReplyTo = toId,
                Attachment = uploaded?.ItemId,
                AttachmentName = uploaded?.Name,
            };
        ChatMessage posted;
        try
        {
            posted = chat.Post(chatId, payload);
        }
        catch (DeliveryUnknownException error)
        {
            session.Store.Add(Record(null, null, "unknown"));
            session.Log($"{verb} {chatId}: no answer ({error.Message})");
            output.WriteLine($"No answer from Teams: {error.Message}");
            output.WriteLine($"It may have been sent, so it wasn't sent again. Check with: teams-post show {chatId}");
            return 1;
        }
        catch (GraphRefusedException error)
        {
            session.Log($"{verb} {chatId}: refused ({error.Status})");
            output.WriteLine($"Not sent: {error.Message}");
            return 1;
        }
        if (string.IsNullOrEmpty(posted.Id))
        {
            session.Store.Add(Record(null, null, "unknown"));
            output.WriteLine($"Teams accepted the {verb} but didn't say its id, so it can't be read back. Check with: teams-post show {chatId}");
            return 1;
        }
        session.Store.Add(Record(posted.Id, posted.Created, "posted"));
        session.Log($"{(verb == "reply" ? "replied in" : "posted in")} {chatId} {posted.Id} from {full}"
                    + (uploaded is null ? "" : $", {uploaded.Name} attached"));
        output.WriteLine($"Sent {posted.Id} at {posted.Created ?? postedAt}.");
        return ReadBack(session, chat, chatId, chatId, posted.Id, sent, Mention.None, toId, uploaded?.AttachmentId);
    }

    /// <summary>A file attached as Teams' own attachments are: an attachment tag at the end of the body, and a
    /// "reference" attachment with the id from the file's eTag, its WebDAV URL and its name.</summary>
    public static void Attach(JsonObject payload, DriveFile file)
    {
        var body = (JsonObject)payload["body"]!;
        body["content"] = (string)body["content"]! + $"<attachment id=\"{file.AttachmentId}\"></attachment>";
        if (payload["attachments"] is not JsonArray attachments)
            payload["attachments"] = attachments = [];
        attachments.Add((JsonNode)new JsonObject
        {
            ["id"] = file.AttachmentId,
            ["contentType"] = "reference",
            ["contentUrl"] = file.ContentUrl,
            ["name"] = file.Name,
        });
    }
}
