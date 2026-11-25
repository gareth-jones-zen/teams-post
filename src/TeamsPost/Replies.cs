using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace TeamsPost;

/// <summary>unanswered and reply: finding the chats where someone else spoke last, and answering one from a file.</summary>
public static partial class Commands
{
    /// <summary>The state file's target for a reply (not a valid target name, so it can't clash with one).</summary>
    public const string ReplyTarget = "(reply)";

    /// <summary>How many of a chat's newest messages unanswered reads, and how many of the others' it shows.</summary>
    internal const int ChatMessagesRead = 30;
    internal const int MessagesShown = 5;
    internal const int TextLimit = 500;
    internal const int QuotePreviewLimit = 200;

    [GeneratedRegex(@"^19:[^\s/?#%\\]+\z", RegexOptions.CultureInvariant)]
    private static partial Regex ChatIdPattern();

    [GeneratedRegex(@"@thread\.(tacv2|skype)\z", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex ChannelIdPattern();

    [GeneratedRegex(@"<attachment\b[^>]*>\s*</attachment>", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex AttachmentTag();

    [GeneratedRegex(@"\b(?:https?|mailto):\S+", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex Address();

    /// <summary>A chat id that's safe in a URL and isn't a channel's. Throws UsageException.</summary>
    public static void CheckChatId(string chatId)
    {
        if (ChannelIdPattern().IsMatch(chatId))
            throw new UsageException($"'{chatId}' is a channel's id: reply only answers in chats (one-to-one, group and meeting chats).");
        if (!ChatIdPattern().IsMatch(chatId))
            throw new UsageException($"'{chatId}' isn't a chat id (\"19:...@thread.v2\" or \"19:..._...@unq.gbl.spaces\").");
    }

    /// <summary>A message's text, one line per paragraph, without the quote of another message it may start with.</summary>
    public static string TextOf(string html) => PlainText.Of(AttachmentTag().Replace(html, ""));

    private static string Trim(string text, int limit) => text.Length <= limit ? text : text[..(limit - 1)].TrimEnd() + "…";

    // ---------------------------------------------------------------- unanswered

    /// <summary>One chat where someone else spoke last: their messages since the signed-in person's last one,
    /// newest first.</summary>
    public sealed record UnansweredChat(ChatSummary Chat, string Title, IReadOnlyList<ChatMessage> Theirs, bool Direct,
        bool MentionsYou, bool Question);

    /// <summary>unanswered: the chats (not channels) where the newest message a person wrote is someone else's and
    /// is inside the window, newest first, flagged direct / mentionsYou / question. Meeting chats only where the
    /// signed-in person is @mentioned in the window. Bots, apps, system events and deleted messages don't count.</summary>
    public static int Unanswered(Session session, int? hours, bool json)
    {
        var settings = session.Config.Unanswered;
        var now = session.Now();
        var since = hours is not null
            ? UnansweredSettings.Since(now, hours, null)
            : UnansweredSettings.Since(now, settings.Hours, settings.WorkingDays);
        var window = hours is not null ? $"{hours} hours"
            : settings.WorkingDays is { } days ? $"{days} working day{(days == 1 ? "" : "s")}"
            : $"{settings.Hours} hours";

        var chat = session.Chat();
        var (me, _) = chat.Me();
        var found = new List<UnansweredChat>();
        var unreadable = new List<string>();
        foreach (var summary in chat.ChatsSince(since))
        {
            if (summary.Last?.CreatedAt is not { } lastAt || lastAt < since)
                continue;
            if (settings.SkipChats.Any(s => s.ChatId is not null && s.Matches(summary.Id, "")))
                continue;
            if (summary.Last.ByPerson && summary.Last.FromUserId == me)
                continue;   // his is the last word

            IReadOnlyList<ChatMessage> messages;
            try
            {
                messages = chat.Latest(summary.Id, ChatMessagesRead);
            }
            catch (GraphRefusedException error)
            {
                unreadable.Add($"{summary.Topic ?? summary.Id} ({error.Status})");
                continue;
            }
            var theirs = new List<ChatMessage>();
            foreach (var message in messages.OrderByDescending(m => m.CreatedAt))
            {
                if (!message.ByPerson)
                    continue;
                if (message.FromUserId == me)
                    break;
                if (!settings.SkipsSender(message.From, message.FromUserId))
                    theirs.Add(message);
            }
            if (theirs.Count == 0 || theirs[0].CreatedAt is not { } newest || newest < since)
                continue;

            var direct = summary.ChatType == "oneOnOne";
            var title = summary.Topic is { Length: > 0 } topic ? topic
                : direct ? theirs[0].From ?? "(someone)"
                : string.Join(", ", theirs.Select(m => m.From ?? "?").Distinct().Take(4));
            if (settings.SkipChats.Any(s => s.Name is not null && s.Matches(summary.Id, title)))
                continue;
            bool Mentions(ChatMessage m) => m.Mentions.Any(x => x.UserId == me);
            if (summary.ChatType == "meeting" && !theirs.Any(m => m.CreatedAt >= since && Mentions(m)))
                continue;
            found.Add(new UnansweredChat(summary, title, theirs, direct, theirs.Any(Mentions),
                theirs.Any(m => Address().Replace(TextOf(m.Html), "").Contains('?'))));
        }
        found = found.OrderByDescending(f => f.Theirs[0].CreatedAt).ToList();

        if (json)
            WriteUnansweredJson(session.Out, since, window, found, unreadable);
        else
            WriteUnanswered(session.Out, since, window, found, unreadable);
        return 0;
    }

    private static IEnumerable<string> Flags(UnansweredChat chat)
    {
        if (chat.Direct)
            yield return "direct";
        if (chat.MentionsYou)
            yield return "mentionsYou";
        if (chat.Question)
            yield return "question";
    }

    private static string ChatTypeName(string type) => type switch
    {
        "oneOnOne" => "one-to-one",
        "group" => "group chat",
        "meeting" => "meeting chat",
        _ => type,
    };

    private static string When(DateTimeOffset? at) =>
        at?.ToLocalTime().ToString("ddd d MMM HH:mm", CultureInfo.InvariantCulture) ?? "?";

    private static void WriteUnanswered(TextWriter output, DateTimeOffset since, string window, List<UnansweredChat> found, List<string> unreadable)
    {
        output.WriteLine($"Since {When(since)} ({window}): {found.Count} chat{(found.Count == 1 ? "" : "s")} where someone else spoke last.");
        foreach (var chat in found)
        {
            output.WriteLine();
            var flags = string.Join(", ", Flags(chat));
            output.WriteLine($"{chat.Title}  ({ChatTypeName(chat.Chat.ChatType)}{(flags.Length > 0 ? "; " + flags : "")})");
            output.WriteLine($"  chat {chat.Chat.Id}");
            output.WriteLine($"  reply --to {chat.Theirs[0].Id}");
            if (chat.Theirs.Count > MessagesShown)
                output.WriteLine($"  ({chat.Theirs.Count - MessagesShown} earlier message{(chat.Theirs.Count - MessagesShown == 1 ? "" : "s")} since yours not shown)");
            foreach (var message in chat.Theirs.Take(MessagesShown).Reverse())
            {
                output.WriteLine($"  {When(message.CreatedAt)}  {message.From}");
                foreach (var line in Trim(TextOf(message.Html), TextLimit).Split('\n'))
                    output.WriteLine("    " + line);
            }
        }
        if (unreadable.Count > 0)
        {
            output.WriteLine();
            output.WriteLine($"Couldn't read {unreadable.Count} chat{(unreadable.Count == 1 ? "" : "s")}: {string.Join("; ", unreadable)}");
        }
    }

    private static void WriteUnansweredJson(TextWriter output, DateTimeOffset since, string window, List<UnansweredChat> found, List<string> unreadable)
    {
        var chats = new JsonArray();
        foreach (var chat in found)
        {
            var messages = new JsonArray();
            foreach (var message in chat.Theirs.Take(MessagesShown).Reverse())
                messages.Add((JsonNode)new JsonObject
                {
                    ["id"] = message.Id,
                    ["from"] = message.From,
                    ["created"] = message.Created,
                    ["text"] = Trim(TextOf(message.Html), TextLimit),
                });
            chats.Add((JsonNode)new JsonObject
            {
                ["chat"] = chat.Chat.Id,
                ["title"] = chat.Title,
                ["chatType"] = chat.Chat.ChatType,
                ["flags"] = new JsonObject { ["direct"] = chat.Direct, ["mentionsYou"] = chat.MentionsYou, ["question"] = chat.Question },
                ["replyTo"] = chat.Theirs[0].Id,
                ["lastAt"] = chat.Theirs[0].Created,
                ["since"] = chat.Theirs.Count,
                ["messages"] = messages,
            });
        }
        var root = new JsonObject
        {
            ["since"] = since.ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture),
            ["window"] = window,
            ["chats"] = chats,
            ["unreadable"] = new JsonArray(unreadable.Select(u => (JsonNode?)JsonValue.Create(u)).ToArray()),
        };
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }))
            root.WriteTo(writer);
        output.WriteLine(Encoding.UTF8.GetString(stream.ToArray()));
    }

    // ---------------------------------------------------------------- reply

    /// <summary>The quote of another message, as Teams' own Reply stores it in a chat: an attachment tag at the
    /// start of the body, and a messageReference attachment whose content is the quoted message's id, the first
    /// 200 characters of its text and its sender.</summary>
    public static JsonObject ReplyPayload(string html, ChatMessage? quoted)
    {
        if (quoted is null)
            return TeamsMarkdown.Payload(html, Mention.None, "");
        var content = new JsonObject
        {
            ["messageId"] = quoted.Id,
            ["messagePreview"] = Trim(TextOf(quoted.Html).Replace('\n', ' '), QuotePreviewLimit),
            ["messageSender"] = SenderForQuote(quoted.FromJson),
        };
        return new JsonObject
        {
            ["body"] = new JsonObject
            {
                ["contentType"] = "html",
                ["content"] = $"<attachment id=\"{quoted.Id}\"></attachment>{html}",
            },
            ["attachments"] = new JsonArray(new JsonObject
            {
                ["id"] = quoted.Id,
                ["contentType"] = "messageReference",
                ["contentUrl"] = null,
                ["content"] = content.ToJsonString(),
                ["name"] = null,
                ["thumbnailUrl"] = null,
                ["teamsAppId"] = null,
            }),
        };
    }

    /// <summary>The quoted message's sender (Graph's from), without the @odata.type annotations.</summary>
    private static JsonObject SenderForQuote(JsonObject? from)
    {
        JsonNode? Part(string name)
        {
            if (from?[name] is not JsonObject part)
                return null;
            var copy = (JsonObject)part.DeepClone();
            copy.Remove("@odata.type");
            return copy;
        }
        return new JsonObject { ["application"] = Part("application"), ["device"] = Part("device"), ["user"] = Part("user") };
    }

}
