using System.Globalization;
using System.Text.RegularExpressions;
using PyCompat;

namespace TeamsPost;

/// <summary>What a command needs from outside: the settings, Teams, where it writes, the time and the waits.</summary>
public sealed class Session
{
    public required Config Config { get; init; }
    public required Func<TeamsChat> Chat { get; init; }
    public required TextWriter Out { get; init; }
    public Action<string> Log { get; init; } = _ => { };
    public Func<DateTimeOffset> Now { get; init; } = () => DateTimeOffset.Now;
    public Action<TimeSpan> Wait { get; init; } = wait => Thread.Sleep(wait);

    public PostStore Store => new(Config.StateFile);
}

/// <summary>post, edit and show. Each returns the exit code: 0 done, 1 refused or failed (a read-back that differs,
/// an answer that never came). Settings and arguments problems are thrown as UsageException (exit 2).</summary>
public static partial class Commands
{
    /// <summary>How long a just-posted message may answer "not found" before the read-back gives up.</summary>
    internal const int ReadBackTries = 5;
    internal static readonly TimeSpan ReadBackWait = TimeSpan.FromSeconds(2);

    [GeneratedRegex(@"<li[\s>]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ListItemTag();

    [GeneratedRegex(@"<p[\s>]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ParagraphTag();

    [GeneratedRegex(@"<at\s+id\s*=\s*[""']?0[""']?\s*>\s*Everyone\s*</at>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex EveryoneTag();

    [GeneratedRegex(@"^\d{1,20}\z", RegexOptions.CultureInvariant)]
    private static partial Regex MessageId();

    public static void CheckMessageId(string id)
    {
        if (!MessageId().IsMatch(id))
            throw new UsageException($"'{id}' isn't a Teams message id (they're all digits, e.g. 1763371234567).");
    }

    /// <summary>The file's message, as HTML for the target. Throws RefusedException (a draft, or nothing in it).</summary>
    public static string HtmlOf(Config config, string file, Target target)
    {
        string text;
        try
        {
            text = File.ReadAllText(file);
        }
        catch (Exception error) when (error is FileNotFoundException or DirectoryNotFoundException)
        {
            throw new RefusedException($"No file at {PostStore.Normalise(file)}.");
        }
        var markdown = TeamsMarkdown.MessageText(text, file, config.DraftMarker, config.StripFrontmatter);
        return TeamsMarkdown.HtmlFor(markdown, target.Mention);
    }

    /// <summary>Why a read-back doesn't match what was sent, or null if it does: the same number of list items and
    /// of paragraphs, and @Everyone there when it was sent.</summary>
    public static string? Mismatch(string sent, Mention mention, string chatId, ChatMessage back)
    {
        var problems = new List<string>();
        int sentItems = ListItemTag().Count(sent), backItems = ListItemTag().Count(back.Html);
        if (sentItems != backItems)
            problems.Add($"{sentItems} list items sent, {backItems} in Teams");
        int sentParagraphs = ParagraphTag().Count(sent), backParagraphs = ParagraphTag().Count(back.Html);
        if (sentParagraphs != backParagraphs)
            problems.Add($"{sentParagraphs} paragraphs sent, {backParagraphs} in Teams");
        if (mention == Mention.Everyone)
        {
            var mentioned = back.Mentions.Any(m => m.Text == "Everyone" && m.ConversationId == chatId && m.ConversationType == "chat");
            if (!mentioned || !EveryoneTag().IsMatch(back.Html))
                problems.Add("the @Everyone mention isn't there");
        }
        return problems.Count == 0 ? null : string.Join("; ", problems);
    }

    /// <summary>post: preview, or (yes) post once, record it, read it back.</summary>
    public static int Post(Session session, string targetName, string file, bool yes, bool again)
    {
        var config = session.Config;
        var target = config.TargetNamed(targetName);
        var chatId = target.ChatId;
        var html = HtmlOf(config, file, target);
        var full = PostStore.Normalise(file);
        var earlier = session.Store.Load().Where(r => r.Target == target.Name && PostStore.SameFile(r.File, full)).ToList();

        var output = session.Out;
        output.WriteLine($"To {target.Name} (chat {chatId}):");
        output.WriteLine();
        foreach (var line in PlainText.Lines(html))
            output.WriteLine("  " + line);
        output.WriteLine();

        if (earlier.Count > 0 && !again)
        {
            var last = earlier[^1];
            var what = last.Landed
                ? $"was already posted to {target.Name} on {last.PostedAt} (message {last.Id})"
                : $"was sent to {target.Name} on {last.PostedAt}, but no answer came back, so it may be there";
            output.WriteLine($"Not posted: {full} {what}.");
            output.WriteLine(last.Landed
                ? $"To change that message: teams-post edit \"{file}\". To post it again anyway: add --again."
                : $"Check with: teams-post show {target.Name}. To post it anyway: add --again.");
            return 1;
        }
        if (!yes)
        {
            output.WriteLine("Nothing has been sent. Run again with --yes to post.");
            return 0;
        }

        var chat = session.Chat();
        var sha = PostStore.Hash(html);
        var postedAt = Stamp(session.Now());
        ChatMessage posted;
        try
        {
            posted = chat.Post(chatId, TeamsMarkdown.Payload(html, target.Mention, chatId));
        }
        catch (DeliveryUnknownException error)
        {
            session.Store.Add(new PostRecord(full, target.Name, chatId, null, null, sha, "unknown", postedAt));
            session.Log($"post {target.Name}: no answer ({error.Message})");
            output.WriteLine($"No answer from Teams: {error.Message}");
            output.WriteLine($"It may have been posted, so it wasn't sent again. Check with: teams-post show {target.Name}");
            return 1;
        }
        catch (GraphRefusedException error)
        {
            session.Log($"post {target.Name}: refused ({error.Status})");
            output.WriteLine($"Not posted: {error.Message}");
            return 1;
        }

        if (string.IsNullOrEmpty(posted.Id))
        {
            session.Store.Add(new PostRecord(full, target.Name, chatId, null, null, sha, "unknown", postedAt));
            output.WriteLine($"Teams accepted the post but didn't say its id, so it can't be read back. Check with: teams-post show {target.Name}");
            return 1;
        }
        session.Store.Add(new PostRecord(full, target.Name, chatId, posted.Id, posted.Created, sha, "posted", postedAt));
        session.Log($"posted {target.Name} {posted.Id} from {full}");
        output.WriteLine($"Posted {posted.Id} at {posted.Created ?? postedAt}.");
        return ReadBack(session, chat, target.Name, chatId, posted.Id, html, target.Mention);
    }

    /// <summary>edit: find the message, show what would change, or (yes) replace it once and read it back.</summary>
    public static int Edit(Session session, string file, bool yes, string? targetName, string? messageId)
    {
        var config = session.Config;
        var full = PostStore.Normalise(file);
        PostRecord? record = null;
        Target target;
        if (messageId is not null)
        {
            CheckMessageId(messageId);
            if (targetName is null)
                throw new UsageException("edit --id needs --target, to say which chat the message is in.");
            target = config.TargetNamed(targetName);
            record = session.Store.Load().LastOrDefault(r => r.Id == messageId && r.Chat == target.ChatId);
        }
        else
        {
            var records = session.Store.Load().Where(r => PostStore.SameFile(r.File, full)).ToList();
            if (targetName is not null)
                records = records.Where(r => r.Target == targetName).ToList();
            var landed = records.Where(r => r.Landed).ToList();
            if (landed.Count == 0)
            {
                session.Out.WriteLine(records.Count > 0
                    ? $"{full} was sent{(targetName is null ? "" : $" to {targetName}")} but no answer came back, so which message it became isn't known. Find it with `teams-post show <target>`, then use --target and --id."
                    : $"No record of {full} being posted{(targetName is null ? "" : $" to {targetName}")}. If it was, use --target and --id.");
                return 1;
            }
            var targets = landed.Select(r => r.Target).Distinct().ToList();
            if (targets.Count > 1)
                throw new UsageException($"{full} was posted to {string.Join(" and ", targets)}: say which with --target.");
            record = landed[^1];
            target = config.TargetNamed(record.Target);
            messageId = record.Id!;
        }

        var chatId = record?.Chat is { Length: > 0 } recorded ? recorded : target.ChatId;
        var html = HtmlOf(config, file, target);
        var chat = session.Chat();
        ChatMessage current;
        try
        {
            current = chat.Get(chatId, messageId);
        }
        catch (GraphRefusedException error)
        {
            session.Out.WriteLine($"Couldn't read message {messageId} in {target.Name}: {error.Message}");
            return 1;
        }
        if (current.Deleted)
        {
            session.Out.WriteLine($"Message {messageId} in {target.Name} has been deleted, so it can't be edited.");
            return 1;
        }

        var oldLines = PlainText.Lines(current.Html);
        var newLines = PlainText.Lines(html);
        var sha = PostStore.Hash(html);
        if (oldLines.SequenceEqual(newLines) && (record is null || record.Sha256 == sha))
        {
            session.Out.WriteLine($"Message {messageId} in {target.Name} already reads as {Path.GetFileName(full)}: nothing to change.");
            return 0;
        }
        session.Out.WriteLine($"Message {messageId} in {target.Name} (chat {chatId}):");
        session.Out.WriteLine();
        var diff = UnifiedDiff.Lines(oldLines, newLines, "in Teams", Path.GetFileName(full));
        if (diff.Count == 0)
            session.Out.WriteLine("(the text reads the same; only formatting changes)");
        foreach (var line in diff)
            session.Out.WriteLine(line);
        session.Out.WriteLine();
        if (!yes)
        {
            session.Out.WriteLine("Nothing has been sent. Run again with --yes to edit the message.");
            return 0;
        }

        try
        {
            chat.Edit(chatId, messageId, TeamsMarkdown.Payload(html, target.Mention, chatId));
        }
        catch (DeliveryUnknownException error)
        {
            session.Log($"edit {target.Name} {messageId}: no answer ({error.Message})");
            session.Out.WriteLine($"No answer from Teams: {error.Message}");
            session.Out.WriteLine($"The edit may have been applied, so it wasn't sent again. Check with: teams-post show {target.Name} {messageId}");
            return 1;
        }
        catch (GraphRefusedException error)
        {
            session.Log($"edit {target.Name} {messageId}: refused ({error.Status})");
            session.Out.WriteLine($"Not edited: {error.Message}");
            if (error.Status == 403)
                session.Out.WriteLine("Editing needs Chat.ReadWrite, and only the person who posted a message can edit it.");
            return 1;
        }

        var editedAt = Stamp(session.Now());
        session.Store.Upsert(record is null
            ? new PostRecord(full, target.Name, chatId, messageId, current.Created, sha, "posted", editedAt, editedAt)
            : record with { File = full, Sha256 = sha, EditedAt = editedAt });
        session.Log($"edited {target.Name} {messageId} from {full}");
        session.Out.WriteLine($"Edited {messageId}.");
        return ReadBack(session, chat, target.Name, chatId, messageId, html, target.Mention);
    }

    /// <summary>show: one message, or the newest few, oldest first.</summary>
    public static int Show(Session session, string targetName, string? messageId, int last)
    {
        var target = session.Config.TargetNamed(targetName);
        var chatId = target.ChatId;
        var chat = session.Chat();
        IReadOnlyList<ChatMessage> messages;
        try
        {
            if (messageId is not null)
            {
                CheckMessageId(messageId);
                messages = [chat.Get(chatId, messageId)];
            }
            else
                messages = chat.Latest(chatId, last).Reverse().ToList();
        }
        catch (GraphRefusedException error)
        {
            session.Out.WriteLine(error.Message);
            return 1;
        }
        var first = true;
        foreach (var message in messages)
        {
            if (!first)
                session.Out.WriteLine();
            first = false;
            session.Out.WriteLine($"{message.Id}  {LocalTime(message.Created)}  {message.From ?? "(system)"}"
                + (message.MessageType == "message" ? "" : $"  [{message.MessageType}]"));
            if (message.Deleted)
                session.Out.WriteLine("  (deleted)");
            else
                foreach (var line in PlainText.Lines(message.Html))
                    session.Out.WriteLine("  " + line);
        }
        if (messages.Count == 0)
            session.Out.WriteLine($"No messages in {target.Name}.");
        return 0;
    }

    private static int ReadBack(Session session, TeamsChat chat, string targetName, string chatId, string messageId, string html, Mention mention)
    {
        ChatMessage? back = null;
        for (var attempt = 1; back is null; attempt++)
        {
            try
            {
                back = chat.Get(chatId, messageId);
            }
            catch (GraphRefusedException error) when (error.Status == 404 && attempt < ReadBackTries)
            {
                session.Wait(ReadBackWait);
            }
            catch (Exception error) when (error is GraphRefusedException or HttpRequestException or IOException or TaskCanceledException
                                          or M365Graph.GraphRequestException)
            {
                session.Out.WriteLine($"Read back FAILED: {error.Message}. Check it with: teams-post show {targetName} {messageId}");
                return 1;
            }
        }
        var problem = Mismatch(html, mention, chatId, back);
        if (problem is not null)
        {
            session.Log($"read back {messageId}: different ({problem})");
            session.Out.WriteLine($"Read back DIFFERENT: {problem}. Check it in Teams.");
            return 1;
        }
        session.Out.WriteLine("Read back OK.");
        return 0;
    }

    private static string Stamp(DateTimeOffset time) => time.ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture);

    private static string LocalTime(string? created) =>
        DateTimeOffset.TryParse(created, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var time)
            ? time.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
            : created ?? "?";
}
