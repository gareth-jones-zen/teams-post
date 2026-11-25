using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using M365Graph;

namespace TeamsPost;

/// <summary>A message as Graph returns it: the parts teams-post reads.</summary>
public sealed record ChatMessage(string Id, string? Created, string? From, string MessageType, bool Deleted, string Html,
    IReadOnlyList<MentionInfo> Mentions)
{
    /// <summary>The sender's user id; null for an app, a bot or a system event.</summary>
    public string? FromUserId { get; init; }

    /// <summary>Sent by an app or a bot (from.application), not a person.</summary>
    public bool FromApplication { get; init; }

    /// <summary>The sender as Graph gave it (the "from" object), for quoting the message.</summary>
    public JsonObject? FromJson { get; init; }

    /// <summary>The attachments' ids and content types (a quoted message is a "messageReference").</summary>
    public IReadOnlyList<(string? Id, string? ContentType)> Attachments { get; init; } = [];

    /// <summary>A message a person wrote: not a system event, not from an app, not deleted.</summary>
    public bool ByPerson => MessageType == "message" && !Deleted && !FromApplication && FromUserId is not null;

    public DateTimeOffset? CreatedAt =>
        DateTimeOffset.TryParse(Created, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeUniversal, out var at) ? at : null;

    public static ChatMessage Parse(JsonNode? json)
    {
        var mentions = new List<MentionInfo>();
        if (json?["mentions"] is JsonArray array)
            foreach (var m in array)
                mentions.Add(new MentionInfo((int?)m?["id"] ?? -1, (string?)m?["mentionText"],
                    (string?)m?["mentioned"]?["conversation"]?["id"], (string?)m?["mentioned"]?["conversation"]?["conversationIdentityType"])
                {
                    UserId = (string?)m?["mentioned"]?["user"]?["id"],
                });
        var attachments = new List<(string?, string?)>();
        if (json?["attachments"] is JsonArray items)
            foreach (var a in items)
                attachments.Add(((string?)a?["id"], (string?)a?["contentType"]));
        var from = json?["from"] as JsonObject;
        return new ChatMessage(
            (string?)json?["id"] ?? "",
            (string?)json?["createdDateTime"],
            (string?)from?["user"]?["displayName"] ?? (string?)from?["application"]?["displayName"],
            (string?)json?["messageType"] ?? "message",
            json?["deletedDateTime"] is JsonValue || json?["isDeleted"] is JsonValue deleted && deleted.TryGetValue<bool>(out var isDeleted) && isDeleted,
            (string?)json?["body"]?["content"] ?? "",
            mentions)
        {
            FromUserId = (string?)from?["user"]?["id"],
            FromApplication = from?["application"] is JsonObject,
            FromJson = (JsonObject?)from?.DeepClone(),
            Attachments = attachments,
        };
    }
}

public sealed record MentionInfo(int Id, string? Text, string? ConversationId, string? ConversationType)
{
    /// <summary>The person mentioned, when it's a person.</summary>
    public string? UserId { get; init; }
}

/// <summary>A chat as /me/chats lists it, with its last message's preview (null if it has none).</summary>
public sealed record ChatSummary(string Id, string? Topic, string ChatType, ChatMessage? Last);

/// <summary>Someone in a chat: their user id and name, and (Graph's aadUserConversationMember) email and home tenant.</summary>
public sealed record ChatMember(string? UserId, string? Name)
{
    public string? Email { get; init; }
    public string? TenantId { get; init; }
}

/// <summary>A file uploaded for a chat: its drive item id and name, the GUID from its eTag (a reference attachment's
/// id) and its WebDAV URL (the attachment's contentUrl), as Microsoft's "send a message with a file attachment"
/// example takes them.</summary>
public sealed record DriveFile(string ItemId, string Name, string AttachmentId, string ContentUrl, long Size);

/// <summary>Attaching a file needs Files.ReadWrite, and no sign-in holds it (never asked for, refused, or waiting
/// for an admin's approval). Nothing was uploaded or sent.</summary>
public sealed class FilesConsentException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>Graph answered a request with a refusal (4xx): nothing was changed.</summary>
public sealed class GraphRefusedException(int status, string message) : Exception(message)
{
    public int Status { get; } = status;
}

/// <summary>A post or an edit was sent, but whether it was applied isn't known (no answer, or a 5xx). It is never
/// sent again; reading the chat tells the truth.</summary>
public sealed class DeliveryUnknownException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>The Teams chat calls teams-post makes, through M365Graph.
///
/// The scopes are asked for per call, so each command needs only its own consent: posting needs ChatMessage.Send,
/// reading Chat.Read, and editing Chat.ReadWrite (Graph's only delegated permission for changing a chat message).
///
/// Writes go through SendRawAsync with GraphRetryPolicy.Write: sent again only after a 429 (refused before it was
/// looked at) or a connection that never opened. A timeout, a dropped connection after sending or a 5xx may have
/// been applied, so it's never resent: M365Graph throws GraphDeliveryUnknownException, which becomes
/// DeliveryUnknownException. A 4xx is a refusal (GraphRefusedException): nothing changed. Reads (GET) are retried
/// on a 429, a 5xx or a dropped connection.</summary>
public sealed partial class TeamsChat(GraphClient graph)
{
    public static readonly string[] SendScopes = ["ChatMessage.Send"];
    public static readonly string[] ReadScopes = ["Chat.Read"];
    public static readonly string[] EditScopes = ["Chat.ReadWrite"];

    /// <summary>Uploading a file to the signed-in person's OneDrive and sharing it (post --attach): asked for only
    /// when first needed, as editing's Chat.ReadWrite is.</summary>
    public static readonly string[] FileScopes = ["Files.ReadWrite"];

    /// <summary>The folder in a person's OneDrive where Teams keeps the files they attach in chats.</summary>
    public const string ChatFilesFolder = "Microsoft Teams Chat Files";
    public static readonly string[] MeScopes = ["User.Read"];

    /// <summary>What sign-in asks for: enough to post, read back, and know who's signed in.</summary>
    public static readonly string[] SignInScopes = ["ChatMessage.Send", "Chat.Read", "User.Read"];

    public GraphClient Graph { get; } = graph;

    private static string Messages(string chatId) => $"/chats/{chatId}/messages";

    /// <summary>Post a new message; returns it as Graph created it.</summary>
    public ChatMessage Post(string chatId, JsonObject payload) =>
        ChatMessage.Parse(Write(HttpMethod.Post, Messages(chatId), payload, SendScopes));

    /// <summary>Replace a message's body (and mentions).</summary>
    public void Edit(string chatId, string messageId, JsonObject payload) =>
        Write(HttpMethod.Patch, $"{Messages(chatId)}/{messageId}", payload, EditScopes);

    public ChatMessage Get(string chatId, string messageId) =>
        ChatMessage.Parse(Read($"{Messages(chatId)}/{messageId}", ReadScopes));

    /// <summary>The newest messages, newest first (Graph allows up to 50 at a time).</summary>
    public IReadOnlyList<ChatMessage> Latest(string chatId, int count)
    {
        var page = Read($"{Messages(chatId)}?$top={count}&$orderby=createdDateTime%20desc", ReadScopes);
        return page["value"] is JsonArray items ? items.Select(ChatMessage.Parse).ToList() : [];
    }

    /// <summary>Who's signed in: their user id and name (User.Read).</summary>
    public (string Id, string? Name) Me()
    {
        var me = Read("/me?$select=id,displayName", MeScopes);
        return ((string?)me["id"] is { Length: > 0 } id ? id : throw new InvalidOperationException("Graph didn't say who's signed in."),
            (string?)me["displayName"]);
    }

    /// <summary>The signed-in person's chats, the most recently active first, until (and including) the first whose
    /// last message is older than since (or has none). At most pages × 50 chats.</summary>
    public IReadOnlyList<ChatSummary> ChatsSince(DateTimeOffset since, int pages = 20)
    {
        var chats = new List<ChatSummary>();
        string? next = "/me/chats?$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime%20desc&$top=50";
        for (var page = 0; page < pages && !string.IsNullOrEmpty(next); page++)
        {
            var reply = Read(next, ReadScopes);
            next = (string?)reply["@odata.nextLink"];
            foreach (var item in reply["value"] as JsonArray ?? [])
            {
                var last = item?["lastMessagePreview"] is JsonObject preview ? ChatMessage.Parse(preview) : null;
                chats.Add(new ChatSummary((string?)item?["id"] ?? "", (string?)item?["topic"], (string?)item?["chatType"] ?? "", last));
                if (last?.CreatedAt is not { } at || at < since)
                    return chats;
            }
        }
        return chats;
    }

    /// <summary>A chat by id: its topic and type. A chat the signed-in person isn't in (or a channel's id) is a
    /// refusal (GraphRefusedException).</summary>
    public ChatSummary Chat(string chatId)
    {
        var chat = Read($"/chats/{chatId}", ReadScopes);
        return new ChatSummary((string?)chat["id"] ?? chatId, (string?)chat["topic"], (string?)chat["chatType"] ?? "", null);
    }

    /// <summary>The people in a chat.</summary>
    public IReadOnlyList<ChatMember> Members(string chatId)
    {
        var reply = Read($"/chats/{chatId}/members", ReadScopes);
        return (reply["value"] as JsonArray ?? []).Select(m => new ChatMember((string?)m?["userId"], (string?)m?["displayName"])
        {
            Email = (string?)m?["email"] is { Length: > 0 } email ? email : null,
            TenantId = (string?)m?["tenantId"],
        }).ToList();
    }

    /// <summary>A Files.ReadWrite token, or FilesConsentException: checked before anything is uploaded, so a missing
    /// consent stops the run with nothing changed. A browser opens to ask only if someone's at the console, as for
    /// every other scope.</summary>
    public void RequireFiles()
    {
        try
        {
            Graph.TokenAsync(FileScopes).GetAwaiter().GetResult();
        }
        catch (SignInNeededException error)
        {
            throw new FilesConsentException(
                "Attaching a file needs the Files.ReadWrite permission (to upload it to your OneDrive and share it with the "
                + $"chat's members), and no sign-in holds it. {error.Message}\n"
                + "At a console, run: teams-post sign-in --files. If it shows \"Approval required\", an admin has to approve it.",
                error);
        }
    }

    /// <summary>Uploads a file to the signed-in person's "Microsoft Teams Chat Files" folder (replacing one of the same
    /// name, so a run that stopped part-way can simply be run again), then reads what a reference attachment needs.
    /// A PUT of the same content changes nothing the second time, so it's retried as M365Graph retries any PUT.</summary>
    public DriveFile Upload(string name, byte[] content)
    {
        var url = $"/me/drive/root:/{Uri.EscapeDataString(ChatFilesFolder)}/{Uri.EscapeDataString(name)}:/content";
        var uploaded = Send(HttpMethod.Put, url, content, "application/octet-stream", GraphRetryPolicy.For(HttpMethod.Put), FileScopes);
        var id = (string?)uploaded?["id"] is { Length: > 0 } itemId ? itemId
            : throw new InvalidOperationException("OneDrive didn't say what the uploaded file's id is.");
        var item = Read($"/me/drive/items/{id}?$select=id,name,eTag,webDavUrl,size", FileScopes);
        var etag = (string?)item["eTag"] ?? "";
        var guid = GuidIn().Match(etag);
        if (!guid.Success)
            throw new InvalidOperationException($"The uploaded file's eTag ({etag}) has no GUID to attach it by.");
        var webDav = (string?)item["webDavUrl"] is { Length: > 0 } dav ? dav
            : throw new InvalidOperationException("OneDrive didn't give the uploaded file's webDavUrl.");
        return new DriveFile(id, (string?)item["name"] ?? name, guid.Value.ToLowerInvariant(), webDav, (long?)item["size"] ?? content.Length);
    }

    [System.Text.RegularExpressions.GeneratedRegex("[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")]
    private static partial System.Text.RegularExpressions.Regex GuidIn();

    /// <summary>Shares the file with these people by name, read-only and without an email (driveItem invite:
    /// requireSignIn, sendInvitation false, roles read): never a link that anyone, or the whole organisation, can open.
    /// Sent once, as every write is. Returns who it couldn't be shared with (a 207 says per person); empty if everyone.</summary>
    public IReadOnlyList<string> Share(string itemId, IReadOnlyList<ChatMember> people)
    {
        var recipients = new JsonArray();
        foreach (var person in people)
            recipients.Add((JsonNode)(person.Email is { } email
                ? new JsonObject { ["email"] = email }
                : new JsonObject { ["objectId"] = person.UserId }));
        var payload = new JsonObject
        {
            ["recipients"] = recipients,
            ["requireSignIn"] = true,
            ["sendInvitation"] = false,
            ["roles"] = new JsonArray("read"),
        };
        var reply = Write(HttpMethod.Post, $"/me/drive/items/{itemId}/invite", payload, FileScopes);
        var granted = reply?["value"] as JsonArray ?? [];
        var failed = new List<string>();
        for (var i = 0; i < people.Count; i++)
        {
            var permission = i < granted.Count ? granted[i] : null;
            if (permission is null || permission["error"] is JsonObject)
                failed.Add(people[i].Name ?? people[i].Email ?? people[i].UserId ?? "?");
        }
        return failed;
    }

    /// <summary>A read: retried by M365Graph after a 429, a 5xx or a dropped connection; a 4xx is a
    /// GraphRefusedException.</summary>
    private JsonObject Read(string url, string[] scopes)
    {
        try
        {
            return Graph.GetAsync(url, scopes).GetAwaiter().GetResult();
        }
        catch (GraphRequestException error) when ((int)error.Status is >= 400 and < 500)
        {
            throw new GraphRefusedException((int)error.Status, error.Message);
        }
    }

    private JsonNode? Write(HttpMethod method, string url, JsonObject payload, string[] scopes)
        => Send(method, url, Encoding.UTF8.GetBytes(payload.ToJsonString()), "application/json", GraphRetryPolicy.Write, scopes);

    private JsonNode? Send(HttpMethod method, string url, byte[] body, string contentType, GraphRetryPolicy policy, string[] scopes)
    {
        GraphRawResponse reply;
        try
        {
            reply = Graph.SendRawAsync(method, url, body, contentType, policy, scopes).GetAwaiter().GetResult();
        }
        catch (GraphDeliveryUnknownException error)
        {
            throw new DeliveryUnknownException(error.Message, error);
        }
        var status = (int)reply.Status;
        var detail = Encoding.UTF8.GetString(reply.Body);
        if (detail.Length > 300)
            detail = detail[..300];
        if (status >= 500)
            throw new DeliveryUnknownException($"Teams said {status} to {method.Method} {url}: {detail} (not sent again: it may have been applied)");
        if (!reply.IsSuccess)
            throw new GraphRefusedException(status, $"Teams refused {method.Method} {url} ({status}): {detail}");
        if (reply.Body.Length == 0)
            return null;
        try
        {
            return JsonNode.Parse(reply.Body);
        }
        catch (JsonException)
        {
            return null;
        }
    }
}
