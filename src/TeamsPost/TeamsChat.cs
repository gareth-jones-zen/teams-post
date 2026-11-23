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

/// <summary>Someone in a chat.</summary>
public sealed record ChatMember(string? UserId, string? Name);

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
public sealed class TeamsChat(GraphClient graph)
{
    public static readonly string[] SendScopes = ["ChatMessage.Send"];
    public static readonly string[] ReadScopes = ["Chat.Read"];
    public static readonly string[] EditScopes = ["Chat.ReadWrite"];
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
        return (reply["value"] as JsonArray ?? []).Select(m => new ChatMember((string?)m?["userId"], (string?)m?["displayName"])).ToList();
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
    {
        GraphRawResponse reply;
        try
        {
            reply = Graph.SendRawAsync(method, url, Encoding.UTF8.GetBytes(payload.ToJsonString()), "application/json",
                GraphRetryPolicy.Write, scopes).GetAwaiter().GetResult();
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
