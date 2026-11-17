using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using M365Graph;

namespace TeamsPost;

/// <summary>A message as Graph returns it: the parts teams-post reads.</summary>
public sealed record ChatMessage(string Id, string? Created, string? From, string MessageType, bool Deleted, string Html,
    IReadOnlyList<MentionInfo> Mentions)
{
    public static ChatMessage Parse(JsonNode? json)
    {
        var mentions = new List<MentionInfo>();
        if (json?["mentions"] is JsonArray array)
            foreach (var m in array)
                mentions.Add(new MentionInfo((int?)m?["id"] ?? -1, (string?)m?["mentionText"],
                    (string?)m?["mentioned"]?["conversation"]?["id"], (string?)m?["mentioned"]?["conversation"]?["conversationIdentityType"]));
        return new ChatMessage(
            (string?)json?["id"] ?? "",
            (string?)json?["createdDateTime"],
            (string?)json?["from"]?["user"]?["displayName"] ?? (string?)json?["from"]?["application"]?["displayName"],
            (string?)json?["messageType"] ?? "message",
            json?["deletedDateTime"] is JsonValue,
            (string?)json?["body"]?["content"] ?? "",
            mentions);
    }
}

public sealed record MentionInfo(int Id, string? Text, string? ConversationId, string? ConversationType);

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
/// been applied, so it's never resent; it becomes DeliveryUnknownException. Reads are retried as usual.</summary>
public sealed class TeamsChat(GraphClient graph)
{
    public static readonly string[] SendScopes = ["ChatMessage.Send"];
    public static readonly string[] ReadScopes = ["Chat.Read"];
    public static readonly string[] EditScopes = ["Chat.ReadWrite"];

    /// <summary>What sign-in asks for: enough to post and read back.</summary>
    public static readonly string[] SignInScopes = ["ChatMessage.Send", "Chat.Read"];

    public GraphClient Graph { get; } = graph;

    private static string Messages(string chatId) => $"/chats/{chatId}/messages";

    /// <summary>Post a new message; returns it as Graph created it.</summary>
    public ChatMessage Post(string chatId, JsonObject payload) =>
        ChatMessage.Parse(Write(HttpMethod.Post, Messages(chatId), payload, SendScopes));

    /// <summary>Replace a message's body (and mentions).</summary>
    public void Edit(string chatId, string messageId, JsonObject payload) =>
        Write(HttpMethod.Patch, $"{Messages(chatId)}/{messageId}", payload, EditScopes);

    public ChatMessage Get(string chatId, string messageId)
    {
        try
        {
            return ChatMessage.Parse(Graph.GetAsync($"{Messages(chatId)}/{messageId}", ReadScopes).GetAwaiter().GetResult());
        }
        catch (GraphRequestException error) when ((int)error.Status is >= 400 and < 500)
        {
            throw new GraphRefusedException((int)error.Status, error.Message);
        }
    }

    /// <summary>The newest messages, newest first (Graph allows up to 50 at a time).</summary>
    public IReadOnlyList<ChatMessage> Latest(string chatId, int count)
    {
        try
        {
            var page = Graph.GetAsync($"{Messages(chatId)}?$top={count}&$orderby=createdDateTime%20desc", ReadScopes).GetAwaiter().GetResult();
            return page["value"] is JsonArray items ? items.Select(ChatMessage.Parse).ToList() : [];
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
