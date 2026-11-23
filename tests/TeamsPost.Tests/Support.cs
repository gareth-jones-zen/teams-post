using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using M365Graph;

namespace TeamsPost.Tests;

/// <summary>A fake Teams behind a fake HTTP handler, so the real GraphClient (and its retry policy for writes) is
/// what's tested. It keeps each chat's messages, answers POST, PATCH and GET as Graph does, and can be told to lose
/// an answer, fail with a status, or store a message differently from how it was sent.</summary>
public sealed partial class FakeTeams : HttpMessageHandler
{
    public sealed class Message
    {
        public required string Id { get; init; }
        public required string Chat { get; init; }
        public required string Created { get; set; }
        public string From { get; set; } = "Test User";
        public string FromId { get; set; } = "user-test";
        public bool App { get; set; }
        public string MessageType { get; set; } = "message";
        public JsonArray Attachments { get; set; } = [];
        public string Html { get; set; } = "";
        public JsonArray Mentions { get; set; } = [];
        public bool Deleted { get; set; }
    }

    public List<Message> Messages { get; } = [];
    public List<(string Method, string Url, string? Body)> Requests { get; } = [];
    public List<IReadOnlyList<string>> Scopes { get; } = [];
    private long nextId = 1763370000000;

    /// <summary>The next write is applied, then its answer is lost (a dropped connection).</summary>
    public bool LoseNextWrite { get; set; }

    /// <summary>The next write is applied, then the request times out before the answer comes.</summary>
    public bool TimeOutNextWrite { get; set; }

    /// <summary>The next GET gets this status (Graph having a bad moment), and nothing is read.</summary>
    public HttpStatusCode? FailNextRead { get; set; }

    /// <summary>The next write gets this status, and isn't applied.</summary>
    public HttpStatusCode? FailNextWrite { get; set; }

    /// <summary>The next write is applied, but answered with this status (a 5xx after the work was done).</summary>
    public HttpStatusCode? AppliedButAnswered { get; set; }

    /// <summary>GETs of a single message answer 404 this many times first.</summary>
    public int NotFoundFirst { get; set; }

    /// <summary>Changes the HTML as Teams stores it.</summary>
    public Func<string, string> Store { get; set; } = html => html;

    /// <summary>Drops the mentions as Teams stores them.</summary>
    public bool DropMentions { get; set; }

    [GeneratedRegex(@"^/v1\.0/chats/([^/]+)/messages(?:/(\d+))?$")]
    private static partial Regex Route();

    public GraphClient Client() => new(new GraphClientOptions { Log = _ => { } }, this,
        token: (scopes, _) => { Scopes.Add(scopes); return Task.FromResult("token"); },
        delay: (_, _) => Task.CompletedTask);

    public const string MeId = "me-id";
    public const string MeName = "Me Example";

    public sealed class Chat
    {
        public required string Id { get; init; }
        public string Type { get; init; } = "group";
        public string? Topic { get; init; }
        public List<(string Id, string Name)> Members { get; init; } = [];
    }

    public Dictionary<string, Chat> Chats { get; } = [];

    /// <summary>How many chats /me/chats gives per page.</summary>
    public int ChatPageSize { get; set; } = 50;

    /// <summary>Drops a reply's quote (its attachments) as Teams stores it.</summary>
    public bool DropAttachments { get; set; }

    public Chat AddChat(string id, string type = "group", string? topic = null, params (string Id, string Name)[] others)
    {
        var chat = new Chat { Id = id, Type = type, Topic = topic, Members = [(MeId, MeName), .. others] };
        Chats[id] = chat;
        return chat;
    }

    /// <summary>A message from someone (or from me, with MeId), at a time.</summary>
    public Message Say(string chat, (string Id, string Name) from, string html, string created, JsonArray? mentions = null)
    {
        var message = Add(chat, html, mentions);
        message.From = from.Name;
        message.FromId = from.Id;
        message.Created = created;
        return message;
    }

    public static JsonArray MentionOf(string userId, string name) =>
        [new JsonObject { ["id"] = 0, ["mentionText"] = name, ["mentioned"] = new JsonObject { ["user"] = new JsonObject { ["id"] = userId, ["displayName"] = name } } }];

    public Message Add(string chat, string html, JsonArray? mentions = null)
    {
        var message = new Message { Id = (nextId++).ToString(System.Globalization.CultureInfo.InvariantCulture), Chat = chat,
            Created = "2026-11-17T10:00:00.000Z", Html = html, Mentions = mentions ?? [] };
        Messages.Add(message);
        return message;
    }

    private static HttpResponseMessage Json(HttpStatusCode status, JsonNode? body) =>
        new(status) { Content = new StringContent(body?.ToJsonString() ?? "", Encoding.UTF8, "application/json") };

    private static JsonObject ToJson(Message m) => new()
    {
        ["id"] = m.Id,
        ["createdDateTime"] = m.Created,
        ["deletedDateTime"] = m.Deleted ? "2026-11-17T11:00:00Z" : null,
        ["chatId"] = m.Chat,
        ["messageType"] = m.MessageType,
        ["from"] = From(m),
        ["attachments"] = m.Attachments.DeepClone(),
        ["body"] = new JsonObject { ["contentType"] = "html", ["content"] = m.Html },
        ["mentions"] = m.Mentions.DeepClone(),
    };

    private static JsonObject? From(Message m) =>
        m.MessageType != "message" ? null
        : m.App ? new JsonObject { ["application"] = new JsonObject { ["id"] = m.FromId, ["displayName"] = m.From }, ["device"] = null, ["user"] = null }
        : new JsonObject
        {
            ["application"] = null,
            ["device"] = null,
            ["user"] = new JsonObject
            {
                ["@odata.type"] = "#microsoft.graph.teamworkUserIdentity", ["id"] = m.FromId, ["displayName"] = m.From,
                ["userIdentityType"] = "aadUser", ["tenantId"] = "tenant-id",
            },
        };

    private IEnumerable<Message> Newest(string chat) =>
        Messages.Select((m, i) => (m, i)).Where(x => x.m.Chat == chat)
            .OrderByDescending(x => x.m.Created, StringComparer.Ordinal).ThenByDescending(x => x.i).Select(x => x.m);

    private HttpResponseMessage? Other(HttpRequestMessage request, string path, string query)
    {
        if (request.Method != HttpMethod.Get)
            return null;
        if (path == "/v1.0/me")
            return Json(HttpStatusCode.OK, new JsonObject { ["id"] = MeId, ["displayName"] = MeName });
        if (path == "/v1.0/me/chats")
        {
            var skip = Regex.Match(query, @"\$skiptoken=(\d+)") is { Success: true } s ? int.Parse(s.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) : 0;
            var ordered = Chats.Values.Select(c => (c, last: Newest(c.Id).FirstOrDefault()))
                .OrderByDescending(x => x.last?.Created ?? "", StringComparer.Ordinal).ToList();
            var page = new JsonArray();
            foreach (var (c, last) in ordered.Skip(skip).Take(ChatPageSize))
                page.Add((JsonNode)new JsonObject
                {
                    ["id"] = c.Id, ["topic"] = c.Topic, ["chatType"] = c.Type,
                    ["lastMessagePreview"] = last is null ? null : new JsonObject
                    {
                        ["id"] = last.Id, ["createdDateTime"] = last.Created, ["isDeleted"] = last.Deleted, ["messageType"] = last.MessageType,
                        ["body"] = new JsonObject { ["contentType"] = "html", ["content"] = last.Html }, ["from"] = From(last),
                    },
                });
            var reply = new JsonObject { ["value"] = page };
            if (skip + ChatPageSize < ordered.Count)
                reply["@odata.nextLink"] = $"https://graph.microsoft.com/v1.0/me/chats?$top=50&$skiptoken={skip + ChatPageSize}";
            return Json(HttpStatusCode.OK, reply);
        }
        var chat = ChatRoute().Match(path);
        if (!chat.Success)
            return null;
        if (!Chats.TryGetValue(chat.Groups[1].Value, out var found))
            return Json(HttpStatusCode.NotFound, new JsonObject { ["error"] = "no chat (fake)" });
        if (!chat.Groups[2].Success)
            return Json(HttpStatusCode.OK, new JsonObject { ["id"] = found.Id, ["topic"] = found.Topic, ["chatType"] = found.Type });
        var members = new JsonArray();
        foreach (var (id, name) in found.Members)
            members.Add((JsonNode)new JsonObject { ["userId"] = id, ["displayName"] = name });
        return Json(HttpStatusCode.OK, new JsonObject { ["value"] = members });
    }

    [GeneratedRegex(@"^/v1\.0/chats/([^/]+)(/members)?$")]
    private static partial Regex ChatRoute();

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancel)
    {
        var body = request.Content is null ? null : await request.Content.ReadAsStringAsync(cancel);
        var uri = request.RequestUri!;
        Requests.Add((request.Method.Method, Uri.UnescapeDataString(uri.PathAndQuery), body));
        if (Other(request, Uri.UnescapeDataString(uri.AbsolutePath), Uri.UnescapeDataString(uri.Query)) is { } other)
            return other;
        var route = Route().Match(Uri.UnescapeDataString(uri.AbsolutePath));
        if (!route.Success)
            return Json(HttpStatusCode.NotFound, new JsonObject { ["error"] = "no route (fake)" });
        var chat = route.Groups[1].Value;
        var id = route.Groups[2].Success ? route.Groups[2].Value : null;

        if (request.Method == HttpMethod.Get)
        {
            if (FailNextRead is { } failRead)
            {
                FailNextRead = null;
                return Json(failRead, new JsonObject { ["error"] = "unavailable (fake)" });
            }
            if (id is null)
            {
                var top = int.Parse(Regex.Match(uri.Query, @"\$top=(\d+)").Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                var list = new JsonArray();
                foreach (var m in Newest(chat).Take(top))
                    list.Add((JsonNode)ToJson(m));
                return Json(HttpStatusCode.OK, new JsonObject { ["value"] = list });
            }
            if (NotFoundFirst > 0)
            {
                NotFoundFirst--;
                return Json(HttpStatusCode.NotFound, new JsonObject { ["error"] = "not yet (fake)" });
            }
            var found = Messages.FirstOrDefault(m => m.Chat == chat && m.Id == id);
            return found is null ? Json(HttpStatusCode.NotFound, new JsonObject { ["error"] = "no message (fake)" }) : Json(HttpStatusCode.OK, ToJson(found));
        }

        if (FailNextWrite is { } fail)
        {
            FailNextWrite = null;
            return Json(fail, new JsonObject { ["error"] = "refused (fake)" });
        }
        var payload = JsonNode.Parse(body!)!;
        Message target;
        HttpResponseMessage answer;
        if (request.Method == HttpMethod.Post && id is null)
        {
            target = Add(chat, "");
            target.From = MeName;
            target.FromId = MeId;
            answer = Json(HttpStatusCode.Created, null);
        }
        else if (request.Method == HttpMethod.Patch && id is not null)
        {
            target = Messages.FirstOrDefault(m => m.Chat == chat && m.Id == id)
                ?? throw new InvalidOperationException("PATCH of a message that isn't there (fake)");
            answer = new HttpResponseMessage(HttpStatusCode.NoContent);
        }
        else
            return Json(HttpStatusCode.MethodNotAllowed, null);

        target.Html = Store((string)payload["body"]!["content"]!);
        target.Mentions = DropMentions || payload["mentions"] is not JsonArray mentions ? [] : (JsonArray)mentions.DeepClone();
        target.Attachments = DropAttachments || payload["attachments"] is not JsonArray attachments ? [] : (JsonArray)attachments.DeepClone();
        if (answer.StatusCode == HttpStatusCode.Created)
            answer = Json(HttpStatusCode.Created, ToJson(target));
        if (LoseNextWrite)
        {
            LoseNextWrite = false;
            throw new IOException("connection reset (fake)");
        }
        if (TimeOutNextWrite)
        {
            TimeOutNextWrite = false;
            throw new TaskCanceledException("timed out (fake)", new TimeoutException("the answer never came (fake)"));
        }
        if (AppliedButAnswered is { } status)
        {
            AppliedButAnswered = null;
            return Json(status, new JsonObject { ["error"] = "gateway (fake)" });
        }
        return answer;
    }

    public int Writes => Requests.Count(r => r.Method is "POST" or "PATCH");
}

/// <summary>A temporary folder with settings, a message file and a fake Teams; runs the command line against them.</summary>
public sealed class Env : IDisposable
{
    public const string General = "19:general00000000000000000000000@thread.v2";
    public const string Marketing = "19:marketing000000000000000000000@thread.v2";

    public const string Settings =
        "login_hint = \"me@example.com\"\nstate_file = \"posts.json\"\nlog_file = \"teams-post.log\"\n\n"
        + "[target.marketing]\nchat = \"" + Marketing + "\"\n\n"
        + "[target.general]\nchat = \"" + General + "\"\nmention = \"everyone\"\n";

    public const string Message =
        "---\ndate: 2026-11-16\nleft-out: [press]\n---\nFrom 16/11:\n- **Offer** - PCW promo to 30/11\n  - alt-net too\n- Strapi go-live 24/11\n\n";

    public Env(string settings = Settings)
    {
        Folder = Path.Combine(Path.GetTempPath(), "teams-post-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Folder);
        ConfigPath = Combine("config.toml");
        System.IO.File.WriteAllText(ConfigPath, settings, new UTF8Encoding(false));
        MessageFile = Combine("2026-11-16.md");
        System.IO.File.WriteAllText(MessageFile, Message, new UTF8Encoding(false));
    }

    public string Folder { get; }
    public string ConfigPath { get; }
    public string MessageFile { get; }
    public FakeTeams Teams { get; } = new();
    public StringWriter Out { get; private set; } = new();
    public StringWriter Err { get; private set; } = new();
    public List<TimeSpan> Waits { get; } = [];

    public string Combine(params string[] parts) => Path.Combine([Folder, .. parts]);

    public string? Variable(string name) => name switch
    {
        "APPDATA" => Combine("appdata"),
        "LOCALAPPDATA" => Combine("local"),
        "USERPROFILE" => Folder,
        _ => null,
    };

    public int Run(params string[] args)
    {
        Out = new StringWriter();
        Err = new StringWriter();
        var client = Teams.Client();
        return Cli.Run(["--config", ConfigPath, .. args], new CliHost
        {
            Out = Out,
            Err = Err,
            Env = Variable,
            Graph = _ => client,
            SignIn = _ => { },
            Now = () => new DateTimeOffset(2026, 11, 17, 10, 30, 0, TimeSpan.Zero),
            Wait = Waits.Add,
            RunLock = (_, _) => new Nothing(),
        });
    }

    public List<PostRecord> Records => new PostStore(Combine("posts.json")).Load();

    private sealed class Nothing : IDisposable
    {
        public void Dispose()
        {
        }
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Folder, true);
        }
        catch (IOException)
        {
        }
    }
}
