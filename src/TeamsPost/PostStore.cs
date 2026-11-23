using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace TeamsPost;

/// <summary>One post teams-post made (or tried to). Status "posted": it landed and Id is the message. Status
/// "unknown": the post was sent but no answer came back, so it may be in the chat; Id is null. EditUnknown: the
/// SHA-256 of an edit that was sent but not confirmed (no answer, or a 5xx), so the next edit can tell it landed.</summary>
public sealed record PostRecord(string File, string Target, string Chat, string? Id, string? Created, string Sha256,
    string Status, string PostedAt, string? EditedAt = null, string? EditUnknown = null)
{
    public bool Landed => Status == "posted" && !string.IsNullOrEmpty(Id);
}

/// <summary>The state file: every post, as a JSON array, oldest first. Written through a .tmp beside it, so a crash
/// never leaves half a file.</summary>
public sealed class PostStore(string path)
{
    public string Path { get; } = path;

    public static string Hash(string html) =>
        Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(html)));

    /// <summary>A file's path as the store keys it: absolute and normalised.</summary>
    public static string Normalise(string file) => System.IO.Path.GetFullPath(file);

    public static bool SameFile(string a, string b) =>
        string.Equals(Normalise(a), Normalise(b), StringComparison.OrdinalIgnoreCase);

    public List<PostRecord> Load()
    {
        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(Path);
        }
        catch (Exception error) when (error is FileNotFoundException or DirectoryNotFoundException)
        {
            return [];
        }
        JsonNode? root;
        try
        {
            root = JsonNode.Parse(bytes);
        }
        catch (JsonException error)
        {
            throw new UsageException($"{Path} isn't readable JSON ({error.Message}). Fix or move it; it records what was posted.");
        }
        if (root is not JsonArray items)
            throw new UsageException($"{Path} isn't a JSON array of posts.");
        var records = new List<PostRecord>();
        foreach (var item in items)
        {
            string? Text(string key) => item?[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
            records.Add(new PostRecord(Text("file") ?? "", Text("target") ?? "", Text("chat") ?? "", Text("id"), Text("created"),
                Text("sha256") ?? "", Text("status") ?? "posted", Text("posted_at") ?? "", Text("edited_at"), Text("edit_unknown")));
        }
        return records;
    }

    public void Save(IEnumerable<PostRecord> records)
    {
        var array = new JsonArray();
        foreach (var r in records)
        {
            var item = new JsonObject
            {
                ["file"] = r.File,
                ["target"] = r.Target,
                ["chat"] = r.Chat,
                ["id"] = r.Id,
                ["created"] = r.Created,
                ["sha256"] = r.Sha256,
                ["status"] = r.Status,
                ["posted_at"] = r.PostedAt,
            };
            if (r.EditedAt is not null)
                item["edited_at"] = r.EditedAt;
            if (r.EditUnknown is not null)
                item["edit_unknown"] = r.EditUnknown;
            array.Add((JsonNode)item);
        }
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(Path))!);
        var temp = Path + ".tmp";
        using (var stream = File.Create(temp))
        {
            using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }))
                array.WriteTo(writer);
            stream.WriteByte((byte)'\n');
        }
        File.Move(temp, Path, overwrite: true);
    }

    public void Add(PostRecord record)
    {
        var records = Load();
        records.Add(record);
        Save(records);
    }

    /// <summary>Replace the record for this message (by chat and id), or add it.</summary>
    public void Upsert(PostRecord record)
    {
        var records = Load();
        var at = records.FindIndex(r => r.Id == record.Id && r.Chat == record.Chat && r.Id is not null);
        if (at >= 0)
            records[at] = record;
        else
            records.Add(record);
        Save(records);
    }
}
