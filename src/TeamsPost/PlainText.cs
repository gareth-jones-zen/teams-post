using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace TeamsPost;

/// <summary>A Teams message's HTML as plain lines, for the preview, show and the edit diff: one line per paragraph,
/// "- " or "1. " per list item (indented two spaces per level), @ before a mention, a link's address after its text
/// when they differ. Bold, italics and code lose their markup. Used on what teams-post sends and on what Teams
/// returns, so a preview and a read-back compare like with like.</summary>
public static partial class PlainText
{
    [GeneratedRegex(@"<(/?)([A-Za-z][A-Za-z0-9]*)([^>]*)>|([^<]+)|<", RegexOptions.CultureInvariant)]
    private static partial Regex Tokens();

    [GeneratedRegex(@"\b(href|alt)\s*=\s*(?:""([^""]*)""|'([^']*)')", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex Attribute();

    [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant)]
    private static partial Regex Spaces();

    private sealed class ListState(bool ordered)
    {
        public bool Ordered { get; } = ordered;
        public int Count { get; set; }
    }

    public static IReadOnlyList<string> Lines(string html)
    {
        var lines = new List<string>();
        var line = new StringBuilder();
        var lists = new List<ListState>();
        var links = new Stack<(string? Href, int Start)>();

        void Break()
        {
            var text = line.ToString().Trim();
            if (text.Length > 0)
                lines.Add(new string(' ', Math.Max(0, lists.Count - 1) * 2) + text);
            line.Clear();
        }

        foreach (Match m in Tokens().Matches(html))
        {
            if (m.Groups[4].Success || m.Value == "<")
            {
                var text = Spaces().Replace(WebUtility.HtmlDecode(m.Value).Replace(' ', ' '), " ");
                if (line.Length == 0 || line[^1] == ' ')
                    text = text.TrimStart();
                line.Append(text);
                continue;
            }
            var closing = m.Groups[1].Value == "/";
            var tag = m.Groups[2].Value.ToLowerInvariant();
            var attributes = m.Groups[3].Value;
            switch (tag, closing)
            {
                case ("p" or "div" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "blockquote" or "pre" or "table" or "tr", _):
                case ("br", false):
                    Break();
                    break;
                case ("ul" or "ol", false):
                    Break();
                    lists.Add(new ListState(tag == "ol"));
                    break;
                case ("ul" or "ol", true):
                    Break();
                    if (lists.Count > 0)
                        lists.RemoveAt(lists.Count - 1);
                    break;
                case ("li", false):
                    Break();
                    if (lists.Count > 0)
                    {
                        var list = lists[^1];
                        list.Count++;
                        line.Append(list.Ordered ? $"{list.Count}. " : "- ");
                    }
                    break;
                case ("li", true):
                    Break();
                    break;
                case ("at", false):
                    line.Append('@');
                    break;
                case ("a", false):
                    links.Push((AttributeValue(attributes, "href"), line.Length));
                    break;
                case ("a", true):
                    if (links.Count > 0)
                    {
                        var (href, start) = links.Pop();
                        var shown = start <= line.Length ? line.ToString(start, line.Length - start).Trim() : "";
                        if (!string.IsNullOrEmpty(href) && shown != href)
                            line.Append(" <").Append(href).Append('>');
                    }
                    break;
                case ("emoji", false):
                    line.Append(AttributeValue(attributes, "alt") ?? "");
                    break;
                case ("img", false):
                    line.Append("[image]");
                    break;
                case ("attachment", false):
                    line.Append("[attachment]");
                    break;
            }
        }
        Break();
        return lines;
    }

    public static string Of(string html) => string.Join('\n', Lines(html));

    private static string? AttributeValue(string attributes, string name)
    {
        foreach (Match m in Attribute().Matches(attributes))
            if (string.Equals(m.Groups[1].Value, name, StringComparison.OrdinalIgnoreCase))
                return WebUtility.HtmlDecode(m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value);
        return null;
    }
}
