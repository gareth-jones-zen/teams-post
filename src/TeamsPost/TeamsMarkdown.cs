using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace TeamsPost;

/// <summary>A request that's refused before anything is sent (a draft, an empty file, a second post): exit 1.</summary>
public sealed class RefusedException(string message) : Exception(message);

/// <summary>A small, deliberate subset of markdown, as Teams HTML:
/// - each non-blank line that isn't a bullet is a paragraph (a "# heading" is a bold paragraph);
/// - "- " or "* " bullets are a list, "1. " a numbered list; two spaces of indent nest one level (a tab counts as
///   two); an indented line under a bullet that isn't a bullet carries on that bullet; a blank line between bullets
///   doesn't end the list, a plain line does;
/// - inline: **bold**, *italic* or _italic_, `code`, [text](https://...); a bare URL is left as text (Teams links it);
/// - everything else is HTML-escaped, so it shows as typed.</summary>
public static partial class TeamsMarkdown
{
    /// <summary>The @Everyone mention, as Teams stores it when someone types it in a group chat.</summary>
    public const string EveryoneTag = "<at id=\"0\">Everyone</at>";

    [GeneratedRegex(@"^( *)([-*]|\d{1,9}[.)]) +(.*)$", RegexOptions.CultureInvariant)]
    private static partial Regex ListItem();

    [GeneratedRegex(@"^#{1,6} +(.*?)( +#+)? *$", RegexOptions.CultureInvariant)]
    private static partial Regex Heading();

    // 1 code; 2, 3 link text and url; 4 bold; 5 *italic*; 6 _italic_.
    [GeneratedRegex(@"`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\)|\*\*(?!\s)(.+?)(?<!\s)\*\*|(?<![\w*])\*(?![\s*])(.+?)(?<![\s*])\*(?![\w*])|(?<![\w_])_(?![\s_])(.+?)(?<![\s_])_(?![\w_])",
        RegexOptions.CultureInvariant)]
    private static partial Regex InlinePattern();

    /// <summary>The message part of a file's text: refused if it holds the draft marker (anywhere, frontmatter
    /// included); a leading frontmatter block dropped if asked; trailing blank lines dropped. Throws
    /// RefusedException for a draft or a file with no message.</summary>
    public static string MessageText(string text, string file, string draftMarker, bool stripFrontmatter)
    {
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (text.StartsWith('﻿'))
            text = text[1..];
        if (draftMarker.Length > 0 && text.Contains(draftMarker, StringComparison.Ordinal))
            throw new RefusedException($"{file} is still marked {draftMarker}: take the marker off once it's approved.");
        if (stripFrontmatter)
            text = StripFrontmatter(text);
        var lines = text.Split('\n').ToList();
        while (lines.Count > 0 && lines[^1].Trim().Length == 0)
            lines.RemoveAt(lines.Count - 1);
        while (lines.Count > 0 && lines[0].Trim().Length == 0)
            lines.RemoveAt(0);
        if (lines.Count == 0)
            throw new RefusedException($"{file} has no message in it.");
        return string.Join('\n', lines);
    }

    /// <summary>Text without a leading "---" ... "---" (or "...") block. An unclosed block is left as it is.</summary>
    public static string StripFrontmatter(string text)
    {
        var lines = text.Split('\n');
        if (lines.Length == 0 || lines[0].TrimEnd() != "---")
            return text;
        for (var i = 1; i < lines.Length; i++)
            if (lines[i].TrimEnd() is "---" or "...")
                return string.Join('\n', lines[(i + 1)..]);
        return text;
    }

    /// <summary>The message's HTML.</summary>
    public static string ToHtml(string markdown)
    {
        var html = new StringBuilder();
        var open = new List<string>();   // the open lists, outermost first; each has an <li> open

        void CloseTo(int depth)
        {
            while (open.Count > depth)
            {
                html.Append("</li></").Append(open[^1]).Append('>');
                open.RemoveAt(open.Count - 1);
            }
        }

        foreach (var raw in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            var line = ExpandLeadingTabs(raw);
            if (line.Trim().Length == 0)
                continue;
            var item = ListItem().Match(line);
            if (item.Success)
            {
                var tag = char.IsAsciiDigit(item.Groups[2].Value[0]) ? "ol" : "ul";
                var depth = Math.Min(item.Groups[1].Length / 2, open.Count);
                CloseTo(depth + 1);
                if (open.Count == depth + 1 && open[^1] != tag)
                    CloseTo(depth);
                if (open.Count == depth + 1)
                    html.Append("</li><li>");
                else
                {
                    html.Append('<').Append(tag).Append("><li>");
                    open.Add(tag);
                }
                html.Append(Inline(item.Groups[3].Value.Trim()));
                continue;
            }
            if (open.Count > 0 && char.IsWhiteSpace(line[0]))
            {
                html.Append(' ').Append(Inline(line.Trim()));   // a wrapped bullet
                continue;
            }
            CloseTo(0);
            var heading = Heading().Match(line);
            html.Append("<p>")
                .Append(heading.Success ? "<strong>" + Inline(heading.Groups[1].Value) + "</strong>" : Inline(line.Trim()))
                .Append("</p>");
        }
        CloseTo(0);
        return html.ToString();
    }

    private static string ExpandLeadingTabs(string line)
    {
        var i = 0;
        var indent = new StringBuilder();
        for (; i < line.Length && line[i] is ' ' or '\t'; i++)
            indent.Append(line[i] == '\t' ? "  " : " ");
        return indent.Append(line, i, line.Length - i).ToString();
    }

    /// <summary>One line's inline markup as HTML, everything else escaped.</summary>
    public static string Inline(string text)
    {
        var html = new StringBuilder();
        var at = 0;
        foreach (Match m in InlinePattern().Matches(text))
        {
            html.Append(Escape(text[at..m.Index]));
            at = m.Index + m.Length;
            if (m.Groups[1].Success)
                html.Append("<code>").Append(Escape(m.Groups[1].Value)).Append("</code>");
            else if (m.Groups[2].Success)
            {
                var url = m.Groups[3].Value;
                if (url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || url.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                    || url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                    html.Append("<a href=\"").Append(Escape(url, attribute: true)).Append("\">").Append(Inline(m.Groups[2].Value)).Append("</a>");
                else
                    html.Append(Escape(m.Value));   // not a web link: shown as typed
            }
            else if (m.Groups[4].Success)
                html.Append("<strong>").Append(Inline(m.Groups[4].Value)).Append("</strong>");
            else
                html.Append("<em>").Append(Inline(m.Groups[5].Success ? m.Groups[5].Value : m.Groups[6].Value)).Append("</em>");
        }
        return html.Append(Escape(text[at..])).ToString();
    }

    public static string Escape(string text, bool attribute = false)
    {
        var escaped = text.Replace("&", "&amp;", StringComparison.Ordinal).Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal);
        return attribute ? escaped.Replace("\"", "&quot;", StringComparison.Ordinal) : escaped;
    }

    /// <summary>The HTML with @Everyone at the start of the first paragraph (followed by a space), or in a
    /// paragraph of its own when the message starts with a list.</summary>
    public static string WithEveryone(string html) =>
        html.StartsWith("<p>", StringComparison.Ordinal)
            ? "<p>" + EveryoneTag + " " + html[3..]
            : "<p>" + EveryoneTag + "</p>" + html;

    /// <summary>The message's HTML for a target: with the mention, if the target has one.</summary>
    public static string HtmlFor(string markdown, Mention mention)
    {
        var html = ToHtml(markdown);
        return mention == Mention.Everyone ? WithEveryone(html) : html;
    }

    /// <summary>The JSON Graph is sent, for a new message or an edit: the HTML body, and for @Everyone the mention
    /// of the whole chat, as Teams stores it when someone types @Everyone.</summary>
    public static JsonObject Payload(string html, Mention mention, string chatId)
    {
        var payload = new JsonObject
        {
            ["body"] = new JsonObject { ["contentType"] = "html", ["content"] = html },
        };
        if (mention == Mention.Everyone)
        {
            payload["mentions"] = new JsonArray(new JsonObject
            {
                ["id"] = 0,
                ["mentionText"] = "Everyone",
                ["mentioned"] = new JsonObject
                {
                    ["conversation"] = new JsonObject
                    {
                        ["id"] = chatId,
                        ["displayName"] = "Everyone",
                        ["conversationIdentityType"] = "chat",
                    },
                },
            });
        }
        return payload;
    }
}
