using System.Text.Json.Nodes;

namespace TeamsPost.Tests;

public class MarkdownTests
{
    [Fact]
    public void Lines_are_paragraphs()
    {
        Assert.Equal("<p>one</p><p>two</p><p>three</p>", TeamsMarkdown.ToHtml("one\ntwo\n\nthree"));
    }

    [Fact]
    public void Bullets_are_a_list_with_either_marker()
    {
        Assert.Equal("<ul><li>a</li><li>b</li></ul>", TeamsMarkdown.ToHtml("- a\n* b"));
    }

    [Fact]
    public void Numbers_are_an_ordered_list()
    {
        Assert.Equal("<ol><li>first</li><li>second</li></ol>", TeamsMarkdown.ToHtml("1. first\n2. second"));
    }

    [Fact]
    public void Two_spaces_nest_inside_the_item()
    {
        Assert.Equal("<ul><li>a<ul><li>b<ul><li>c</li></ul></li></ul></li><li>d</li></ul>",
            TeamsMarkdown.ToHtml("- a\n  - b\n    - c\n- d"));
    }

    [Fact]
    public void A_tab_is_two_spaces_and_a_deep_jump_nests_one_level()
    {
        // a tab is one level; "c" at six spaces (level 3) goes one level under "b", not two
        Assert.Equal("<ul><li>a<ul><li>b<ul><li>c</li></ul></li></ul></li></ul>", TeamsMarkdown.ToHtml("- a\n\t- b\n      - c"));
        // a first bullet indented by mistake is still top level
        Assert.Equal("<ul><li>x</li></ul>", TeamsMarkdown.ToHtml("    - x"));
    }

    [Fact]
    public void A_numbered_list_can_nest_in_bullets()
    {
        Assert.Equal("<ul><li>a<ol><li>x</li><li>y</li></ol></li></ul>", TeamsMarkdown.ToHtml("- a\n  1. x\n  2. y"));
    }

    [Fact]
    public void Changing_list_type_starts_a_new_list()
    {
        Assert.Equal("<ul><li>a</li></ul><ol><li>b</li></ol>", TeamsMarkdown.ToHtml("- a\n1. b"));
    }

    [Fact]
    public void A_plain_line_ends_a_list_a_blank_line_doesnt()
    {
        Assert.Equal("<ul><li>a</li><li>b</li></ul><p>after</p><ul><li>c</li></ul>", TeamsMarkdown.ToHtml("- a\n\n- b\nafter\n- c"));
    }

    [Fact]
    public void An_indented_line_carries_on_its_bullet()
    {
        Assert.Equal("<ul><li>one wrapped</li></ul>", TeamsMarkdown.ToHtml("- one\n  wrapped"));
    }

    [Fact]
    public void A_heading_is_a_bold_paragraph()
    {
        Assert.Equal("<p><strong>Notes</strong></p><p>x</p>", TeamsMarkdown.ToHtml("## Notes\nx"));
    }

    [Theory]
    [InlineData("**bold**", "<strong>bold</strong>")]
    [InlineData("*italic*", "<em>italic</em>")]
    [InlineData("_italic_", "<em>italic</em>")]
    [InlineData("`a < b`", "<code>a &lt; b</code>")]
    [InlineData("**bold with *italic* in**", "<strong>bold with <em>italic</em> in</strong>")]
    [InlineData("[the form](https://forms.example.com/r?a=1&b=2)", "<a href=\"https://forms.example.com/r?a=1&amp;b=2\">the form</a>")]
    [InlineData("[**bold** link](https://x.example)", "<a href=\"https://x.example\"><strong>bold</strong> link</a>")]
    [InlineData("[bad](javascript:alert(1))", "[bad](javascript:alert(1))")]
    [InlineData("see https://example.com/a_b_c here", "see https://example.com/a_b_c here")]
    [InlineData("snake_case_name stays", "snake_case_name stays")]
    [InlineData("2 * 3 * 4", "2 * 3 * 4")]
    [InlineData("** not bold **", "** not bold **")]
    [InlineData("Fish & chips <b>not bold</b> \"q\"", "Fish &amp; chips &lt;b&gt;not bold&lt;/b&gt; \"q\"")]
    [InlineData("`**not bold**`", "<code>**not bold**</code>")]
    public void Inline_markup(string markdown, string html) => Assert.Equal(html, TeamsMarkdown.Inline(markdown));

    [Fact]
    public void Frontmatter_is_dropped_and_trailing_blanks_too()
    {
        var text = TeamsMarkdown.MessageText("---\ntitle: x\nleft-out: y\n---\n\nHello\n- a\n\n\n", "f.md", "[DRAFT]", true);
        Assert.Equal("Hello\n- a", text);
    }

    [Fact]
    public void Frontmatter_is_kept_when_not_stripping_and_an_unclosed_block_is_kept()
    {
        Assert.Equal("---\na: b\n---\nx", TeamsMarkdown.MessageText("---\na: b\n---\nx", "f.md", "[DRAFT]", false));
        Assert.Equal("---\nno end", TeamsMarkdown.StripFrontmatter("---\nno end"));
    }

    [Fact]
    public void Crlf_and_a_byte_order_mark_are_read()
    {
        Assert.Equal("Hello\nWorld", TeamsMarkdown.MessageText("﻿---\r\na: 1\r\n---\r\nHello\r\nWorld\r\n", "f.md", "[DRAFT]", true));
    }

    [Fact]
    public void The_draft_marker_anywhere_is_refused()
    {
        var error = Assert.Throws<RefusedException>(() => TeamsMarkdown.MessageText("---\nstatus: x\n---\n# [DRAFT] hi\n", "f.md", "[DRAFT]", true));
        Assert.Contains("still marked [DRAFT]", error.Message);
        Assert.Throws<RefusedException>(() => TeamsMarkdown.MessageText("---\ntitle: [DRAFT]\n---\nhi\n", "f.md", "[DRAFT]", true));
    }

    [Fact]
    public void An_empty_draft_marker_checks_nothing() =>
        Assert.Equal("[DRAFT] hi", TeamsMarkdown.MessageText("[DRAFT] hi", "f.md", "", true));

    [Fact]
    public void A_file_with_no_message_is_refused() =>
        Assert.Contains("no message", Assert.Throws<RefusedException>(() => TeamsMarkdown.MessageText("---\na: 1\n---\n\n\n", "f.md", "[DRAFT]", true)).Message);

    [Fact]
    public void Everyone_goes_at_the_start_of_the_first_paragraph()
    {
        Assert.Equal("<p><at id=\"0\">Everyone</at> can you fill this in please</p><ul><li>x</li></ul>",
            TeamsMarkdown.HtmlFor("can you fill this in please\n- x", Mention.Everyone));
    }

    [Fact]
    public void Everyone_gets_its_own_paragraph_before_a_list() =>
        Assert.Equal("<p><at id=\"0\">Everyone</at></p><ul><li>x</li></ul>", TeamsMarkdown.HtmlFor("- x", Mention.Everyone));

    [Fact]
    public void The_payload_is_html_with_the_chat_mention()
    {
        var payload = TeamsMarkdown.Payload("<p><at id=\"0\">Everyone</at> hi</p>", Mention.Everyone, "19:abc@thread.v2");
        Assert.Equal(
            "{\"body\":{\"contentType\":\"html\",\"content\":\"\\u003Cp\\u003E\\u003Cat id=\\u00220\\u0022\\u003EEveryone\\u003C/at\\u003E hi\\u003C/p\\u003E\"},"
            + "\"mentions\":[{\"id\":0,\"mentionText\":\"Everyone\",\"mentioned\":{\"conversation\":{\"id\":\"19:abc@thread.v2\","
            + "\"displayName\":\"Everyone\",\"conversationIdentityType\":\"chat\"}}}]}",
            payload.ToJsonString());
        var back = JsonNode.Parse(payload.ToJsonString())!;
        Assert.Equal("<p><at id=\"0\">Everyone</at> hi</p>", (string?)back["body"]!["content"]);
    }

    [Fact]
    public void No_mention_no_mentions_array() =>
        Assert.Null(TeamsMarkdown.Payload("<p>hi</p>", Mention.None, "19:abc@thread.v2")["mentions"]);
}

public class PlainTextTests
{
    [Fact]
    public void Paragraphs_lists_and_nesting()
    {
        var html = TeamsMarkdown.ToHtml("Intro\n- a\n  - b\n  1. c\n- **d** and [link](https://x.example)\n1. one\n2. two");
        Assert.Equal(["Intro", "- a", "  - b", "  1. c", "- d and link <https://x.example>", "1. one", "2. two"], PlainText.Lines(html));
    }

    [Fact]
    public void Teams_markup_reads_as_plain_text()
    {
        var html = "<div><div><p>Hi&nbsp;<at id=\"0\">Everyone</at>,\n  look <emoji id=\"1f440_eyes\" alt=\"👀\" title=\"Eyes\"></emoji></p>"
                   + "<p>A &amp; B<br>next</p><img src=\"x\"><a href=\"https://same.example\">https://same.example</a></div></div>";
        Assert.Equal(["Hi @Everyone, look 👀", "A & B", "next", "[image]https://same.example"], PlainText.Lines(html));
    }

    [Fact]
    public void Escaped_text_comes_back_as_typed() =>
        Assert.Equal(["Fish & chips <b>"], PlainText.Lines(TeamsMarkdown.ToHtml("Fish & chips <b>")));
}
