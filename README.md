# teams-post

Posts a markdown file to a Microsoft Teams chat, reads it back to check it landed, and can edit the
message later from the same file.

```
teams-post post marketing 2026-11-16.md          # shows the message as it will appear; nothing is sent
teams-post post marketing 2026-11-16.md --yes
  Posted 1763371234567 at 2026-11-17T10:30:01.123Z.
  Read back OK.
teams-post edit 2026-11-16.md                    # shows what would change in that message
teams-post edit 2026-11-16.md --yes
teams-post show marketing --last 3
```

It's for messages someone has written and approved as a file (a team update, a reminder with
@Everyone), where the file is the record and a person decides when it goes.

## Install

Windows (x64) only; nothing else is needed.

1. Download `teams-post-<version>-win-x64.zip` from [Releases](https://github.com/gareth-jones-zen/teams-post/releases).
2. Unzip it to `%LOCALAPPDATA%\Programs\teams-post`.
3. Run `teams-post config` to create the settings file, fill in your targets, then `teams-post sign-in`.

To upgrade, unzip the newer release over the same folder.

## Permissions

It signs in as you through Microsoft's "Graph Command Line Tools" app (delegated permissions), and
asks for each permission only when a command needs it:

| Command | Graph call | Permission |
|---|---|---|
| `post --yes` | `POST /chats/{chat}/messages` | **ChatMessage.Send** |
| `post --yes`, `edit`, `show` (reading) | `GET /chats/{chat}/messages[/{id}]` | **Chat.Read** (or Chat.ReadWrite) |
| `edit --yes` | `PATCH /chats/{chat}/messages/{id}` | **Chat.ReadWrite** |

`sign-in` asks for ChatMessage.Send and Chat.Read. Chat.ReadWrite is asked for the first time you
edit, because it's the only delegated permission Graph accepts for changing a chat message. Many
organisations need an admin to approve one or more of these: the sign-in shows an "Approval
required" screen, where you can ask. Only the person who posted a message can edit it.

## Sharing settings with other tools: `base`

Several tools built on the shared library can keep their common settings in one file, named at the top of each
tool's settings: `base = "base-config.toml"` (relative to the settings file). Then:

- **the base's `[paths]` are variables** for this tool's paths, looked up before the environment: with
  `EA_LOCAL = "%LOCALAPPDATA%/claude-ea"` there, `state_file = "%EA_LOCAL%/teams-post.json"` works here;
- **its other settings that teams-post knows and this file doesn't set are used**, typically `login_hint`, `tenant`
  and `msal_cache`;
- **a table such as `[target]` applies only if this file has that section too**, its own entries winning.

The loader is background-tools-lib's `PyCompat.BaseSettings`.

## Settings

`%APPDATA%\teams-post\config.toml` by default (`--config <file>`, before the command, or
`TEAMS_POST_CONFIG` for another).

```toml
base = "base-config.toml"   # optional: settings shared with other tools
login_hint = "you@example.com"
tenant = "organizations"            # or a tenant id
msal_cache = "..."                  # optional: share another tool's sign-in cache
state_file = "..."                  # optional; default %LOCALAPPDATA%\teams-post\posts.json
log_file = "..."                    # optional; default %LOCALAPPDATA%\teams-post\teams-post.log
draft_marker = "[DRAFT]"            # a file containing this is refused ("" for no check)
strip_frontmatter = true            # a leading --- YAML --- block is ignored

[target.marketing]
chat = "19:...@thread.v2"
[target.general]
chat = "19:...@thread.v2"
mention = "everyone"                # optional: the post starts with @Everyone
```

A chat's id is in its link in Teams (`.../l/chat/19:...@thread.v2/...`), or Graph's `/me/chats`.
Paths may use `~` and `%VARIABLES%`; a relative path is relative to the settings file.

**The file is read strictly:** an unknown setting, a target with no chat, a `mention` other than
`"everyone"`, a value of the wrong type or a key given twice is an error that names the setting,
not something silently ignored.

**Channels are planned.** A target can already be written as
`channel = { team = "<team id>", channel = "<channel id>" }`, but posting to it says it isn't
supported yet: it needs ChannelMessage.Send, which isn't in use yet.

## Commands

```
teams-post config                           create the settings file (if missing) and say where it is
teams-post sign-in                          sign in to Microsoft (once; opens a browser)
teams-post post <target> <file.md>          show what would be posted (sends nothing)
teams-post post <target> <file.md> --yes    post it, then read it back
teams-post edit <file.md>                   show what would change in the message it was posted as
teams-post edit <file.md> --yes             edit that message, then read it back
teams-post show <target> [<message id>]     a message, or the newest ones (--last N, default 5, up to 50)
teams-post --version

    --again                  post: post a file that has already been posted to that target
    --target T               edit: which target, when the file went to more than one
    --id MESSAGE_ID          edit: a message by its id (with --target), when there's no record of it
```

**Exit codes:** 0 done; 1 refused or failed (a draft, a second post, a read-back that differs, an
answer that never came); 2 settings, arguments or sign-in.

## The markdown it understands

A deliberate small subset, so what you see in the file is what the chat shows:

| In the file | In Teams |
|---|---|
| Each line of text (blank lines between them or not) | A paragraph |
| `# Heading` (any level) | A bold paragraph |
| `- item` or `* item` | A bulleted list |
| `1. item` (or `1)`) | A numbered list |
| Two more spaces of indent (a tab counts as two) | One level of nesting, inside the item above |
| An indented line under a bullet that isn't a bullet | Carries on that bullet |
| `**bold**`, `*italic*` or `_italic_`, `` `code` `` | Bold, italic, code |
| `[text](https://...)` | A link (`http`, `https` and `mailto` only; anything else shows as typed) |
| A bare URL | Left as text: Teams makes it a link |

- **Everything else is HTML-escaped**, so `<`, `>` and `&` show as typed.
- A blank line between bullets doesn't end the list; a line of plain text does. Switching between
  bullets and numbers at the same level starts a new list.
- An indent deeper than one level below the bullet above nests one level, not more.
- `snake_case_names` and `2 * 3 * 4` aren't italics: emphasis needs a non-word character either side.
- **Frontmatter** (a leading `---` ... `---` block) is dropped when `strip_frontmatter` is on, so
  notes for yourself can sit there. Trailing blank lines are dropped.
- **The draft marker anywhere in the file, frontmatter included, refuses it**, before anything is
  sent: take the marker off once the message is approved.

**@Everyone** (a target with `mention = "everyone"`): the message starts with
`<at id="0">Everyone</at>` followed by a space, in the first paragraph (or in a paragraph of its own
if the message starts with a list), with the mention of the whole chat:

```json
"mentions": [{ "id": 0, "mentionText": "Everyone",
  "mentioned": { "conversation": { "id": "<chat id>", "displayName": "Everyone", "conversationIdentityType": "chat" } } }]
```

That's exactly how Teams stores an @Everyone typed by hand, so everyone in the chat is notified.

## Safety

- **Preview by default.** Without `--yes`, `post` shows the target and the message as plain text
  (what the chat will show), and `edit` a line diff of the message in Teams against the file.
  Nothing is sent.
- **A file is posted once per target.** Every post is recorded in the state file (the file's full
  path, target, chat, message id, created time and a SHA-256 of the HTML sent). Posting the same file
  to the same target again is refused, with a pointer to `edit`; `--again` posts it anyway. Two runs
  at once wait for each other, so they can't both decide a file hasn't been posted.
- **A change is never sent twice.** A post or an edit is sent again only after a 429 (refused before
  it was looked at) or a connection that never opened. After a timeout, a dropped connection or a
  5xx it may have been applied, so it isn't resent: the run stops, says so, and exits 1. A post whose
  answer was lost is recorded as `unknown`, so the next `post` of that file is refused until you've
  looked (`show`). An edit whose answer was lost leaves its record as it was; `show <target> <id>`
  tells you whether it landed. Reads are retried as usual.
- **Every post and edit is read back.** The message must have the same number of list items and
  paragraphs as was sent, and the @Everyone mention when one was sent; otherwise it says what differs
  and exits 1. A message just posted can answer "not found" for a moment, so the read waits a few
  seconds for it.
- **The log** (`log_file`) records each post and edit by target and message id, not the text.

## Building

It's .NET 11, published as Native AOT. It needs [background-tools-lib](https://github.com/gareth-jones-zen/background-tools-lib)
checked out beside this repository (or set the `BackgroundToolsLib` property to a checkout elsewhere), then:

```
dotnet test -c Release
dotnet publish src/TeamsPost.Console -c Release -r win-x64 -o out
```

The tests run the real Graph client against a fake Teams behind a fake HTTP handler, which can lose
an answer, fail a write or store a message differently from how it was sent. A tag `v*` builds and
publishes a release.
