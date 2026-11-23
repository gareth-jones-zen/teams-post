# teams-post

Posts a markdown file to a Microsoft Teams chat, reads it back to check it landed, and can edit the
message later from the same file. It also lists the chats where someone else spoke last, and replies
in any of your chats from a file, quoting the message it answers.

```
teams-post post marketing 2026-11-16.md          # shows the message as it will appear; nothing is sent
teams-post post marketing 2026-11-16.md --yes
  Posted 1763371234567 at 2026-11-17T10:30:01.123Z.
  Read back OK.
teams-post edit 2026-11-16.md                    # shows what would change in that message
teams-post edit 2026-11-16.md --yes
teams-post show marketing --last 3
teams-post unanswered --hours 72                 # chats waiting on you (reads only)
teams-post reply "19:...@unq.gbl.spaces" reply.md --to 1763371234567        # preview
teams-post reply "19:...@unq.gbl.spaces" reply.md --to 1763371234567 --yes
```

It's for messages someone has written and approved as a file (a team update, a reminder with
@Everyone, a reply drafted for someone to check), where the file is the record and a person decides
when it goes.

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
| `reply --yes` | `POST /chats/{chat}/messages` | **ChatMessage.Send** |
| `unanswered`, `reply` (reading) | `GET /me/chats`, `GET /chats/{chat}[/members]`, `GET /chats/{chat}/messages` | **Chat.Read** |
| `unanswered`, `reply` (who you are) | `GET /me` | **User.Read** |

`sign-in` asks for ChatMessage.Send, Chat.Read and User.Read. Chat.ReadWrite is asked for the first time you
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

[unanswered]                        # optional: what `unanswered` looks at
hours = 48                          # how far back (the default); or
# working_days = 2                  # that many working days back (Monday to Friday), not both
skip_chats = ["19:...@thread.v2", "^Team .* General$"]   # chat ids, or regular expressions (case ignored)
                                    # matched against the chat's name: its topic, or the other person's name
skip_senders = ["Workflows", "Some Person"]              # display names (case ignored) or user ids
```

`working_days` keeps a Monday morning from missing Friday afternoon: counted back from now, a
Saturday or Sunday doesn't count (bank holidays do).

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
teams-post show <target> [<message id>]     a message, or the newest ones (--last N, default 5, up to 50);
                                            a chat id works in place of a target
teams-post unanswered [--hours N] [--json]  chats where someone else spoke last (reads only)
teams-post reply <chat id> <file.md>        show the reply that would be sent (sends nothing)
teams-post reply <chat id> <file.md> --yes  send it into that chat, then read it back
teams-post --version

    --again                  post, reply: send a file that has already been sent there
    --target T               edit: which target, when the file went to more than one
    --id MESSAGE_ID          edit: a message by its id (with --target), when there's no record of it
    --to MESSAGE_ID          reply: quote that message, as Teams' own Reply does
    --hours N                unanswered: look back N hours (1 to 744), instead of [unanswered]
    --json                   unanswered: as JSON
```

**Exit codes:** 0 done; 1 refused or failed (a draft, a second post, a read-back that differs, an
answer that never came, a chat you aren't in); 2 settings, arguments or sign-in.

## Chats waiting on you: `unanswered`

Lists your one-to-one and group chats where the newest message a person wrote is someone else's
and falls inside the window (48 hours, or `[unanswered]`), newest first. It reads only. It doesn't
decide what needs an answer: it lists them all, flagged, for you (or an assistant) to judge.

- **Not counted:** system events (someone added, a meeting started), messages from apps and bots,
  deleted messages, and anyone in `skip_senders`. If one of those is the newest, the message before it decides.
- **Meeting chats** are listed only when you're @mentioned in the window. Channels aren't looked at.
- **Flags:** `direct` (a one-to-one), `mentionsYou` (you're @mentioned), `question` (a "?", not
  counting one inside a link).
- **For each chat:** its name (topic, or the other person), type, id, the newest five of the
  messages since yours (oldest first, each as plain text, up to 500 characters), and the id of the
  newest, for `reply --to`.

```
Since Fri 13 Nov 10:30 (2 working days): 2 chats where someone else spoke last.

Alex Example  (one-to-one; direct, question)
  chat 19:..._...@unq.gbl.spaces
  reply --to 1763371234567
  Tue 17 Nov 08:55  Alex Example
    Have you seen the design?
  Tue 17 Nov 09:00  Alex Example
    Can you look at this <https://example.com/doc> today?

Design review  (group chat; mentionsYou)
  ...
```

`--json` gives the same as one object: `since`, `window`, `chats` (each with `chat`, `title`,
`chatType`, `flags` {`direct`, `mentionsYou`, `question`}, `replyTo`, `lastAt`, `since` (how many
messages since yours) and `messages` [{`id`, `from`, `created`, `text`}]), and `unreadable` (chats
whose messages couldn't be read).

## Replying in a chat: `reply`

Sends a markdown file (the same subset as `post`) into any chat you're in, by its id: a target isn't
needed. With `--to`, the reply quotes that message the way Teams' own Reply does in a chat: the body
starts with `<attachment id="<message id>"></attachment>`, with an attachment

```json
{ "id": "<message id>", "contentType": "messageReference",
  "content": "{\"messageId\":\"<message id>\",\"messagePreview\":\"<its first 200 characters>\",\"messageSender\":{\"application\":null,\"device\":null,\"user\":{\"userIdentityType\":\"aadUser\",\"tenantId\":\"...\",\"id\":\"<sender's id>\",\"displayName\":\"<sender>\"}}}" }
```

That's how Teams stores a reply typed by hand.

- **The preview** shows the chat (its topic or the people in it), the message being answered
  (sender, time, its first 200 characters) and the reply as plain text.
- **Refused before anything is sent:** a channel's id (`@thread.tacv2`, `@thread.skype`), a chat
  you can't read or aren't a member of, a `--to` message that isn't in the chat, is deleted or is a
  system event, and a draft.
- **A reply can't be edited by teams-post yet** (`edit` says so): edit it in Teams.

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

- **Preview by default.** Without `--yes`, `post` and `reply` show where it's going and the message
  as plain text (what the chat will show), and `edit` a line diff of the message in Teams against
  the file. Nothing is sent. `unanswered` and `show` only read.
- **A file is posted once per target.** Every post is recorded in the state file (the file's full
  path, target, chat, message id, created time and a SHA-256 of the HTML sent). Posting the same file
  to the same target again is refused, with a pointer to `edit`; `--again` posts it anyway. Two runs
  at once wait for each other, so they can't both decide a file hasn't been posted.
- **A file is sent as a reply once per chat.** A reply is recorded the same way (target `(reply)`,
  the chat, and the quoted message as `reply_to`); a file already sent to that chat, by `reply` or
  `post`, is refused unless `--again`.
- **A change is never sent twice.** A post or an edit is sent again only after a 429 (refused before
  it was looked at) or a connection that never opened. After a timeout, a dropped connection or a
  5xx it may have been applied, so it isn't resent: the run stops, says so, and exits 1. A post or reply whose
  answer was lost is recorded as `unknown`, so the next `post` or `reply` of that file there is
  refused until you've looked (`show <target or chat id>`). An edit whose answer was lost is recorded as unconfirmed (`edit_unknown`, the
  SHA-256 of what was sent) and `show <target> <id>` tells you whether it landed; the next `edit`
  of that file reads the message first, and if it already reads as the file, records the edit as
  done and sends nothing. A 4xx is a refusal: nothing changed and nothing is recorded. Reads are
  retried after a 429, a 5xx or a dropped connection (three tries in all).
- **Every post, reply and edit is read back.** The message must have the same number of list items
  and paragraphs as was sent, the @Everyone mention when one was sent, and a reply's quote when it
  quoted a message; otherwise it says what differs
  and exits 1. A message just posted can answer "not found" for a moment, so the read waits a few
  seconds for it.
- **The log** (`log_file`) records each post, reply and edit by target or chat and message id, not
  the text.

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
