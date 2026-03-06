# Outlook Skill

> Outlook COM automation: inbox, calendar, search, send, reply.

## Actions

| Action | Description | Params |
|--------|-------------|--------|
| `inbox` | List inbox messages | `--limit 15` |
| `unread` | List unread messages | `--limit 20` |
| `sent` | List sent items | `--limit 15` |
| `read` | Read email by index | `--index 0 --folder inbox\|sent\|drafts` |
| `search` | Search emails | `--query "text" --folder inbox\|sent --limit 10` |
| `calendar` | List upcoming events | `--days 7` |
| `send` | Send or draft email | `--to addr --subject text --body text [--cc addr] [--draft]` |
| `reply` | Reply to inbox email | `--index 0 --body text [--reply-all] [--draft]` |
| `folders` | List mail folders | |

## Output

All actions return structured JSON with consistent schema.

## Requirements

- Microsoft Outlook (desktop, COM-enabled)
- Non-admin PowerShell session (admin session cannot access user's Outlook profile)
