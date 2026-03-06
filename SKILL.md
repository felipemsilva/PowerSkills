---
name: powerskills
description: Windows automation for AI agents via PowerShell. Outlook email/calendar, Edge browser (CDP), desktop screenshots/window management, shell commands. Use when needing to interact with Outlook, control Edge browser, take screenshots, manage windows, or run system commands on Windows.
---

# PowerSkills

Windows capabilities for AI agents. Call `powerskills.ps1` with a skill and action to get structured JSON output.

## Usage

```powershell
.\powerskills.ps1 <skill> <action> [--param value ...]
.\powerskills.ps1 list                          # List available skills
.\powerskills.ps1 outlook help                   # Show skill help
.\powerskills.ps1 outlook inbox --limit 10       # Run action
```

## Setup

If scripts are blocked, run once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

## Skills

### outlook
| Action | Params | Description |
|--------|--------|-------------|
| `inbox` | `--limit N` | List inbox messages |
| `unread` | `--limit N` | List unread messages |
| `sent` | `--limit N` | List sent items |
| `read` | `--index N --folder inbox\|sent\|drafts` | Read full email |
| `search` | `--query "text" --folder inbox\|sent --limit N` | Search emails |
| `calendar` | `--days N` | List upcoming events |
| `send` | `--to addr --subject text --body text [--cc addr] [--draft]` | Send/draft email |
| `reply` | `--index N --body text [--reply-all] [--draft]` | Reply to inbox email |
| `folders` | | List mail folders |

Requires: Outlook desktop (COM-enabled), non-admin PowerShell.

### browser
| Action | Params | Description |
|--------|--------|-------------|
| `tabs` | | List open Edge tabs |
| `navigate` | `--url URL` | Navigate to URL |
| `screenshot` | `--out-file path.png` | Capture page screenshot |
| `content` | | Get page text |
| `html` | | Get page HTML |
| `evaluate` | `--expression "js"` | Execute JavaScript |
| `click` | `--selector "#btn"` | Click element |
| `type` | `--selector "#input" --text "hello"` | Type into element |
| `new-tab` | `--url URL` | Open new tab |
| `close-tab` | `--target-id id` | Close tab |
| `scroll` | `--scroll-target top\|bottom\|selector` | Scroll page |
| `fill` | `--fields-json '[...]'` | Fill multiple form fields |

Requires: Edge running with `--remote-debugging-port=9222`.

### desktop
| Action | Params | Description |
|--------|--------|-------------|
| `screenshot` | `--out-file path.png [--window "title"]` | Full screen or window screenshot |
| `windows` | | List visible windows |
| `focus` | `--window "title"` | Focus window |
| `minimize` | `--window "title"` | Minimize window |
| `maximize` | `--window "title"` | Maximize window |
| `keys` | `--keys "{ENTER}" [--window "title"]` | Send keystrokes |
| `launch` | `--app notepad [--app-args "file.txt"]` | Launch application |

### system
| Action | Params | Description |
|--------|--------|-------------|
| `exec` | `--command "whoami" [--timeout 30]` | Run PowerShell command |
| `info` | | Hostname, OS, user, uptime |
| `processes` | `--limit N` | Top processes by CPU |
| `env` | `--name PATH` | Get environment variable |

## Output Format

All actions return JSON:

```json
{"status": "success", "exit_code": 0, "data": {...}, "timestamp": "..."}
```

## Configuration

Edit `config.json`:

```json
{
  "edge_debug_port": 9222,
  "default_timeout": 30,
  "outlook_body_max_chars": 5000
}
```
