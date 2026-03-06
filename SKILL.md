---
name: powerskills
description: Remote Windows automation via PowerSkills CLI and OneDrive relay. Outlook email/calendar, Edge browser (CDP), desktop screenshots/windows, shell commands. Use when checking Outlook inbox, reading/sending work emails, taking Windows screenshots, controlling Edge browser, or running commands on the Windows work PC. NOT for tasks that can be done locally on Mac.
---

# PowerSkills — Windows Automation via Relay

Control the Windows work PC (DESKTOP-PBITIC6) from Mac via OneDrive file sync relay.

## Architecture

```
Mac → write JSON to relay inbox → OneDrive sync (~30-60s) → Windows worker executes → outbox JSON → sync back → Mac reads result
```

Round trip: ~1-2 minutes.

## Paths

- **Relay inbox (Mac):** `/Volumes/Data/.ODContainer-OneDrive/OneDrive/4Jarvis/relay/inbox/`
- **Relay outbox (Mac):** `/Volumes/Data/.ODContainer-OneDrive/OneDrive/4Jarvis/relay/outbox/`
- **PowerSkills (Windows):** `C:\Users\alloth\OneDrive\4Jarvis\relay\PowerSkills\`
- **PowerSkills (Mac mirror):** `/Volumes/MacDev/PowerSkills/` (source of truth)

## How to Call PowerSkills

Write a JSON file to the relay inbox with `action: "exec"` and the PowerSkills command:

```python
import json, time, os

INBOX = "/Volumes/Data/.ODContainer-OneDrive/OneDrive/4Jarvis/relay/inbox"
OUTBOX = "/Volumes/Data/.ODContainer-OneDrive/OneDrive/4Jarvis/relay/outbox"
PS_DIR = "C:\\Users\\alloth\\OneDrive\\4Jarvis\\relay\\PowerSkills"

def powerskills(skill, action, params=""):
    cmd_id = f"ps-{skill}-{int(time.time())}"
    cmd = f"Set-Location '{PS_DIR}'; .\\powerskills.ps1 {skill} {action} {params}"
    request = {"id": cmd_id, "action": "exec", "command": cmd}
    with open(f"{INBOX}/{cmd_id}.json", "w") as f:
        json.dump(request, f, indent=2)
    return cmd_id

def read_result(cmd_id, timeout=180):
    """Poll outbox for result. Returns parsed JSON."""
    path = f"{OUTBOX}/{cmd_id}.json"
    for _ in range(timeout // 5):
        if os.path.exists(path):
            with open(path, encoding="utf-8-sig") as f:  # BOM!
                data = json.loads(f.read())
            stdout = data.get("stdout", "")
            if stdout:
                return json.loads(stdout)
            return data
        time.sleep(5)
    return None
```

## Available Skills

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

### browser
| Action | Params | Description |
|--------|--------|-------------|
| `tabs` | | List open Edge tabs |
| `navigate` | `--url URL` | Navigate to URL |
| `screenshot` | `--out-file path.png` | Capture page screenshot |
| `content` | | Get page text |
| `evaluate` | `--expression "js"` | Execute JavaScript |
| `click` | `--selector "#btn"` | Click element |
| `type` | `--selector "#input" --text "hello"` | Type into element |

### desktop
| Action | Params | Description |
|--------|--------|-------------|
| `screenshot` | `--out-file path.png [--window "title"]` | Screenshot |
| `windows` | | List visible windows |
| `focus` | `--window "title"` | Focus window |
| `keys` | `--keys "{ENTER}" [--window "title"]` | Send keystrokes |
| `launch` | `--app notepad` | Launch application |

### system
| Action | Params | Description |
|--------|--------|-------------|
| `exec` | `--command "whoami"` | Run PowerShell command |
| `info` | | System info |
| `processes` | `--limit N` | Top processes |

## Checking Relay Availability

Before running commands, verify the relay is active:

```bash
# Write a ping, wait up to 90s for response
ID="ping-$(date +%s)"
echo '{"id":"'$ID'","action":"exec","command":"echo pong"}' > "$INBOX/$ID.json"
# Poll outbox for $ID.json
```

If no response within 90 seconds, the relay worker is not running on Windows. Notify the user.

## Windows Setup

If running interactively on Windows, scripts may be blocked by execution policy:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

The relay worker already uses `-ExecutionPolicy Bypass` so this only affects direct interactive use.

## Known Quirks

- **UTF-8 BOM:** Windows writes JSON with BOM. Always read with `encoding='utf-8-sig'`.
- **OneDrive sync delay:** 30-60 seconds each direction. Total round trip ~1-2 min.
- **Outlook COM:** Requires non-admin PowerShell. Admin session cannot access user's Outlook profile.
- **Edge CDP:** Edge must be running with `--remote-debugging-port=9222`.
- **iCloud file locks:** If writing to iCloud-synced Obsidian folders fails with `errno 11` (Resource deadlock), write to a different filename.
- **Search is slow:** Outlook search via relay can be unreliable. Prefer listing + filtering locally.
- **Sent items folder:** Use `powerskills outlook sent` (dedicated action). The old relay's `read` action ignored the `folder` parameter — PowerSkills fixes this.

## Deployment Sync

When updating PowerSkills source on Mac (`/Volumes/MacDev/PowerSkills/`), copy to relay mirror:

```bash
rsync -av --exclude='.git' /Volumes/MacDev/PowerSkills/ "/Volumes/Data/.ODContainer-OneDrive/OneDrive/4Jarvis/relay/PowerSkills/"
```

OneDrive will sync to Windows automatically.
