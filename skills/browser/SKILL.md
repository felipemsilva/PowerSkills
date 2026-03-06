# Browser Skill

> Edge CDP browser automation: tabs, navigate, screenshot, content, click, type.

## Actions

| Action | Description | Params |
|--------|-------------|--------|
| `tabs` | List open browser tabs | |
| `navigate` | Navigate to URL | `--url https://...` |
| `screenshot` | Capture page screenshot | `--out-file path.png [--target-id id]` |
| `content` | Get page text content | `[--target-id id]` |
| `html` | Get full page HTML | `[--target-id id]` |
| `evaluate` | Execute JavaScript | `--expression "document.title"` |
| `click` | Click element by selector | `--selector "#btn"` |
| `type` | Type into element | `--selector "#input" --text "hello"` |
| `new-tab` | Open new tab | `--url https://...` |
| `close-tab` | Close tab | `--target-id id` |
| `scroll` | Scroll page | `--scroll-target top\|bottom\|selector` |
| `fill` | Fill multiple form fields | `--fields-json '[{"selector":"#a","value":"b"}]'` |
| `wait` | Wait N seconds | `--seconds 3` |

## Requirements

- Microsoft Edge running with `--remote-debugging-port=9222`
- Launch: `msedge --remote-debugging-port=9222`
