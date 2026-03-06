# Desktop Skill

> Desktop automation: screenshots, window management, keystrokes.

## Actions

| Action | Description | Params |
|--------|-------------|--------|
| `screenshot` | Full screen or window screenshot | `[--window "title"] --out-file path.png` |
| `windows` | List visible windows | |
| `focus` | Focus window by title | `--window "title"` |
| `minimize` | Minimize window | `--window "title"` |
| `maximize` | Maximize window | `--window "title"` |
| `keys` | Send keystrokes | `--keys "{ENTER}" [--window "title"]` |
| `launch` | Launch application | `--app notepad [--app-args "file.txt"] [--wait-ms 3000]` |

## Requirements

- Windows with .NET Framework (System.Windows.Forms, System.Drawing)
