# RemindersToExcel

AppleScript that exports reminders from the macOS Reminders app to a CSV file on your Desktop.

## What It Does

Exports every reminder across all lists (or a single list) into a CSV with these columns:

| Column | Description |
|--------|-------------|
| List Name | Which reminder list it belongs to |
| Title | Reminder name |
| Notes | Body text |
| Due Date | YYYY-MM-DD HH:MM format |
| Creation Date | YYYY-MM-DD HH:MM format |
| Completion Date | YYYY-MM-DD HH:MM format |
| Completed | Yes / No |
| Priority | None, Low, Medium, High |
| Flagged | Yes / No |

Output file is saved to `~/Desktop/Reminders_Export_YYYY-MM-DD_HHMM.csv` (or `Reminders_<ListName>_YYYY-MM-DD_HHMM.csv` for single-list exports).

## Usage

### From Script Editor or Terminal

Run directly and it exports all lists by default:

```bash
osascript RemindersToCSV.applescript
```

Export a specific list:

```bash
osascript RemindersToCSV.applescript "My List Name"
```

### From Apple Shortcuts

Set up a shortcut with these actions:

1. **Choose from List** -- present your reminder list names plus "All Lists"
2. **Text** -- pass the chosen value as text
3. **Run AppleScript** -- paste the script contents; the chosen list name flows in as input

The script detects it was called from Shortcuts and runs silently (no dialogs). It returns the output file path, which you can pass to subsequent actions like **Quick Look** or **Show Notification**.

When no input is provided (Script Editor, direct run), it defaults to exporting all lists and shows a completion dialog.

## Performance

Uses bulk property fetching (`name of every reminder of aList`) to pull all values per property in a single Apple Event call, rather than querying each reminder individually. This reduces IPC overhead from ~8N calls to ~8 calls per list.
