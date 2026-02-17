# mcp-winauto

Windows Desktop App Automation MCP Server.

Enables LLMs to operate Windows desktop applications (WinForms/.NET and any UIA-compatible app) via UI Automation, non-intrusively.

## Installation

```
uv sync
```

## Configuration - Claude Code

```bash
claude mcp add winauto -- uv run --directory /path/to/mcp-winauto python server.py
```

## Configuration - Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "winauto": {
      "command": "uv",
      "args": ["run", "--directory", "D:\\Workspaces\\mcp-winauto", "python", "server.py"]
    }
  }
}
```

## Tools Reference

| Tool | Description |
|---|---|
| `connect_app` | Connect to a running application by window title (regex match). Must be called first. |
| `get_ui_tree` | Get the UI element tree of the connected window. |
| `click_element` | Click a UI element via InvokePattern (no mouse movement). |
| `set_text` | Set text on a UI element via ValuePattern (no keyboard emulation). |
| `select_item` | Select an item in a combo box or list by item name. |
| `select_grid_row` | Select a row in a data grid by row index (0-based). |
| `select_menu` | Select a menu item by path, e.g. `"File->Open"`. |
| `close_window` | Close the connected application window. |
| `send_keys` | Send keyboard shortcuts to the connected window. |
| `get_text` | Get the text content of a UI element. |
| `save_screenshot` | Capture a screenshot of the connected window and save to a file. |

## Selector Format

Most tools accept a `selector` parameter as a JSON string to identify UI elements:

```json
{"title": "Name", "control_type": "Button", "auto_id": "btnOK", "parent": {"title": "Panel1"}}
```

All fields are optional, but at least one of `title`, `control_type`, or `auto_id` is required. The `parent` field accepts a nested selector to narrow the search scope.

Use `get_ui_tree` to discover element names, control types, and automation IDs.

## Key Format

The `send_keys` tool uses pywinauto key format:

| Modifier | Symbol |
|---|---|
| Ctrl | `^` |
| Alt | `%` |
| Shift | `+` |

Special keys are wrapped in braces: `{ENTER}`, `{TAB}`, `{F1}`, `{DELETE}`, etc.

Examples: `^s` (Ctrl+S), `%{F4}` (Alt+F4), `+{TAB}` (Shift+Tab), `{ENTER}` (Enter).

## Constraints

- The target application window **must not be minimized**. Restore it before performing operations.
- `send_keys` and `save_screenshot` will **briefly activate (foreground) the window** to perform their actions. All other tools operate non-intrusively via UIA patterns.
