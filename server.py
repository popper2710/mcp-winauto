"""MCP server entry point for mcp-winauto.

Exposes 13 tools for Windows desktop application automation via UI Automation.
"""

import json

from mcp.server.fastmcp import FastMCP
from automation import WindowManager

mcp = FastMCP("winauto")
wm = WindowManager()


# ------------------------------------------------------------------
# Tool 1: connect_app
# ------------------------------------------------------------------

@mcp.tool()
def connect_app(app_name_regex: str) -> str:
    """Connect to a running Windows application by window title (regex match).
    Example: connect_app(".*Notepad.*") or connect_app(".*メモ帳.*")
    Must be called before using any other tools."""
    try:
        title = wm.connect(app_name_regex)
        return f"Connected to: {title}"
    except LookupError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 2: get_ui_tree
# ------------------------------------------------------------------

@mcp.tool()
def get_ui_tree() -> str:
    """Get the UI element tree of the connected application window.
    Returns hierarchical text showing ControlType, Name, and AutomationId.
    Use this to discover elements before operating on them."""
    try:
        tree = wm.get_ui_tree()
        return tree
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 3: click_element
# ------------------------------------------------------------------

@mcp.tool()
def click_element(selector: str) -> str:
    """Click a UI element using non-intrusive InvokePattern (no mouse movement).
    selector: JSON string, e.g. {"title": "OK", "control_type": "Button"}
    Selector fields (use any combination): title, control_type, auto_id, parent"""
    try:
        sel = json.loads(selector)
        result = wm.click(sel)
        return result
    except json.JSONDecodeError as e:
        return f"Error: {e}"
    except LookupError as e:
        return f"Error: {e}"
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 4: set_text
# ------------------------------------------------------------------

@mcp.tool()
def set_text(selector: str, text: str) -> str:
    """Set text on a UI element using non-intrusive ValuePattern (no keyboard emulation).
    selector: JSON string, e.g. {"auto_id": "txtName"}"""
    try:
        sel = json.loads(selector)
        result = wm.set_text(sel, text)
        return result
    except json.JSONDecodeError as e:
        return f"Error: {e}"
    except LookupError as e:
        return f"Error: {e}"
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 5: select_item
# ------------------------------------------------------------------

@mcp.tool()
def select_item(selector: str, item_name: str) -> str:
    """Select an item in a combo box or list by item name.
    selector: JSON string identifying the combo/list element.
    item_name: The visible text of the item to select."""
    try:
        sel = json.loads(selector)
        result = wm.select_item(sel, item_name)
        return result
    except json.JSONDecodeError as e:
        return f"Error: {e}"
    except LookupError as e:
        return f"Error: {e}"
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 6: select_grid_row
# ------------------------------------------------------------------

@mcp.tool()
def select_grid_row(selector: str, row_index: int) -> str:
    """Select a row in a data grid by row index (0-based).
    selector: JSON string identifying the grid element."""
    try:
        sel = json.loads(selector)
        result = wm.select_grid_row(sel, row_index)
        return result
    except json.JSONDecodeError as e:
        return f"Error: {e}"
    except LookupError as e:
        return f"Error: {e}"
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 7: select_menu
# ------------------------------------------------------------------

@mcp.tool()
def select_menu(menu_path: str) -> str:
    """Select a menu item from the menu bar.
    menu_path: Arrow-separated path, e.g. "File->Open" or "Edit->Find->Replace"."""
    try:
        result = wm.select_menu(menu_path)
        return result
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 8: close_window
# ------------------------------------------------------------------

@mcp.tool()
def close_window() -> str:
    """Close the connected application window."""
    try:
        result = wm.close_window()
        return result
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 9: send_keys
# ------------------------------------------------------------------

@mcp.tool()
def send_keys(keys: str) -> str:
    """Send keyboard shortcuts to the connected window.
    Uses pywinauto key format: ^ = Ctrl, % = Alt, + = Shift.
    Examples: "^s" (Ctrl+S), "%{F4}" (Alt+F4), "{ENTER}" (Enter key).
    NOTE: This will briefly bring the window to the foreground."""
    try:
        result = wm.send_keys(keys)
        return result
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 10: get_text
# ------------------------------------------------------------------

@mcp.tool()
def get_text(selector: str) -> str:
    """Get the text content of a UI element.
    selector: JSON string, e.g. {"auto_id": "lblStatus"}
    Returns the element's current text value."""
    try:
        sel = json.loads(selector)
        result = wm.get_text(sel)
        return result
    except json.JSONDecodeError as e:
        return f"Error: {e}"
    except LookupError as e:
        return f"Error: {e}"
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 11: save_screenshot
# ------------------------------------------------------------------

@mcp.tool()
def save_screenshot(filename: str) -> str:
    """Capture a screenshot of the connected window and save to a local file.
    filename: File name or path (e.g. "screenshot.png"). Saved to current directory if relative.
    NOTE: This will briefly bring the window to the foreground."""
    try:
        result = wm.save_screenshot(filename)
        return result
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 12: list_windows
# ------------------------------------------------------------------

@mcp.tool()
def list_windows() -> str:
    """List all visible windows of the connected application.
    Returns window index, title, and which window is the current target.
    Use switch_window to change the target window."""
    try:
        windows = wm.list_windows()
        if not windows:
            return "No visible windows found."
        lines = []
        for w in windows:
            markers = []
            if w["is_main"]:
                markers.append("main")
            if w["is_current"]:
                markers.append("current")
            suffix = f"  [{', '.join(markers)}]" if markers else ""
            lines.append(f"  {w['index']}: {w['title']}{suffix}")
        return "Windows:\n" + "\n".join(lines)
    except RuntimeError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Tool 13: switch_window
# ------------------------------------------------------------------

@mcp.tool()
def switch_window(title: str = None, index: int = None) -> str:
    """Switch which window is the target for all subsequent operations.
    Provide either title (substring match) or index (from list_windows).
    Switching to the main window re-enables automatic dialog detection."""
    try:
        new_title = wm.switch_window(title=title, index=index)
        return f"Switched to: {new_title}"
    except (ValueError, IndexError, LookupError, RuntimeError) as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"


# ------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------

def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
