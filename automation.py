"""WindowManager: wraps all pywinauto (UIA backend) operations."""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import os
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError

from pywinauto import Application


def _get_window_rects(hwnd: int):
    """Return (full_rect, visible_rect) for the given window handle.

    full_rect: from GetWindowRect (includes DWM shadow/extended frame).
    visible_rect: from DwmGetWindowAttribute DWMWA_EXTENDED_FRAME_BOUNDS
                   (the actual visible area without shadow).

    Both are (left, top, right, bottom) in screen coordinates.
    Returns (full_rect, None) if the DWM call fails.
    """
    hwnd_c = ctypes.wintypes.HWND(hwnd)

    full = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd_c, ctypes.byref(full))
    full_rect = (full.left, full.top, full.right, full.bottom)

    vis = ctypes.wintypes.RECT()
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
        hwnd_c,
        ctypes.wintypes.DWORD(DWMWA_EXTENDED_FRAME_BOUNDS),
        ctypes.byref(vis),
        ctypes.sizeof(vis),
    )
    if hr == 0:  # S_OK
        vis_rect = (vis.left, vis.top, vis.right, vis.bottom)
        return full_rect, vis_rect
    return full_rect, None

# Timeout (seconds) for UIA operations that may block (e.g. invoke on
# a button that opens a modal dialog).
_OP_TIMEOUT = 5

_executor = ThreadPoolExecutor(max_workers=1)


def _run_with_timeout(func, *args, timeout=_OP_TIMEOUT):
    """Run *func* in a worker thread with a timeout.

    Returns the result of *func* on success.
    Raises RuntimeError if the call does not complete within *timeout* seconds.
    """
    future = _executor.submit(func, *args)
    try:
        return future.result(timeout=timeout)
    except FuturesTimeoutError:
        raise RuntimeError(
            f"Operation timed out after {timeout}s. "
            "A modal dialog may be blocking the target application. "
            "Close any open dialogs and retry."
        )


class WindowManager:
    """High-level wrapper around pywinauto UIA-backend operations."""

    def __init__(self) -> None:
        self._app: Application | None = None
        self._window = None

    # ------------------------------------------------------------------
    # Property
    # ------------------------------------------------------------------

    @property
    def window(self):
        """Return the connected window wrapper.

        Raises RuntimeError if no app is connected or the window is gone.
        """
        if self._app is None or self._window is None:
            raise RuntimeError("No app connected. Call connect_app first.")
        try:
            if not self._window.exists():
                raise RuntimeError("Connected window no longer exists.")
        except Exception as exc:
            if "no longer exists" in str(exc):
                raise
            raise RuntimeError("Connected window no longer exists.") from exc
        # Check for minimized state
        try:
            wrapper = self._window.wrapper_object()
            if hasattr(wrapper, 'is_minimized') and wrapper.is_minimized():
                raise RuntimeError(
                    "Window is minimized. Restore it before performing operations."
                )
        except RuntimeError:
            raise
        except Exception:
            pass
        return self._window

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self, app_name_regex: str) -> str:
        """Connect to a running application by window-title regex.

        Returns the window title on success.
        Raises LookupError when no matching window is found.
        """
        try:
            app = Application(backend="uia").connect(
                title_re=app_name_regex, timeout=5
            )
            self._app = app
            self._window = app.top_window()
            return self._window.window_text()
        except Exception as exc:
            self._app = None
            self._window = None
            raise LookupError(
                f"Could not connect to app matching '{app_name_regex}': {exc}"
            ) from exc

    # ------------------------------------------------------------------
    # UI tree
    # ------------------------------------------------------------------

    def get_ui_tree(self, max_depth: int = 3) -> str:
        """Return a hierarchical text representation of the UI tree.

        Each line has the format:
            {indent}{ControlType}  Name="{Name}"  AutoId="{AutomationId}"

        Uses 2 spaces per depth level.  Name is truncated to 50 chars.
        """
        lines: list[str] = []

        def _walk(element, depth: int) -> None:
            info = element.element_info
            control_type = info.control_type or ""
            name = info.name or ""
            if len(name) > 50:
                name = name[:50]
            auto_id = info.automation_id or ""
            indent = "  " * depth
            lines.append(
                f'{indent}{control_type}  Name="{name}"  AutoId="{auto_id}"'
            )
            if depth < max_depth:
                try:
                    for child in element.children():
                        _walk(child, depth + 1)
                except Exception:
                    pass

        win = self.window
        _walk(win, 0)
        return "\n".join(lines)

    # ------------------------------------------------------------------
    # Element finding
    # ------------------------------------------------------------------

    def find_element(self, selector: dict):
        """Find a single UI element matching *selector*.

        Supported selector keys:
            title        - matches element_info.name
            control_type - matches element_info.control_type
            auto_id      - matches element_info.automation_id
            parent       - a nested selector dict; search starts from that parent

        Returns the first matching pywinauto wrapper.
        Raises LookupError if nothing matches.
        """
        has_criteria = any(
            k in selector for k in ("title", "control_type", "auto_id")
        )
        if not has_criteria:
            raise LookupError(
                "Selector must contain at least one of: title, control_type, auto_id"
            )

        if "parent" in selector:
            search_from = self.find_element(selector["parent"])
        else:
            search_from = self.window

        title = selector.get("title")
        control_type = selector.get("control_type")
        auto_id = selector.get("auto_id")

        def _search(element):
            for child in element.children():
                info = child.element_info
                if title is not None and info.name != title:
                    pass
                elif control_type is not None and info.control_type != control_type:
                    pass
                elif auto_id is not None and info.automation_id != auto_id:
                    pass
                else:
                    return child
                # Recurse into children
                result = _search(child)
                if result is not None:
                    return result
            return None

        match = _search(search_from)
        if match is None:
            raise LookupError(f"Element not found: {selector}")
        return match

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def click(self, selector: dict) -> str:
        """Click (invoke) the element described by *selector*.

        Tries InvokePattern first, then TogglePattern, then
        ExpandCollapsePattern.  Raises RuntimeError if none work.
        """
        element = self.find_element(selector)
        info = element.element_info
        control_type = info.control_type or ""
        name = info.name or ""

        # Try InvokePattern (with timeout to avoid modal dialog hangs)
        try:
            _run_with_timeout(element.invoke)
            return f"Clicked: {control_type} '{name}'"
        except RuntimeError as exc:
            if "timed out" in str(exc):
                return f"Clicked: {control_type} '{name}' (action may have opened a dialog)"
            pass
        except Exception:
            pass

        # Try TogglePattern (checkboxes)
        try:
            _run_with_timeout(element.toggle)
            return f"Clicked: {control_type} '{name}'"
        except RuntimeError as exc:
            if "timed out" in str(exc):
                return f"Clicked: {control_type} '{name}' (action may have opened a dialog)"
        except Exception:
            pass

        # Try ExpandCollapsePattern
        try:
            try:
                state = element.get_expand_state()
            except Exception:
                state = None
            if state is not None:
                if state == 0:  # Collapsed
                    _run_with_timeout(element.expand)
                else:
                    _run_with_timeout(element.collapse)
                return f"Clicked: {control_type} '{name}'"
        except RuntimeError as exc:
            if "timed out" in str(exc):
                return f"Clicked: {control_type} '{name}' (action may have opened a dialog)"
        except Exception:
            pass

        raise RuntimeError(
            f"Cannot click {control_type} '{name}': "
            "element does not support Invoke, Toggle, or ExpandCollapse patterns."
        )

    def set_text(self, selector: dict, text: str) -> str:
        """Set text on the element described by *selector*.

        Tries ValuePattern.set_value first, falls back to set_edit_text.
        """
        element = self.find_element(selector)
        info = element.element_info
        control_type = info.control_type or ""
        name = info.name or ""

        # Try ValuePattern
        try:
            element.iface_value.SetValue(text)
            return f"Set text on {control_type} '{name}'"
        except Exception:
            pass

        # Fallback: set_edit_text
        try:
            element.set_edit_text(text)
            return f"Set text on {control_type} '{name}'"
        except Exception as exc:
            raise RuntimeError(
                f"Cannot set text on {control_type} '{name}': {exc}"
            ) from exc

    def get_text(self, selector: dict) -> str:
        """Return text from the element described by *selector*.

        Tries ValuePattern.value first, then window_text().
        """
        element = self.find_element(selector)

        # Try ValuePattern
        try:
            return element.iface_value.CurrentValue
        except Exception:
            pass

        # Fallback: window_text
        try:
            return element.window_text()
        except Exception as exc:
            raise RuntimeError(
                f"Cannot get text from element: {exc}"
            ) from exc

    def select_item(self, selector: dict, item_name: str) -> str:
        """Select an item by name inside a combo-box / list element.

        Expands the element, finds the child whose Name matches *item_name*,
        selects it via SelectionItemPattern, and collapses the element.
        """
        element = self.find_element(selector)
        info = element.element_info
        control_type = info.control_type or ""
        name = info.name or ""

        # Expand (sleep briefly to let dropdown populate)
        try:
            element.expand()
            time.sleep(0.3)
        except Exception:
            pass

        # Find child item matching item_name
        item = None
        for child in element.descendants():
            if child.element_info.name == item_name:
                item = child
                break

        if item is None:
            try:
                element.collapse()
            except Exception:
                pass
            raise LookupError(
                f"Item '{item_name}' not found in {control_type} '{name}'"
            )

        # Select
        try:
            item.iface_selection_item.Select()
        except Exception as exc:
            try:
                element.collapse()
            except Exception:
                pass
            raise RuntimeError(
                f"Cannot select '{item_name}': {exc}"
            ) from exc

        # Collapse
        try:
            element.collapse()
        except Exception:
            pass

        return f"Selected '{item_name}' in {control_type} '{name}'"

    def select_grid_row(self, selector: dict, row_index: int) -> str:
        """Select a row in a data-grid by index.

        Finds row elements among the grid's children and selects
        the target row via SelectionItemPattern.
        """
        element = self.find_element(selector)
        info = element.element_info
        control_type = info.control_type or ""
        name = info.name or ""

        # Collect row elements (filter out headers, scrollbars, etc.)
        rows = []
        for child in element.children():
            child_name = child.element_info.name or ""
            child_type = child.element_info.control_type or ""
            # WinForms DataGridView rows are typically "Custom" or "DataItem" elements
            # Header rows often contain "トップ" or "Header" in their name
            if child_type in ("Custom", "DataItem"):
                if "トップ" not in child_name and "Header" not in child_name.lower():
                    rows.append(child)

        if row_index < 0 or row_index >= len(rows):
            raise IndexError(
                f"Row index {row_index} out of range (0-{len(rows) - 1}) "
                f"in {control_type} '{name}'"
            )

        target_row = rows[row_index]

        # Try SelectionItemPattern on the row itself
        try:
            target_row.iface_selection_item.Select()
            return f"Selected row {row_index} in {control_type} '{name}'"
        except Exception:
            pass

        # Fallback: try SelectionItemPattern on the first cell
        try:
            cells = target_row.children()
            if cells:
                cells[0].iface_selection_item.Select()
                return f"Selected row {row_index} in {control_type} '{name}'"
        except Exception:
            pass

        # Fallback: click the row (with timeout)
        try:
            _run_with_timeout(target_row.click_input)
            return f"Selected row {row_index} in {control_type} '{name}'"
        except Exception as exc:
            raise RuntimeError(
                f"Cannot select row {row_index} in {control_type} '{name}': {exc}"
            ) from exc

    def select_menu(self, menu_path: str) -> str:
        """Select a menu item by navigating the menu hierarchy.

        *menu_path* uses arrow separator, e.g. ``"File->Open"``.
        Each segment is clicked via InvokePattern in sequence.
        """
        segments = [s.strip() for s in menu_path.split("->")]
        if not segments:
            raise ValueError("Empty menu path")

        win = self.window

        # Find the menu bar
        menu_bar = None
        for child in win.children():
            if child.element_info.control_type == "MenuBar":
                menu_bar = child
                break

        if menu_bar is None:
            raise RuntimeError("No menu bar found in the window")

        # Click the first menu item from the menu bar
        current = None
        for item in menu_bar.children():
            if item.element_info.name == segments[0]:
                current = item
                break

        if current is None:
            raise LookupError(f"Menu item '{segments[0]}' not found in menu bar")

        try:
            _run_with_timeout(current.invoke)
        except RuntimeError as exc:
            if "timed out" not in str(exc):
                try:
                    _run_with_timeout(current.expand)
                except Exception as exc2:
                    raise RuntimeError(f"Cannot open menu '{segments[0]}': {exc2}") from exc2
        except Exception:
            try:
                _run_with_timeout(current.expand)
            except Exception as exc:
                raise RuntimeError(f"Cannot open menu '{segments[0]}': {exc}") from exc

        # Navigate through remaining segments
        for i, segment in enumerate(segments[1:], 1):
            time.sleep(0.3)  # wait for submenu to appear
            found = False
            for d in win.descendants():
                info = d.element_info
                if info.control_type == "MenuItem" and info.name == segment:
                    try:
                        _run_with_timeout(d.invoke)
                    except RuntimeError as exc:
                        if "timed out" not in str(exc):
                            try:
                                _run_with_timeout(d.expand)
                            except Exception as exc2:
                                raise RuntimeError(f"Cannot click menu '{segment}': {exc2}") from exc2
                    except Exception:
                        try:
                            _run_with_timeout(d.expand)
                        except Exception as exc:
                            raise RuntimeError(f"Cannot click menu '{segment}': {exc}") from exc
                    found = True
                    break
            if not found:
                # Close any open menus
                try:
                    win.type_keys("{ESC}")
                except Exception:
                    pass
                raise LookupError(f"Menu item '{segment}' not found")

        return f"Selected menu: {menu_path}"

    def close_window(self) -> str:
        """Close the connected window and reset internal state."""
        try:
            self.window.close()
        except Exception:
            pass
        self._app = None
        self._window = None
        return "Window closed"

    def send_keys(self, keys: str) -> str:
        """Send keystrokes to the connected window.

        Brings the window to the foreground first.
        """
        win = self.window
        try:
            win.set_focus()
            win.type_keys(keys, pause=0.05)
        except Exception as exc:
            raise RuntimeError(f"Cannot send keys: {exc}") from exc
        return f"Sent keys: {keys}"

    def save_screenshot(self, filename: str) -> str:
        """Capture a screenshot of the connected window and save it to *filename*.

        If *filename* is not absolute it is resolved relative to the cwd.
        A ``.png`` extension is enforced.  Parent directories are created
        as needed.
        """
        # Resolve path
        if not os.path.isabs(filename):
            save_path = os.path.join(os.getcwd(), filename)
        else:
            save_path = filename

        # Ensure .png extension
        root, ext = os.path.splitext(save_path)
        if ext.lower() != ".png":
            save_path = root + ".png"

        # Ensure parent directory exists
        parent = os.path.dirname(save_path)
        if parent:
            os.makedirs(parent, exist_ok=True)

        win = self.window
        try:
            win.set_focus()
            time.sleep(0.3)
            image = win.capture_as_image()

            # Crop out DWM shadow / extended frame
            hwnd = win.element_info.handle
            if hwnd:
                full_rect, vis_rect = _get_window_rects(hwnd)
                if vis_rect is not None:
                    # Offsets in screen coordinates
                    left_off = vis_rect[0] - full_rect[0]
                    top_off = vis_rect[1] - full_rect[1]
                    right_off = full_rect[2] - vis_rect[2]
                    bottom_off = full_rect[3] - vis_rect[3]
                    if left_off > 0 or top_off > 0 or right_off > 0 or bottom_off > 0:
                        img_w, img_h = image.size
                        screen_w = full_rect[2] - full_rect[0]
                        screen_h = full_rect[3] - full_rect[1]
                        # Account for DPI scaling between screen coords and image pixels
                        scale_x = img_w / screen_w if screen_w else 1
                        scale_y = img_h / screen_h if screen_h else 1
                        image = image.crop((
                            round(left_off * scale_x),
                            round(top_off * scale_y),
                            img_w - round(right_off * scale_x),
                            img_h - round(bottom_off * scale_y),
                        ))

            image.save(save_path)
        except Exception as exc:
            raise RuntimeError(f"Cannot save screenshot: {exc}") from exc

        return f"Screenshot saved to: {os.path.abspath(save_path)}"
