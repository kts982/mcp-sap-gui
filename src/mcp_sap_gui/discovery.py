"""
Discovery Mixin - popups, toolbars, shell content, screen elements, screenshots.

Provides screen discovery, popup handling, toolbar enumeration, and
screenshot capabilities for the SAP GUI controller.
"""

import logging
import re
from typing import Any, Dict, List

from .models import SAPGUIError, ScreenElement

logger = logging.getLogger(__name__)

# GuiFrameWindow.HardCopy image type. It must be the numeric constant: a
# string like "PNG" is ignored and SAP GUI writes an uncompressed BMP
# (18 MB for a maximised window, versus 1.3 MB as PNG).
_HARDCOPY_PNG = 2


class DiscoveryMixin:
    """Mixin for discovery and inspection operations on SAP GUI screens."""

    _POPUP_INTERACTIVE_TYPES = (
        "GuiTextField",
        "GuiCTextField",
        "GuiPasswordField",
        "GuiComboBox",
        "GuiCheckBox",
        "GuiRadioButton",
        "GuiTableControl",
        "GuiGridView",
        "GuiTree",
    )
    _POPUP_ERROR_PATTERNS = (
        "error",
        "failed",
        "cannot",
        "can not",
        "not possible",
        "not found",
        "does not exist",
        "required",
        "invalid",
    )
    _POPUP_INFO_PATTERNS = (
        "information",
        "info",
        "success",
        "completed",
        "saved",
        "created",
        "displayed",
    )
    _POPUP_INPUT_PATTERNS = (
        "enter",
        "select",
        "choose",
        "specify",
        "values",
        "parameters",
        "selection",
        "organization",
        "warehouse",
        "plant",
    )
    _POPUP_CONFIRMATION_PATTERNS = (
        "do you want",
        "are you sure",
        "save changes",
        "discard",
        "overwrite",
        "delete",
        "confirm",
        "continue?",
        "proceed",
    )

    # =========================================================================
    # Popup Window Handling
    # =========================================================================

    def get_popup_window(self) -> Dict[str, Any]:
        """
        Check if a popup window (modal dialog) is currently open.

        SAP popups appear as wnd[1], wnd[2], etc.  Returns the topmost
        popup's title, text content, and available buttons so the AI can
        decide how to respond, plus a coarse classification and safe
        next-action hint.

        Returns:
            Dict with popup info or {"popup_exists": False}
        """
        self._require_session()

        # Find the topmost window above wnd[0]
        popup_wnd = None
        popup_id = None
        for i in range(1, 10):
            wnd_id = f"wnd[{i}]"
            try:
                wnd = self._session.findById(wnd_id)
                popup_wnd = wnd
                popup_id = wnd_id
            except Exception:
                break

        if popup_wnd is None:
            return {"popup_exists": False}

        result: Dict[str, Any] = {
            "popup_exists": True,
            "window_id": popup_id,
            "title": getattr(popup_wnd, 'Text', ''),
        }

        # Read the status bar message in the popup if any
        try:
            sbar = self._session.findById(f"{popup_id}/sbar")
            result["message"] = sbar.Text
            result["message_type"] = getattr(sbar, 'MessageType', '')
        except Exception:
            pass

        # Collect text elements and buttons from the user area
        texts = []
        buttons = []
        interactive_elements = []
        try:
            usr = self._session.findById(f"{popup_id}/usr")
            self._collect_popup_contents(usr, texts, buttons, interactive_elements)
        except Exception:
            pass

        # Also check toolbar buttons (tbar[0] for OK/Cancel)
        for tbar_idx in range(2):
            try:
                tbar = self._session.findById(f"{popup_id}/tbar[{tbar_idx}]")
                for i in range(tbar.Children.Count):
                    btn = tbar.Children(i)
                    if getattr(btn, 'Type', '') in ('GuiButton',):
                        buttons.append({
                            "id": self._normalize_element_id(btn.Id),
                            "text": getattr(btn, 'Text', '').strip(),
                            "tooltip": getattr(btn, 'Tooltip', '').strip(),
                        })
            except Exception:
                pass

        if texts:
            result["texts"] = texts
        if buttons:
            result["buttons"] = buttons
        if interactive_elements:
            result["interactive_elements"] = interactive_elements
            result["has_inputs"] = True
        else:
            result["has_inputs"] = False

        prefilled = self._prefilled_inputs(interactive_elements)
        if prefilled:
            result["prefilled_inputs"] = prefilled
            result["notice"] = self._PREFILL_NOTICE

        return self._classify_popup(result)

    # Deliberately generic: no popup is recognised by name. Any dialog whose
    # changeable inputs already hold values gets the same notice, because a
    # confirm accepts defaults the agent never chose (a transport request
    # prompt pre-filled with another project's request is one such case).
    _PREFILL_TYPES = ("GuiTextField", "GuiCTextField", "GuiComboBox")
    _PREFILL_NOTICE = (
        "This popup has pre-filled input values. Confirming accepts them as "
        "they are: verify each one is what the user intends, and change it "
        "first if it is not."
    )

    def _prefilled_inputs(self, interactive_elements: list) -> list:
        """Changeable inputs of a popup that already hold a value."""
        return [
            {
                "id": el["id"],
                "name": el.get("name", ""),
                "value": self._mask_field_value(el["id"], el["text"]),
            }
            for el in interactive_elements
            if el.get("type") in self._PREFILL_TYPES
            and el.get("changeable")
            and el.get("text")
        ]

    def _popup_digest(self) -> Dict[str, Any]:
        """Compact summary of the open popup, attached to action responses.

        The title is left out: screen info already reads it from the active
        window. Full details (element IDs) stay with get_popup_window.
        """
        try:
            popup = self.get_popup_window()
        except Exception:
            return {}
        if not popup.get("popup_exists"):
            return {}

        digest: Dict[str, Any] = {
            "classification": popup.get("classification", "unknown"),
        }
        texts = popup.get("texts") or []
        if texts:
            digest["texts"] = [str(t)[:200] for t in texts[:6]]
        buttons = [
            b.get("text") or b.get("tooltip") for b in popup.get("buttons", [])
        ]
        buttons = [b for b in buttons if b]
        if buttons:
            digest["buttons"] = buttons
        for key in ("prefilled_inputs", "notice"):
            if key in popup:
                digest[key] = popup[key]
        return digest

    def _collect_popup_contents(
        self,
        container,
        texts: list,
        buttons: list,
        interactive_elements: list,
        depth: int = 0,
    ) -> None:
        """Recursively collect text and buttons from a popup's user area."""
        if depth > 3:
            return
        try:
            for i in range(container.Children.Count):
                child = container.Children(i)
                ctype = getattr(child, 'Type', '')
                text = getattr(child, 'Text', '').strip()

                if ctype == 'GuiButton':
                    buttons.append({
                        "id": self._normalize_element_id(child.Id),
                        "text": text,
                        "tooltip": getattr(child, 'Tooltip', '').strip(),
                    })
                elif ctype in self._POPUP_INTERACTIVE_TYPES:
                    interactive_elements.append({
                        "id": self._normalize_element_id(child.Id),
                        "type": ctype,
                        "name": getattr(child, 'Name', ''),
                        "text": text,
                        "changeable": getattr(child, 'Changeable', None),
                    })
                elif text and ctype in (
                    'GuiTextField', 'GuiCTextField', 'GuiLabel',
                    'GuiTitlebar', 'GuiStatusbar',
                ):
                    texts.append(text)

                if hasattr(child, 'Children'):
                    try:
                        if child.Children.Count > 0:
                            self._collect_popup_contents(
                                child,
                                texts,
                                buttons,
                                interactive_elements,
                                depth + 1,
                            )
                    except Exception:
                        pass
        except Exception:
            pass

    def _popup_text_blob(self, popup: Dict[str, Any]) -> str:
        """Build a lowercase text blob from popup title/message/texts."""
        parts = [
            popup.get("title", ""),
            popup.get("message", ""),
            *popup.get("texts", []),
        ]
        return " ".join(part for part in parts if part).lower()

    def _classify_popup(self, popup: Dict[str, Any]) -> Dict[str, Any]:
        """Classify popup intent and suggest a safe next action."""
        buttons = popup.get("buttons", [])
        text_blob = self._popup_text_blob(popup)
        message_type = popup.get("message_type", "")
        has_inputs = popup.get("has_inputs", False)
        has_confirm = self._match_button(buttons, self._CONFIRM_PATTERNS) is not None
        has_cancel = self._match_button(buttons, self._CANCEL_PATTERNS) is not None

        if message_type == "E" or any(p in text_blob for p in self._POPUP_ERROR_PATTERNS):
            classification = "error"
        elif has_inputs or any(p in text_blob for p in self._POPUP_INPUT_PATTERNS):
            classification = "input_required"
        elif (
            any(p in text_blob for p in self._POPUP_CONFIRMATION_PATTERNS)
            or ("?" in text_blob and (has_confirm or has_cancel))
        ):
            classification = "confirmation"
        elif has_confirm and has_cancel:
            classification = "confirmation"
        elif (
            message_type in ("S", "I")
            or any(p in text_blob for p in self._POPUP_INFO_PATTERNS)
            or (has_confirm and not has_cancel and not has_inputs)
            or (has_cancel and not has_confirm and not has_inputs)
        ):
            classification = "information"
        else:
            classification = "unknown"

        recommended_action = "read"
        safe_auto_action = None
        if classification == "information" and not has_inputs:
            if has_confirm and not has_cancel:
                recommended_action = "confirm"
                safe_auto_action = "confirm"
            elif has_cancel and not has_confirm:
                recommended_action = "cancel"
                safe_auto_action = "cancel"

        popup["classification"] = classification
        popup["recommended_action"] = recommended_action
        if safe_auto_action is not None:
            popup["safe_auto_action"] = safe_auto_action
        return popup

    # =========================================================================
    # Popup Workflow
    # =========================================================================

    _CONFIRM_PATTERNS = ("ok", "yes", "continue", "confirm", "execute", "save")
    _CANCEL_PATTERNS = ("cancel", "no", "abort", "close", "back")

    def _match_button(
        self, buttons: List[Dict[str, Any]], patterns: tuple,
    ) -> Dict[str, Any] | None:
        """Find first button whose text or tooltip matches any pattern."""
        for btn in buttons:
            text_lower = btn.get("text", "").lower()
            tooltip_lower = btn.get("tooltip", "").lower()
            for pattern in patterns:
                if pattern in text_lower or pattern in tooltip_lower:
                    return btn
        return None

    def handle_popup(
        self, action: str = "read", button_text: str = "",
    ) -> Dict[str, Any]:
        """Read and optionally act on the current popup dialog.

        Args:
            action: ``read``, ``confirm``, ``cancel``, ``press``, or ``auto``.
            button_text: required when *action* is ``press``.

        Returns:
            Dict with popup content, classification, and action taken.
        """
        self._require_session()

        popup = self.get_popup_window()
        if not popup.get("popup_exists"):
            return {"popup_exists": False, "action": "none"}

        popup["action_requested"] = action
        popup_id = popup["window_id"]
        all_buttons = popup.get("buttons", [])

        if action == "read":
            popup["action"] = "read"
            return popup

        if action == "auto":
            decision = popup.get("safe_auto_action") or "read"
            popup["auto_decision"] = decision
            if decision == "read":
                popup["action"] = "read"
                return popup
            action = decision

        if action == "press" and not button_text:
            raise ValueError(
                "button_text is required when action='press'"
            )

        # Determine which button to press
        btn = None
        fallback_vkey = None

        if action == "confirm":
            btn = self._match_button(all_buttons, self._CONFIRM_PATTERNS)
            fallback_vkey = 0  # Enter
        elif action == "cancel":
            btn = self._match_button(all_buttons, self._CANCEL_PATTERNS)
            fallback_vkey = 12  # F12
        elif action == "press":
            needle = button_text.lower()
            for b in all_buttons:
                text_l = b.get("text", "").lower()
                tooltip_l = b.get("tooltip", "").lower()
                if needle in text_l or needle in tooltip_l:
                    btn = b
                    break
            if btn is None:
                popup["action"] = "error"
                popup["error"] = (
                    f"No button matching '{button_text}'. "
                    f"Available: {[b.get('text') or b.get('tooltip') for b in all_buttons]}"
                )
                return popup
        else:
            raise ValueError(
                f"Invalid action '{action}'. Use read, confirm, cancel, press, or auto."
            )

        # Press the button or send fallback VKey
        if btn is not None:
            try:
                self._find_element(btn["id"]).press()
            except Exception as e:
                popup["action"] = "error"
                popup["error"] = f"Failed to press button: {e}"
                return popup
            popup["action"] = action + "ed" if action != "press" else "pressed"
            popup["button_pressed"] = btn.get("text") or btn.get("tooltip")
        else:
            # Fallback: send VKey to the popup window
            try:
                self._find_window(popup_id).sendVKey(fallback_vkey)
            except Exception as e:
                popup["action"] = "error"
                popup["error"] = f"Failed to send key: {e}"
                return popup
            popup["action"] = action + "ed"
            popup["button_pressed"] = (
                "Enter (fallback)" if fallback_vkey == 0 else "F12 (fallback)"
            )

        # The popup that was acted on is gone, so its element IDs are dead
        # weight. Keep the trail of what it said and which values were
        # accepted: changeable inputs only, not read-only description fields.
        inputs = {
            (el.get("name") or el.get("id", "")): self._mask_field_value(
                el.get("id", ""), el["text"],
            )
            for el in popup.pop("interactive_elements", [])
            if el.get("type") in self._PREFILL_TYPES
            and el.get("changeable")
            and el.get("text")
        }
        if inputs:
            popup["inputs"] = inputs
        for key in ("buttons", "has_inputs", "recommended_action",
                    "safe_auto_action", "prefilled_inputs", "notice"):
            popup.pop(key, None)

        # Return updated screen state after the action. A follow-up popup is
        # reported in full as popup_after, so the digest would only repeat it.
        try:
            popup["screen"] = self.get_screen_info(include_popup=False)
        except Exception:
            pass

        try:
            popup_after = self.get_popup_window()
            popup["popup_closed"] = not popup_after.get("popup_exists", False)
            if not popup["popup_closed"]:
                # A follow-up popup is the new state: that one stays complete.
                popup["popup_after"] = popup_after
        except Exception:
            pass

        return popup

    # =========================================================================
    # Toolbar Discovery
    # =========================================================================

    def get_toolbar_buttons(self, window_id: str = "wnd[0]") -> Dict[str, Any]:
        """
        List all toolbar buttons on a window's application toolbar.

        Reads buttons from tbar[0] (system toolbar) and tbar[1] (application
        toolbar).  Returns button IDs, text, tooltip, and enabled state.

        Args:
            window_id: Window ID (default "wnd[0]")

        Returns:
            Dict with toolbar button info
        """
        self._require_session()

        normalized_window_id = self._validate_window_id(window_id)
        toolbars: Dict[str, list] = {}
        for tbar_idx, tbar_name in [(0, "system_toolbar"), (1, "application_toolbar")]:
            buttons = []
            try:
                tbar = self._find_element(f"{normalized_window_id}/tbar[{tbar_idx}]")
                for i in range(tbar.Children.Count):
                    btn = tbar.Children(i)
                    btype = getattr(btn, 'Type', '')
                    if btype in ('GuiButton',):
                        buttons.append({
                            "id": self._normalize_element_id(btn.Id),
                            "text": getattr(btn, 'Text', '').strip(),
                            "tooltip": getattr(btn, 'Tooltip', '').strip(),
                            "enabled": getattr(btn, 'Changeable', True) is not False,
                        })
            except Exception:
                pass
            if buttons:
                toolbars[tbar_name] = buttons

        return {
            "window_id": normalized_window_id,
            "toolbars": toolbars,
        }

    # =========================================================================
    # Shell Content Reading
    # =========================================================================

    def read_shell_content(self, shell_id: str) -> Dict[str, Any]:
        """
        Read content from a GuiShell subtype (HTMLViewer, etc.).

        Attempts to extract useful content based on the shell's SubType.
        Supports GuiHTMLViewer (BrowserHandle -> InnerHTML), and falls
        back to generic Text property.

        Args:
            shell_id: SAP GUI shell element ID

        Returns:
            Dict with shell content and metadata
        """
        self._require_session()

        try:
            shell = self._find_element(shell_id)
            shell_type = getattr(shell, 'Type', '')
            sub_type = getattr(shell, 'SubType', '')

            result: Dict[str, Any] = {
                "shell_id": shell_id,
                "type": shell_type,
                "sub_type": sub_type,
            }

            # Try SubType-specific extraction
            if sub_type == "HTMLViewer":
                try:
                    result["inner_html"] = shell.InnerHTML
                except Exception:
                    pass
                try:
                    result["url"] = shell.CurrentUrl
                except Exception:
                    pass

            # Generic fallback: Text property
            try:
                text = getattr(shell, 'Text', None)
                if text is not None:
                    result["text"] = str(text)[:5000]
            except Exception:
                pass

            return result
        except Exception as e:
            return self._error_result(
                {"shell_id": shell_id},
                e,
                "Could not read shell content",
            )

    # =========================================================================
    # Classic List Reading
    # =========================================================================

    # ABAP list colours (FORMAT COLOR). 0 and 2 are the plain background /
    # normal text and carry no meaning worth reporting.
    _LIST_COLORS = {
        1: "heading", 3: "total", 4: "key",
        5: "positive", 6: "negative", 7: "group",
    }
    _LIST_CELL_RE = re.compile(r"/([a-z]+)\[(\d+),(\d+)\]$")

    def read_list(self, window_id: str = "wnd[0]", max_lines: int = 200,
                  scroll_to: int = -1, with_ids: bool = False) -> Dict[str, Any]:
        """
        Read a classic ABAP list (WRITE output) as lines of text.

        A classic list (report output, program SAPMSSY0; also many F4 hit
        lists) has no table object: every word is a GuiLabel named
        ``lbl[col,row]``. Discovery returns one element per label (about 500
        for a single page); this reassembles them into the lines a user sees.

        Args:
            window_id: Window holding the list (a popup for F4 hit lists)
            max_lines: Maximum number of lines to return
            scroll_to: Vertical scrollbar position to move to first; -1 reads
                the page that is currently shown
            with_ids: Also return, per line, the ID of its first label (to set
                the focus on a line before F2 / double-click)

        Returns:
            Dict with the lines of the visible page, colours of the lines that
            have a semantic one, and the scroll state for paging
        """
        self._require_session()

        try:
            window_id = self._validate_window_id(window_id)
            usr_id = f"{window_id}/usr"
            usr = self._session.findById(usr_id)

            if scroll_to >= 0:
                usr.VerticalScrollbar.Position = scroll_to
                # Scrolling re-renders the page: the old reference is stale.
                usr = self._session.findById(usr_id)

            rows: Dict[int, list] = {}
            for i in range(usr.Children.Count):
                child = usr.Children(i)
                match = self._LIST_CELL_RE.search(child.Id)
                if not match:
                    continue
                kind, col, row = match.group(1), int(match.group(2)), int(match.group(3))
                if kind == "chk":
                    text = "[x]" if self._list_prop(child, "Selected", False) else "[ ]"
                    color = None
                else:
                    text = str(self._list_prop(child, "Text", ""))
                    color = self._LIST_COLORS.get(
                        self._list_prop(child, "ColorIndex", 0)
                    ) if kind == "lbl" else None
                rows.setdefault(row, []).append((col, text, color, child.Id))

            if not rows:
                return {
                    "window": window_id,
                    "is_list": False,
                    "lines": [],
                    "note": (
                        "No list labels on this screen. For ALV grids and table "
                        "controls use read_table."
                    ),
                }

            first_row = min(rows)
            last_row = min(max(rows), first_row + max(max_lines, 1) - 1)
            lines: List[str] = []
            colors: Dict[str, List[str]] = {}
            ids: Dict[str, str] = {}
            for row in range(first_row, last_row + 1):
                line = ""
                row_colors: List[str] = []
                cells = sorted(rows.get(row, []), key=lambda cell: cell[0])
                for col, text, color, _cell_id in cells:
                    line = line.ljust(col) + text
                    if color and color not in row_colors:
                        row_colors.append(color)
                lines.append(line.rstrip())
                if row_colors:
                    colors[str(row)] = row_colors
                if with_ids and cells:
                    ids[str(row)] = self._normalize_element_id(cells[0][3])

            result: Dict[str, Any] = {
                "window": window_id,
                "is_list": True,
                "first_row": first_row,
                "lines": lines,
            }
            if colors:
                result["colors"] = colors
            if with_ids:
                result["line_ids"] = ids
            if max(rows) > last_row:
                result["truncated"] = True
            scroll = self._list_scroll_state(usr)
            if scroll:
                result["scroll"] = scroll
            return result

        except Exception as e:
            return self._error_result({"window": window_id}, e, "Could not read list")

    @staticmethod
    def _list_prop(obj, name: str, default):
        """Read a COM property that may be unsupported on this element."""
        try:
            return getattr(obj, name)
        except Exception:
            return default

    def _list_scroll_state(self, usr) -> Dict[str, Any]:
        """Vertical (and, when the list is wider than the window, horizontal)
        scrollbar state, so the caller can page with scroll_to."""
        state: Dict[str, Any] = {}
        for key, attr in (("vertical", "VerticalScrollbar"),
                          ("horizontal", "HorizontalScrollbar")):
            try:
                bar = getattr(usr, attr)
                maximum = bar.Maximum
                if maximum > 0:
                    state[key] = {
                        "position": bar.Position,
                        "maximum": maximum,
                        "page_size": bar.PageSize,
                    }
            except Exception:
                continue
        return state

    # =========================================================================
    # Screen Element Discovery
    # =========================================================================

    def get_screen_elements(self, container_id: str = "wnd[0]/usr",
                            max_depth: int = 3,
                            type_filter: str = "",
                            changeable_only: bool = False,
                            expand_tables: bool = False) -> List[ScreenElement]:
        """
        Enumerate all elements on the current screen.

        Useful for discovering field IDs when automating a new transaction.

        Args:
            container_id: Starting container (default: main user area). A
                bare window ("wnd[0]") lists its top-level children, incl.
                docking containers that live beside the user area.
            max_depth: Maximum recursion depth
            type_filter: Comma-separated SAP element types to include
                (e.g. "GuiTextField,GuiCTextField"). Empty = all types.
            changeable_only: If True, only return editable/input elements
            expand_tables: If True, also list every visible cell of a
                GuiTableControl. Off by default: an empty two-column table is
                76 elements, and read_table / get_column_info describe the
                columns (incl. a cell-ID template) far more compactly.

        Returns:
            List of ScreenElement objects
        """
        self._require_session()

        type_filter_set = None
        if type_filter:
            type_filter_set = {t.strip() for t in type_filter.split(",") if t.strip()}

        try:
            container = self._session.findById(
                self._validate_container_id(container_id)
            )
            elements = self._enumerate_elements(
                container, max_depth,
                type_filter_set=type_filter_set,
                changeable_only=changeable_only,
                expand_tables=expand_tables,
            )
            return elements
        except ValueError:
            raise
        except Exception as e:
            logger.warning(
                "Failed to enumerate elements in %s: %s",
                container_id,
                e,
                exc_info=logger.isEnabledFor(logging.DEBUG),
            )
            raise SAPGUIError(f"Failed to enumerate elements in '{container_id}'")

    def get_docking_containers(self, window_id: str = "wnd[0]") -> List[str]:
        """List a window's docking containers as short IDs.

        Docking containers (``wnd[0]/shellcont``, ``wnd[0]/shellcont[1]``) are
        siblings of the user area, so discovery under ``wnd[0]/usr`` never
        reaches them. They hold the dialog-structure tree of view clusters
        (SM34, most IMG activities) and the SE80 repository tree.
        """
        self._require_session()

        try:
            window = self._find_window(window_id)
            containers = []
            for i in range(window.Children.Count):
                child_id = self._normalize_element_id(window.Children(i).Id)
                if child_id.rsplit("/", 1)[-1].startswith("shellcont"):
                    containers.append(child_id)
            return containers
        except Exception as e:
            logger.debug("Could not list docking containers of %s: %s", window_id, e)
            return []

    def _enumerate_elements(self, container, max_depth: int,
                            current_depth: int = 0,
                            type_filter_set: set = None,
                            changeable_only: bool = False,
                            expand_tables: bool = False) -> List[ScreenElement]:
        """Recursively enumerate screen elements."""
        elements = []

        if current_depth >= max_depth:
            return elements

        def prop(obj, name, default):
            # COM raises more than AttributeError for an unsupported property
            # (GuiTitlebar.Changeable -> IndexError), which getattr's default
            # does not absorb; one such child must not hide its siblings.
            try:
                return getattr(obj, name)
            except Exception:
                return default

        try:
            for i in range(container.Children.Count):
                child = container.Children(i)

                element = ScreenElement(
                    id=self._normalize_element_id(child.Id),
                    type=child.Type,
                    name=prop(child, 'Name', ''),
                    text=str(prop(child, 'Text', ''))[:200],
                    changeable=prop(child, 'Changeable', False),
                    visible=prop(child, 'Visible', True),
                )

                # Apply filters — but always recurse into containers
                include = True
                if type_filter_set and element.type not in type_filter_set:
                    include = False
                if changeable_only and not element.changeable:
                    include = False
                if include:
                    elements.append(element)

                # A table control stays ONE element: its cells are read with
                # the table tools, not dumped one element per cell.
                if element.type == "GuiTableControl" and not expand_tables:
                    continue

                # Recurse into containers regardless of filters
                if hasattr(child, 'Children') and child.Children.Count > 0:
                    child_elements = self._enumerate_elements(
                        child, max_depth, current_depth + 1,
                        type_filter_set=type_filter_set,
                        changeable_only=changeable_only,
                        expand_tables=expand_tables,
                    )
                    elements.extend(child_elements)

        except Exception as e:
            logger.debug(f"Error enumerating at depth {current_depth}: {e}")

        return elements

    # =========================================================================
    # Screenshot & Visual
    # =========================================================================

    def _find_topmost_window(self) -> str:
        """Find the topmost SAP GUI window (highest wnd index that exists).

        Popups appear as wnd[1], wnd[2], etc. This returns the topmost
        window so screenshots and screen reads capture what the user sees.

        Tries Session.ActiveWindow first (faster), falls back to loop.
        """
        try:
            active = self._session.ActiveWindow
            if active is not None:
                return self._normalize_window_id(active.Id)
        except Exception:
            pass

        topmost = "wnd[0]"
        for i in range(1, 10):
            try:
                self._session.findById(f"wnd[{i}]")
                topmost = f"wnd[{i}]"
            except Exception:
                break
        return topmost

    def take_screenshot(self, filepath: str = None) -> Dict[str, Any]:
        """
        Take a screenshot of the current SAP window.

        Args:
            filepath: Optional file path. If not provided, returns base64 data.

        Returns:
            Dict with filepath or base64 encoded image data
        """
        self._require_session()

        import os
        import tempfile

        temp_filepath = None
        try:
            if filepath is None:
                temp_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    prefix="sap_screenshot_",
                    suffix=".png",
                )
                temp_filepath = temp_file.name
                temp_file.close()
                filepath = temp_filepath
                return_base64 = True
            else:
                return_base64 = False

            # Find the topmost window (popups are wnd[1], wnd[2], etc.)
            window_id = self._find_topmost_window()
            window = self._find_window(window_id)
            window.HardCopy(filepath, _HARDCOPY_PNG)

            # Optimize image size with Pillow if available
            self._optimize_screenshot(filepath)

            if return_base64:
                import base64
                with open(filepath, "rb") as f:
                    data = base64.b64encode(f.read()).decode()
                os.remove(filepath)
                return {
                    "format": "png",
                    "encoding": "base64",
                    "window": window_id,
                    "data": data,
                }
            else:
                return {
                    "format": "png",
                    "filepath": filepath,
                    "window": window_id,
                }

        except Exception as e:
            if temp_filepath and os.path.exists(temp_filepath):
                os.remove(temp_filepath)
            return self._error_result({}, e, "Could not capture screenshot")

    def save_screenshot(self, save_path: str, inline: bool = True) -> Dict[str, Any]:
        """
        Save a screenshot of the current SAP window to a PNG file.

        The file keeps the full resolution SAP GUI renders (the inline image
        is downscaled to 1920px), which is what a document needs. An existing
        file is never overwritten.

        Args:
            save_path: Target file; must end in ".png" and its directory must
                exist. Relative paths resolve against the server's working
                directory.
            inline: Also return the (downscaled) image as base64 ``data``.

        Returns:
            Dict with the absolute filepath, pixel size, window and, when
            *inline* is set, the base64 image data
        """
        self._require_session()

        import base64
        import os
        import shutil
        import tempfile

        temp_filepath = None
        try:
            target = self._resolve_screenshot_path(save_path)

            window_id = self._find_topmost_window()
            self._find_window(window_id).HardCopy(target, _HARDCOPY_PNG)
            if not os.path.isfile(target):
                raise SAPGUIError("SAP GUI did not write the screenshot file")

            result: Dict[str, Any] = {
                "format": "png",
                "filepath": target,
                "window": window_id,
                "size_bytes": os.path.getsize(target),
            }
            size = self._png_size(target)
            if size is not None:
                result["width"], result["height"] = size

            if inline:
                # Optimize a copy: the saved file stays at full resolution.
                temp_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    prefix="sap_screenshot_",
                    suffix=".png",
                )
                temp_filepath = temp_file.name
                temp_file.close()
                shutil.copyfile(target, temp_filepath)
                self._optimize_screenshot(temp_filepath)
                with open(temp_filepath, "rb") as f:
                    result["data"] = base64.b64encode(f.read()).decode()

            return result

        except Exception as e:
            return self._error_result({}, e, "Could not save screenshot")
        finally:
            if temp_filepath and os.path.exists(temp_filepath):
                os.remove(temp_filepath)

    def _resolve_screenshot_path(self, save_path: Any) -> str:
        """Validate a screenshot target and return it as an absolute path."""
        import os

        if not isinstance(save_path, str) or not save_path.strip():
            raise ValueError("save_path must be a file path ending in '.png'")

        target = os.path.abspath(os.path.expanduser(save_path.strip()))
        if not target.lower().endswith(".png"):
            raise ValueError(f"save_path must end in '.png': {save_path!r}")
        if not os.path.isdir(os.path.dirname(target)):
            raise ValueError(
                f"Directory does not exist: {os.path.dirname(target)!r}"
            )
        if os.path.exists(target):
            raise ValueError(
                f"File already exists: {target!r}. Screenshots never "
                "overwrite a file; choose another name."
            )
        return target

    @staticmethod
    def _png_size(filepath: str) -> tuple[int, int] | None:
        """Read (width, height) from a PNG header without needing Pillow."""
        import struct

        try:
            with open(filepath, "rb") as f:
                header = f.read(24)
        except OSError:
            return None
        if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
            return None
        width, height = struct.unpack(">II", header[16:24])
        return width, height

    def _optimize_screenshot(self, filepath: str) -> None:
        """
        Optimize screenshot file size using Pillow if available.

        Resizes large images and applies PNG optimization to significantly
        reduce file size (typically 70-90% reduction).
        """
        try:
            from PIL import Image
        except ImportError:
            logger.debug("Pillow not installed, skipping screenshot optimization")
            return

        try:
            img = Image.open(filepath)

            # Downscale if image is very large (> 1920px wide)
            max_width = 1920
            if img.width > max_width:
                ratio = max_width / img.width
                new_size = (max_width, int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)

            # Convert RGBA to RGB if no transparency (smaller file)
            if img.mode == "RGBA":
                background = Image.new("RGB", img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background

            # Save with optimization
            img.save(filepath, "PNG", optimize=True)
        except Exception as e:
            logger.debug(f"Screenshot optimization failed (using original): {e}")
