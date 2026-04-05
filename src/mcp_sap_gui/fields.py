"""
Fields Mixin - field read/write, buttons, checkboxes, combos, textedit.

Provides all field-level operations for the SAP GUI controller.
"""

import logging
from typing import Any, Dict, List

logger = logging.getLogger(__name__)


class FieldsMixin:
    """Mixin for field operations on SAP GUI screens."""

    # =========================================================================
    # Field Operations
    # =========================================================================

    def read_field(self, field_id: str) -> Dict[str, Any]:
        """
        Read a field value from the screen.

        Returns the field value plus metadata like required, max_length,
        numerical flag, and associated labels (when available from
        GuiTextField / GuiCTextField).

        Args:
            field_id: SAP GUI element ID (e.g., "wnd[0]/usr/txtMATNR")

        Returns:
            Dict with field value and properties
        """
        self._require_session()

        try:
            element = self._session.findById(field_id)

            result: Dict[str, Any] = {
                "field_id": field_id,
                "value": getattr(element, 'Text', ''),
                "type": element.Type,
                "name": getattr(element, 'Name', ''),
                "changeable": getattr(element, 'Changeable', None),
            }

            # Extended metadata for text fields
            for attr, key in [
                ("Required", "required"),
                ("MaxLength", "max_length"),
                ("Numerical", "numerical"),
                ("Highlighted", "highlighted"),
            ]:
                try:
                    val = getattr(element, attr, None)
                    if val is not None:
                        result[key] = val
                except Exception:
                    pass

            # Associated labels (GuiTextField / GuiCTextField)
            for attr, key in [
                ("LeftLabel", "left_label"),
                ("RightLabel", "right_label"),
            ]:
                try:
                    label = getattr(element, attr, None)
                    if label is not None:
                        result[key] = getattr(label, 'Text', str(label))
                except Exception:
                    pass

            return result
        except Exception as e:
            return self._error_result(
                {"field_id": field_id},
                e,
                "Could not read field",
            )

    def set_field(self, field_id: str, value: str) -> Dict[str, Any]:
        """
        Set a field value on the screen.

        Args:
            field_id: SAP GUI element ID
            value: Value to set

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            element = self._session.findById(field_id)
            element.text = value

            logger.debug("Set %s = %s", field_id, self._mask_field_value(field_id, value))
            return {
                "field_id": field_id,
                "value": value,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"field_id": field_id},
                e,
                "Could not set field",
            )

    def press_button(self, button_id: str) -> Dict[str, Any]:
        """
        Press a button on the screen.

        Args:
            button_id: SAP GUI button ID (e.g., "wnd[0]/tbar[1]/btn[8]")

        Returns:
            Dict with screen info after button press
        """
        self._require_session()

        try:
            self._session.findById(button_id).press()

            logger.debug(f"Pressed button: {button_id}")
            return {
                "button_id": button_id,
                "status": "pressed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"button_id": button_id},
                e,
                "Could not press button",
            )

    def select_menu(self, menu_id: str) -> Dict[str, Any]:
        """
        Select a menu item from the menu bar or a submenu.

        Args:
            menu_id: SAP GUI menu ID (e.g., 'wnd[0]/mbar/menu[3]/menu[0]')

        Returns:
            Dict with screen info after menu selection
        """
        self._require_session()

        try:
            self._session.findById(menu_id).Select()

            return {
                "menu_id": menu_id,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"menu_id": menu_id},
                e,
                "Could not select menu item",
            )

    def select_checkbox(self, checkbox_id: str, selected: bool = True) -> Dict[str, Any]:
        """
        Select or deselect a checkbox.

        Args:
            checkbox_id: SAP GUI checkbox ID
            selected: True to select, False to deselect
        """
        self._require_session()

        try:
            element = self._session.findById(checkbox_id)
            element.selected = selected

            return {
                "checkbox_id": checkbox_id,
                "selected": selected,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"checkbox_id": checkbox_id},
                e,
                "Could not change checkbox",
            )

    def select_radio_button(self, radio_id: str) -> Dict[str, Any]:
        """
        Select a radio button.

        Args:
            radio_id: SAP GUI radio button ID (e.g., 'wnd[0]/usr/radOPT1')

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            element = self._session.findById(radio_id)
            element.Select()

            return {
                "radio_id": radio_id,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"radio_id": radio_id},
                e,
                "Could not select radio button",
            )

    def select_combobox_entry(self, combobox_id: str, key_or_value: str) -> Dict[str, Any]:
        """
        Select an entry in a combobox/dropdown.

        First tries to set the key directly. If that fails, searches the
        Entries collection by value text.

        Args:
            combobox_id: SAP GUI combobox ID (e.g., 'wnd[0]/usr/cmbLANGU')
            key_or_value: Key or display value text of the entry to select

        Returns:
            Dict with result status and selected key/value
        """
        self._require_session()

        try:
            combobox = self._session.findById(combobox_id)

            # Try setting key directly first
            try:
                combobox.Key = key_or_value
                return {
                    "combobox_id": combobox_id,
                    "key": key_or_value,
                    "status": "success",
                }
            except Exception:
                pass

            # Fallback: search Entries by value text
            entries = combobox.Entries
            for i in range(entries.Count):
                entry = entries(i)
                if entry.Value == key_or_value or entry.Key == key_or_value:
                    combobox.Key = entry.Key
                    return {
                        "combobox_id": combobox_id,
                        "key": entry.Key,
                        "value": entry.Value,
                        "status": "success",
                    }

            return {
                "combobox_id": combobox_id,
                "error": f"Entry '{key_or_value}' not found in combobox",
            }
        except Exception as e:
            return self._error_result(
                {"combobox_id": combobox_id},
                e,
                "Could not select combobox entry",
            )

    def select_tab(self, tab_id: str) -> Dict[str, Any]:
        """
        Select a tab in a tab strip.

        Args:
            tab_id: SAP GUI tab ID (e.g., 'wnd[0]/usr/tabsTAB/tabpTAB1')

        Returns:
            Dict with result status and screen info after selection
        """
        self._require_session()

        try:
            tab = self._session.findById(tab_id)
            tab.Select()

            return {
                "tab_id": tab_id,
                "status": "success",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"tab_id": tab_id},
                e,
                "Could not select tab",
            )

    def get_combobox_entries(self, combobox_id: str) -> Dict[str, Any]:
        """
        List all entries in a combobox/dropdown.

        Returns the available key-value pairs so the caller knows which
        values are valid before attempting to set one.

        Args:
            combobox_id: SAP GUI combobox ID

        Returns:
            Dict with list of entries (key, value pairs)
        """
        self._require_session()

        try:
            combo = self._session.findById(combobox_id)
            entries = []
            for i in range(combo.Entries.Count):
                entry = combo.Entries(i)
                entries.append({
                    "key": entry.Key,
                    "value": entry.Value,
                })
            return {
                "combobox_id": combobox_id,
                "current_key": getattr(combo, 'Key', ''),
                "entry_count": len(entries),
                "entries": entries,
            }
        except Exception as e:
            return self._error_result(
                {"combobox_id": combobox_id},
                e,
                "Could not read combobox entries",
            )

    def set_batch_fields(
        self,
        fields: Dict[str, str],
        *,
        skip_readonly: bool = False,
        validate: bool = False,
    ) -> Dict[str, Any]:
        """
        Set multiple field values at once.

        More efficient than calling set_field repeatedly -- all values are
        set before a single round-trip return.

        Args:
            fields: Dict mapping field_id -> value
            skip_readonly: When True, silently skip fields whose element
                reports ``Changeable == False`` instead of counting them
                as failures.
            validate: When True, press Enter after setting fields and
                include validation feedback (status bar, highlighted
                fields) in the result.

        Returns:
            Dict with per-field results and optional validation info
        """
        self._require_session()

        results: Dict[str, str] = {}
        skipped = 0
        for field_id, value in fields.items():
            try:
                element = self._session.findById(field_id)
                if skip_readonly and not getattr(element, "Changeable", True):
                    results[field_id] = "skipped: read-only"
                    skipped += 1
                    continue
                element.text = value
                results[field_id] = "success"
            except Exception as e:
                results[field_id] = (
                    "error: "
                    + self._sanitize_error_message(e, "Could not set field")
                )

        succeeded = sum(1 for v in results.values() if v == "success")
        failed = len(fields) - succeeded - skipped
        result: Dict[str, Any] = {
            "total": len(fields),
            "succeeded": succeeded,
            "failed": failed,
            "skipped": skipped,
            "results": results,
        }

        if validate:
            succeeded_ids = [
                fid for fid, status in results.items() if status == "success"
            ]
            result["validation"] = self._validate_batch(succeeded_ids, succeeded)

        return result

    def _validate_batch(
        self,
        succeeded_ids: List[str],
        succeeded: int,
    ) -> Dict[str, Any]:
        """Press Enter and gather validation feedback for a batch-set operation."""
        if succeeded == 0:
            return {"performed": False, "reason": "no fields were set"}

        enter_result = self.press_enter()
        screen_info = enter_result.get("screen", {})

        # Check highlighted fields only among those actually set successfully
        highlighted: List[str] = []
        for field_id in succeeded_ids:
            try:
                element = self._session.findById(field_id)
                if getattr(element, "Highlighted", False):
                    highlighted.append(field_id)
            except Exception:
                pass

        validation: Dict[str, Any] = {
            "performed": True,
            "message": screen_info.get("message"),
            "message_type": screen_info.get("message_type", ""),
            "screen": screen_info,
        }
        if highlighted:
            validation["highlighted_fields"] = highlighted
        return validation

    def read_textedit(self, textedit_id: str, max_lines: int = 0) -> Dict[str, Any]:
        """
        Read the content of a multiline text editor (GuiTextedit).

        Args:
            textedit_id: SAP GUI textedit ID
            max_lines: Maximum lines to return (0 = all lines)

        Returns:
            Dict with full text and line count
        """
        self._require_session()

        try:
            textedit = self._session.findById(textedit_id)
            line_count = textedit.LineCount

            read_count = line_count
            truncated = False
            if max_lines > 0 and line_count > max_lines:
                read_count = max_lines
                truncated = True

            lines = []
            for i in range(read_count):
                try:
                    lines.append(textedit.GetLineText(i))
                except Exception:
                    lines.append("")

            result: Dict[str, Any] = {
                "textedit_id": textedit_id,
                "line_count": read_count,
                "text": "\n".join(lines),
                "changeable": getattr(textedit, 'Changeable', None),
            }
            if truncated:
                result["truncated"] = True
                result["total_lines"] = line_count
            return result
        except Exception as e:
            return self._error_result(
                {"textedit_id": textedit_id},
                e,
                "Could not read text editor",
            )

    def set_textedit(self, textedit_id: str, text: str) -> Dict[str, Any]:
        """
        Set the content of a multiline text editor (GuiTextedit).

        Attempts to set via the Text property first, then falls back
        to SetUnprotectedTextPart for protected editors.

        Args:
            textedit_id: SAP GUI textedit ID
            text: Text content to set

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            textedit = self._session.findById(textedit_id)
            try:
                textedit.Text = text
            except Exception:
                # Fallback: set only the unprotected part
                textedit.SetUnprotectedTextPart(text)

            return {
                "textedit_id": textedit_id,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"textedit_id": textedit_id},
                e,
                "Could not set text editor",
            )

    def set_focus(self, element_id: str) -> Dict[str, Any]:
        """
        Set focus to any screen element.

        Args:
            element_id: SAP GUI element ID

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            self._session.findById(element_id).SetFocus()
            return {"element_id": element_id, "status": "success"}
        except Exception as e:
            return self._error_result(
                {"element_id": element_id},
                e,
                "Could not set focus",
            )
