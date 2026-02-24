"""
Fields Mixin - field read/write, buttons, checkboxes, combos, textedit.

Provides all field-level operations for the SAP GUI controller.
"""

import logging
from typing import Any, Dict

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
            return {"field_id": field_id, "error": str(e)}

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

            logger.debug(f"Set {field_id} = {value}")
            return {
                "field_id": field_id,
                "value": value,
                "status": "success",
            }
        except Exception as e:
            return {"field_id": field_id, "error": str(e)}

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
            return {"button_id": button_id, "error": str(e)}

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
            return {"menu_id": menu_id, "error": str(e)}

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
            return {"checkbox_id": checkbox_id, "error": str(e)}

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
            return {"radio_id": radio_id, "error": str(e)}

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
            return {"combobox_id": combobox_id, "error": str(e)}

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
            return {"tab_id": tab_id, "error": str(e)}

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
            return {"combobox_id": combobox_id, "error": str(e)}

    def set_batch_fields(self, fields: Dict[str, str]) -> Dict[str, Any]:
        """
        Set multiple field values at once.

        More efficient than calling set_field repeatedly -- all values are
        set before a single round-trip return.

        Args:
            fields: Dict mapping field_id -> value

        Returns:
            Dict with per-field results
        """
        self._require_session()

        results = {}
        for field_id, value in fields.items():
            try:
                self._session.findById(field_id).text = value
                results[field_id] = "success"
            except Exception as e:
                results[field_id] = f"error: {e}"

        succeeded = sum(1 for v in results.values() if v == "success")
        return {
            "total": len(fields),
            "succeeded": succeeded,
            "failed": len(fields) - succeeded,
            "results": results,
        }

    def read_textedit(self, textedit_id: str) -> Dict[str, Any]:
        """
        Read the content of a multiline text editor (GuiTextedit).

        Args:
            textedit_id: SAP GUI textedit ID

        Returns:
            Dict with full text, line count, and individual lines
        """
        self._require_session()

        try:
            textedit = self._session.findById(textedit_id)
            line_count = textedit.LineCount
            lines = []
            for i in range(line_count):
                try:
                    lines.append(textedit.GetLineText(i))
                except Exception:
                    lines.append("")

            return {
                "textedit_id": textedit_id,
                "line_count": line_count,
                "text": "\n".join(lines),
                "lines": lines,
                "changeable": getattr(textedit, 'Changeable', None),
            }
        except Exception as e:
            return {"textedit_id": textedit_id, "error": str(e)}

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
            return {"textedit_id": textedit_id, "error": str(e)}

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
            return {"element_id": element_id, "error": str(e)}
