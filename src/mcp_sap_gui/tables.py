"""
Tables Mixin - ALV grid and GuiTableControl operations.

Provides table reading, row selection, cell operations, and toolbar
interaction for both ALV grids (GuiGridView) and classic table controls
(GuiTableControl).
"""

import logging
from typing import Any, Dict, List, Optional

from .models import _TOOLBAR_BUTTON_TYPES, VKey

logger = logging.getLogger(__name__)


class TablesMixin:
    """Mixin for table/grid operations on SAP GUI screens."""

    # =========================================================================
    # Table/Grid Operations
    # =========================================================================

    # ---- GuiTableControl helpers ----

    def _get_table_control_columns(self, table) -> list:
        """Get column metadata from a GuiTableControl's Columns collection.

        Each column entry has: index, name, title, tooltip.

        Note: The SAP GUI Scripting API docs state that GuiTableColumn members
        in the Columns collection "do not support properties like id or name".
        Column names are therefore extracted from the first row's cell Name
        property, with Title as fallback.
        """
        columns = []
        col_count = table.Columns.Count

        # Get column names from first row's cells (safer than col.Name
        # which is documented as unsupported on GuiTableColumn)
        cell_names = []
        for i in range(col_count):
            name = None
            try:
                cell = table.GetCell(0, i)
                name = getattr(cell, 'Name', None)
            except Exception:
                pass
            cell_names.append(name)

        # Get column titles and tooltips from the Columns collection
        for i in range(col_count):
            info = {"index": i}
            try:
                col = table.Columns(i)
                try:
                    info["title"] = col.Title
                except Exception:
                    info["title"] = ""
                try:
                    info["tooltip"] = col.Tooltip
                except Exception:
                    info["tooltip"] = ""
            except Exception:
                info["title"] = ""
                info["tooltip"] = ""

            info["name"] = cell_names[i] or info.get("title") or f"col_{i}"
            columns.append(info)
        return columns

    def _read_cell_value(self, cell):
        """Read a cell value, handling different element types."""
        try:
            cell_type = getattr(cell, 'Type', '')
            if cell_type == "GuiCheckBox":
                return bool(cell.Selected)
            elif cell_type == "GuiComboBox":
                return getattr(cell, 'Key', cell.Text)
            else:
                return cell.Text
        except Exception:
            return None

    def _scroll_table_control_to_row(self, table, abs_row: int) -> int:
        """Scroll a GuiTableControl so *abs_row* is visible.

        Returns the **visible-row offset** to pass to ``GetCell()``.
        ``GetCell`` uses visible-row indexing (0 = first visible row),
        so callers must use the returned offset, not the original
        *abs_row*.
        """
        visible = table.VisibleRowCount
        scrollbar = table.VerticalScrollbar
        current_top = scrollbar.Position

        if current_top <= abs_row < current_top + visible:
            return abs_row - current_top  # already visible

        new_pos = max(scrollbar.Minimum, min(abs_row, scrollbar.Maximum))
        scrollbar.Position = new_pos
        return abs_row - new_pos

    def _resolve_table_control_column(self, table, column) -> int:
        """Resolve a column name/index to a numeric column index for GuiTableControl."""
        if isinstance(column, int):
            return column
        if isinstance(column, str) and column.isdigit():
            return int(column)

        # Search by cell Name (from first row) and column Title
        for i in range(table.Columns.Count):
            # Try cell Name first
            try:
                cell = table.GetCell(0, i)
                if getattr(cell, 'Name', None) == column:
                    return i
            except Exception:
                pass
            # Try column Title
            try:
                col = table.Columns(i)
                if col.Title == column:
                    return i
            except Exception:
                pass

        raise ValueError(f"Column '{column}' not found in table")

    # ---- Table reading ----

    def read_table(self, table_id: str, max_rows: int = 100,
                   columns: str = "", columns_only: bool = False,
                   start_row: int = 0) -> Dict[str, Any]:
        """
        Read data from an ALV grid or table control.

        Automatically detects GuiGridView (ALV) vs GuiTableControl and
        uses the appropriate API for each.

        Args:
            table_id: SAP GUI table/grid ID
            max_rows: Maximum rows to read (default 100)
            columns: Comma-separated column names to include (empty = all).
                Reduces response size when only a few columns are needed.
            columns_only: If True, return only column metadata with empty
                data (schema discovery mode — very small response).
            start_row: Row index to start reading from (default 0).
                For ALV: skips first N rows. For TableControl: scrolls to
                that position before reading.

        Returns:
            Dict with table data and metadata
        """
        self._require_session()

        col_filter = []
        if columns:
            col_filter = [c.strip() for c in columns.split(",") if c.strip()]

        try:
            table = self._find_element(table_id)
            if getattr(table, 'Type', '') == "GuiTableControl":
                return self._read_table_control(
                    table, table_id, max_rows, col_filter, columns_only, start_row,
                )
            return self._read_alv_grid(
                table, table_id, max_rows, col_filter, columns_only, start_row,
            )
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not read table",
            )

    def _read_alv_grid(self, grid, table_id: str, max_rows: int,
                       col_filter: List[str] = None,
                       columns_only: bool = False,
                       start_row: int = 0) -> Dict[str, Any]:
        """Read data from an ALV grid (GuiGridView)."""
        all_columns = []
        for i in range(grid.ColumnCount):
            all_columns.append(grid.ColumnOrder(i))

        all_column_info = []
        for col in all_columns:
            info = {"name": col}
            try:
                info["tooltip"] = grid.GetColumnTooltip(col)
            except Exception:
                info["tooltip"] = ""
            try:
                info["title"] = grid.GetDisplayedColumnTitle(col)
            except Exception:
                info["title"] = col
            all_column_info.append(info)

        # Apply column filter
        if col_filter:
            col_set = set(col_filter)
            columns = [c for c in all_columns if c in col_set]
            column_info = [ci for ci in all_column_info if ci["name"] in col_set]
        else:
            columns = all_columns
            column_info = all_column_info

        data = []
        start_row = max(0, start_row)
        if not columns_only:
            end_row = min(start_row + max_rows, grid.RowCount)
            for row in range(start_row, end_row):
                row_data = {}
                for col in columns:
                    try:
                        row_data[col] = grid.GetCellValue(row, col)
                    except Exception:
                        row_data[col] = None
                row_data["_absolute_row_index"] = row
                data.append(row_data)

        result = {
            "table_id": table_id,
            "table_type": "GuiGridView",
            "total_rows": grid.RowCount,
            "start_row": start_row,
            "rows_returned": len(data),
            "columns": columns,
            "column_info": column_info,
            "data": data,
        }
        if columns_only:
            result["columns_only"] = True
        return result

    def _read_table_control(self, table, table_id: str, max_rows: int,
                            col_filter: List[str] = None,
                            columns_only: bool = False,
                            start_row: int = 0) -> Dict[str, Any]:
        """Read visible rows from a GuiTableControl.

        Reads from the current scroll position without changing it (unless
        start_row is explicitly provided).  This preserves the user's
        navigated position (e.g. after "Position..." in SPRO/SM30) and
        avoids crashing SAP GUI's COM server -- programmatic
        scrollbar.Position changes can destabilize certain table views.

        The read is capped at VisibleRowCount.  GuiTableControl.RowCount often
        includes padding rows (empty rows that fill the visible area beyond
        the actual data), so reading stops early when an all-empty row is
        encountered.
        """
        all_columns_info = self._get_table_control_columns(table)
        all_column_names = [c["name"] for c in all_columns_info]
        all_col_count = len(all_columns_info)

        # Build column filter index mapping
        if col_filter:
            col_set = set(col_filter)
            filtered_col_indices = [
                i for i, name in enumerate(all_column_names) if name in col_set
            ]
            column_names = [all_column_names[i] for i in filtered_col_indices]
            columns_info = [all_columns_info[i] for i in filtered_col_indices]
        else:
            filtered_col_indices = list(range(all_col_count))
            column_names = all_column_names
            columns_info = all_columns_info

        total_rows = table.RowCount
        visible_rows = table.VisibleRowCount

        data = []
        if total_rows > 0 and all_col_count > 0:
            scrollbar = table.VerticalScrollbar

            # Scroll to start_row if explicitly requested
            if start_row > 0:
                new_pos = max(scrollbar.Minimum, min(start_row, scrollbar.Maximum))
                try:
                    scrollbar.Position = new_pos
                except Exception:
                    pass  # best-effort scroll
                start_position = scrollbar.Position
            else:
                start_position = scrollbar.Position

            if not columns_only:
                rows_to_read = min(visible_rows, max_rows)
                for vis_idx in range(rows_to_read):
                    # Padding detection must check ALL columns, not just
                    # filtered ones — a row may have data in unfiltered
                    # columns while filtered columns are blank.
                    all_empty = True
                    if col_filter:
                        for ci in range(all_col_count):
                            try:
                                cell = table.GetCell(vis_idx, ci)
                                val = self._read_cell_value(cell)
                            except Exception:
                                val = None
                            if val is not None and val != "":
                                all_empty = False
                                break
                    row_data = {}
                    for col_idx in filtered_col_indices:
                        try:
                            cell = table.GetCell(vis_idx, col_idx)
                            value = self._read_cell_value(cell)
                        except Exception:
                            value = None
                        col_name = all_column_names[col_idx]
                        row_data[col_name] = value
                        if not col_filter:
                            if value is not None and value != "":
                                all_empty = False
                    if all_empty:
                        break
                    row_data["_absolute_row_index"] = start_position + vis_idx
                    data.append(row_data)
        else:
            start_position = 0

        result = {
            "table_id": table_id,
            "table_type": "GuiTableControl",
            "table_field_name": getattr(table, 'TableFieldName', ''),
            "total_rows": total_rows,
            "first_visible_row": start_position,
            "visible_rows": visible_rows,
            "rows_returned": len(data),
            "columns": column_names,
            "column_info": columns_info,
            "data": data,
        }
        if columns_only:
            result["columns_only"] = True
        return result

    def get_alv_toolbar(self, grid_id: str) -> Dict[str, Any]:
        """
        Get all toolbar buttons from an ALV grid.

        Args:
            grid_id: SAP GUI grid ID

        Returns:
            Dict with list of toolbar buttons (id, text, type)
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)
            button_count = grid.ToolbarButtonCount

            buttons = []
            for i in range(button_count):
                btn_id = grid.GetToolbarButtonId(i)
                btn_text = grid.GetToolbarButtonText(i)
                btn_type = grid.GetToolbarButtonType(i)

                btn_info = {
                    "index": i,
                    "id": btn_id,
                    "text": btn_text,
                    "type": _TOOLBAR_BUTTON_TYPES.get(btn_type, str(btn_type)),
                }

                # Add tooltip if available
                try:
                    btn_info["tooltip"] = grid.GetToolbarButtonTooltip(i)
                except Exception:
                    btn_info["tooltip"] = ""

                # Add enabled state if available
                try:
                    btn_info["enabled"] = bool(grid.GetToolbarButtonEnabled(i))
                except Exception:
                    btn_info["enabled"] = True

                buttons.append(btn_info)

            return {
                "grid_id": grid_id,
                "button_count": button_count,
                "buttons": buttons,
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id},
                e,
                "Could not read ALV toolbar",
            )

    def press_alv_toolbar_button(self, grid_id: str, button_id: str) -> Dict[str, Any]:
        """
        Press a toolbar button on an ALV grid.

        Automatically detects Menu/ButtonAndMenu types and uses
        PressToolbarContextButton instead of PressToolbarButton.
        For menu types, this opens the context menu -- use
        select_alv_context_menu_item() to pick an item.

        Args:
            grid_id: SAP GUI grid ID
            button_id: Toolbar button ID (from get_alv_toolbar)

        Returns:
            Dict with result status, screen info, and menu items if a menu was opened
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)

            # Try PressToolbarContextButton first (works for Menu/ButtonAndMenu)
            # Fall back to PressToolbarButton for regular buttons
            menu_opened = False
            try:
                grid.PressToolbarContextButton(button_id)
                menu_opened = True
            except Exception:
                # Not a menu button, use regular press
                grid.PressToolbarButton(button_id)

            if menu_opened:
                return {
                    "grid_id": grid_id,
                    "button_id": button_id,
                    "status": "menu_opened",
                    "hint": "Use select_alv_context_menu_item to pick an item",
                    "screen": self.get_screen_info(),
                }
            else:
                return {
                    "grid_id": grid_id,
                    "button_id": button_id,
                    "status": "pressed",
                    "screen": self.get_screen_info(),
                }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "button_id": button_id},
                e,
                "Could not press ALV toolbar button",
            )

    def select_alv_context_menu_item(
        self,
        grid_id: str,
        menu_item_id: str,
        toolbar_button_id: Optional[str] = None,
        select_by: str = "auto",
    ) -> Dict[str, Any]:
        """
        Select an item from an ALV context menu.

        If toolbar_button_id is provided, opens the context menu first and then
        immediately selects the item -- all in a single call to avoid timing issues.

        Args:
            grid_id: SAP GUI grid ID
            menu_item_id: Function code, visible text, or position descriptor
            toolbar_button_id: Optional toolbar button to open the menu first
            select_by: Selection mode:
                - auto: if menu_item_id contains spaces, use text; otherwise try ID first
                - id: function code selection (toolbar or regular context menu API)
                - text: visible text selection
                - position: position descriptor selection

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)
            select_mode = (select_by or "auto").lower()

            # Open context menu first if toolbar button specified
            if toolbar_button_id:
                grid.PressToolbarContextButton(toolbar_button_id)

            if select_mode not in {"auto", "id", "text", "position"}:
                raise ValueError(
                    "Invalid select_by value. Use one of: auto, id, text, position"
                )

            def _select_by_id() -> None:
                if toolbar_button_id:
                    grid.SelectToolbarMenuItem(menu_item_id)
                else:
                    try:
                        grid.SelectContextMenuItem(menu_item_id)
                    except Exception:
                        # If menu origin is unknown, try toolbar API as fallback.
                        grid.SelectToolbarMenuItem(menu_item_id)

            if select_mode == "text":
                grid.SelectContextMenuItemByText(menu_item_id)
            elif select_mode == "position":
                grid.SelectContextMenuItemByPosition(menu_item_id)
            elif select_mode == "id":
                _select_by_id()
            elif ' ' in menu_item_id:
                # Auto mode: visible text tends to contain spaces.
                grid.SelectContextMenuItemByText(menu_item_id)
            else:
                # Auto mode: function-code first, with text fallback.
                try:
                    _select_by_id()
                except Exception:
                    grid.SelectContextMenuItemByText(menu_item_id)

            return {
                "grid_id": grid_id,
                "menu_item_id": menu_item_id,
                "status": "selected",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "menu_item_id": menu_item_id},
                e,
                "Could not select ALV context menu item",
            )

    def select_table_row(self, table_id: str, row: int) -> Dict[str, Any]:
        """Select a row in a table/grid.

        For GuiTableControl, ``row`` is an **absolute** row index.  The
        table is scrolled if necessary to make the row visible, then
        ``GetAbsoluteRow(row).Selected`` is set.

        Per the SAP GUI Scripting API, ``GetAbsoluteRow`` uses absolute
        indexing (independent of scroll position), unlike ``Rows()``
        which resets after scrolling.  Confirmed via SAP GUI script
        recording which produces ``getAbsoluteRow(N).selected = true``.
        """
        self._require_session()

        try:
            table = self._find_element(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                # Scroll is best-effort — GetAbsoluteRow uses absolute
                # indexing and works regardless of scroll position.
                # Some popup tables throw COM errors on scroll.
                try:
                    self._scroll_table_control_to_row(table, row)
                except Exception:
                    pass
                table.GetAbsoluteRow(row).Selected = True
            else:
                table.selectedRows = str(row)

            return {
                "table_id": table_id,
                "selected_row": row,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not select table row",
            )

    def double_click_table_cell(self, table_id: str, row: int, column: str) -> Dict[str, Any]:
        """Double-click a cell in a table/grid.

        For GuiTableControl, ``row`` is an **absolute** row index.  The
        table is scrolled if necessary, the row is selected via
        ``GetAbsoluteRow``, focus is set on the target cell via
        ``GetCell`` (visible-row indexing), and F2 is sent to the
        owning window.
        """
        self._require_session()

        try:
            table = self._find_element(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                table.GetAbsoluteRow(row).Selected = True
                cell = table.GetCell(vis_row, col_idx)
                cell.SetFocus()
                # Send F2 to the window that owns this table
                wnd_id = table_id.split("/usr")[0]
                self._find_window(wnd_id).sendVKey(VKey.F2)
            else:
                table.DoubleClick(row, column)

            return {
                "table_id": table_id,
                "row": row,
                "column": column,
                "status": "double_clicked",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id, "row": row, "column": column},
                e,
                "Could not double-click table cell",
            )

    def modify_cell(self, grid_id: str, row: int, column: str, value: str) -> Dict[str, Any]:
        """
        Modify the value of a cell in an ALV grid or table control.

        Args:
            grid_id: SAP GUI grid/table ID
            row: Row index (0-based)
            column: Column name (ALV) or column name/index (table control)
            value: New cell value

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._find_element(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                cell = table.GetCell(vis_row, col_idx)
                cell.Text = value
            else:
                table.ModifyCell(row, column, value)

            return {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "value": value,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "row": row, "column": column},
                e,
                "Could not modify cell",
            )

    def set_current_cell(self, grid_id: str, row: int, column: str) -> Dict[str, Any]:
        """
        Set the current (focused) cell in an ALV grid or table control.

        Args:
            grid_id: SAP GUI grid/table ID
            row: Row index (0-based)
            column: Column name (ALV) or column name/index (table control)

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._find_element(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                vis_row = self._scroll_table_control_to_row(table, row)
                col_idx = self._resolve_table_control_column(table, column)
                cell = table.GetCell(vis_row, col_idx)
                cell.SetFocus()
            else:
                table.SetCurrentCell(row, column)

            return {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "row": row, "column": column},
                e,
                "Could not set current cell",
            )

    def get_column_info(self, grid_id: str) -> Dict[str, Any]:
        """
        Get detailed column information from an ALV grid or table control.

        Returns column names, displayed titles, tooltips for each column.

        Args:
            grid_id: SAP GUI grid/table ID

        Returns:
            Dict with column details
        """
        self._require_session()

        try:
            table = self._find_element(grid_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                columns = self._get_table_control_columns(table)
                return {
                    "grid_id": grid_id,
                    "table_type": "GuiTableControl",
                    "column_count": len(columns),
                    "columns": columns,
                }

            # ALV grid
            columns = []
            for i in range(table.ColumnCount):
                col_name = table.ColumnOrder(i)
                col_info = {"name": col_name, "index": i}

                try:
                    col_info["title"] = table.GetDisplayedColumnTitle(col_name)
                except Exception:
                    col_info["title"] = col_name

                try:
                    col_info["tooltip"] = table.GetColumnTooltip(col_name)
                except Exception:
                    col_info["tooltip"] = ""

                columns.append(col_info)

            return {
                "grid_id": grid_id,
                "column_count": len(columns),
                "columns": columns,
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id},
                e,
                "Could not read column information",
            )

    # ---- TableControl-specific operations ----

    def scroll_table_control(self, table_id: str, position: int) -> Dict[str, Any]:
        """
        Scroll a GuiTableControl to a specific row position.

        Since read_table does not scroll (it reads only visible rows),
        use this tool to navigate to a different section of the table
        before reading.

        Args:
            table_id: SAP GUI table control ID
            position: Absolute row position to scroll to

        Returns:
            Dict with new scroll position and visible data summary
        """
        self._require_session()

        try:
            table = self._find_element(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {
                    "table_id": table_id,
                    "error": "Not a GuiTableControl. Use ALV grid methods.",
                }

            scrollbar = table.VerticalScrollbar
            scroll_max = scrollbar.Maximum
            new_pos = max(scrollbar.Minimum, min(position, scroll_max))
            try:
                scrollbar.Position = new_pos
            except Exception as scroll_err:
                # Some popup tables throw COM errors on scroll.
                # Return diagnostic info so the caller can use
                # navigation buttons or Position... dialog instead.
                return {
                    "table_id": table_id,
                    "error": self._sanitize_error_message(
                        scroll_err,
                        "Could not scroll table control",
                    ),
                    "requested_position": position,
                    "clamped_position": new_pos,
                    "scroll_max": scroll_max,
                    "total_rows": table.RowCount,
                    "hint": "Scrollbar failed; use navigation buttons or Position dialog.",
                }

            return {
                "table_id": table_id,
                "status": "success",
                "position": new_pos,
                "visible_rows": table.VisibleRowCount,
                "total_rows": table.RowCount,
                "scroll_max": scroll_max,
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not scroll table control",
            )

    def get_table_control_row_info(self, table_id: str,
                                    rows: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Get row metadata from a GuiTableControl.

        Returns whether each row is selectable and currently selected,
        using GetAbsoluteRow() for absolute indexing.

        Args:
            table_id: SAP GUI table control ID
            rows: List of absolute row indices to query.
                  If None, queries all currently visible rows.

        Returns:
            Dict with row info list
        """
        self._require_session()

        try:
            table = self._find_element(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {"table_id": table_id, "error": "Not a GuiTableControl"}

            if rows is None:
                start = table.VerticalScrollbar.Position
                rows = list(range(start, start + table.VisibleRowCount))

            row_info = []
            for r in rows:
                info: Dict[str, Any] = {"row": r}
                try:
                    abs_row = table.GetAbsoluteRow(r)
                    info["selectable"] = getattr(abs_row, 'Selectable', True)
                    info["selected"] = getattr(abs_row, 'Selected', False)
                except Exception as e:
                    info["error"] = self._sanitize_error_message(
                        e,
                        "Could not inspect table row",
                    )
                row_info.append(info)

            return {
                "table_id": table_id,
                "row_count": len(row_info),
                "rows": row_info,
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not read table row info",
            )

    def select_all_table_control_columns(self, table_id: str,
                                          select: bool = True) -> Dict[str, Any]:
        """
        Select or deselect all columns in a GuiTableControl.

        Args:
            table_id: SAP GUI table control ID
            select: True to select all, False to deselect all

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            table = self._find_element(table_id)
            if getattr(table, 'Type', '') != "GuiTableControl":
                return {"table_id": table_id, "error": "Not a GuiTableControl"}

            if select:
                table.SelectAllColumns()
            else:
                table.DeselectAllColumns()

            return {
                "table_id": table_id,
                "status": "all_selected" if select else "all_deselected",
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not change table column selection",
            )

    # ---- ALV-specific operations ----

    def get_cell_info(self, grid_id: str, row: int,
                      column: str) -> Dict[str, Any]:
        """
        Get detailed cell metadata from an ALV grid (GuiGridView).

        Returns whether the cell is editable, its color/style, and tooltip.

        Args:
            grid_id: SAP GUI grid ID (ALV)
            row: Row index (0-based)
            column: Column name

        Returns:
            Dict with cell properties
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)
            info: Dict[str, Any] = {
                "grid_id": grid_id,
                "row": row,
                "column": column,
                "value": grid.GetCellValue(row, column),
            }

            for method, key in [
                ("GetCellChangeable", "changeable"),
                ("GetCellColor", "color"),
                ("GetCellTooltip", "tooltip"),
                ("GetCellStyle", "style"),
                ("GetCellMaxLength", "max_length"),
            ]:
                try:
                    info[key] = getattr(grid, method)(row, column)
                except Exception:
                    pass

            return info
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "row": row, "column": column},
                e,
                "Could not read cell information",
            )

    def press_column_header(self, grid_id: str,
                             column: str) -> Dict[str, Any]:
        """
        Click a column header in an ALV grid (triggers sort/filter).

        Args:
            grid_id: SAP GUI grid ID (ALV)
            column: Column name

        Returns:
            Dict with result status and screen info
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)
            grid.PressColumnHeader(column)
            return {
                "grid_id": grid_id,
                "column": column,
                "status": "pressed",
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id, "column": column},
                e,
                "Could not press column header",
            )

    def select_all_rows(self, grid_id: str) -> Dict[str, Any]:
        """
        Select all rows in an ALV grid.

        Args:
            grid_id: SAP GUI grid ID (ALV)

        Returns:
            Dict with result status
        """
        self._require_session()

        try:
            grid = self._find_element(grid_id)
            grid.SelectAll()
            return {"grid_id": grid_id, "status": "all_selected"}
        except Exception as e:
            return self._error_result(
                {"grid_id": grid_id},
                e,
                "Could not select all rows",
            )

    # ---- Operations for both table types ----

    def get_current_cell(self, table_id: str) -> Dict[str, Any]:
        """
        Get the currently focused cell in an ALV grid or table control.

        Args:
            table_id: SAP GUI grid/table ID

        Returns:
            Dict with current row and column
        """
        self._require_session()

        try:
            table = self._find_element(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                return {
                    "table_id": table_id,
                    "table_type": "GuiTableControl",
                    "current_row": getattr(table, 'CurrentRow', -1),
                    "current_col": getattr(table, 'CurrentCol', -1),
                }
            else:
                return {
                    "table_id": table_id,
                    "table_type": "GuiGridView",
                    "current_row": getattr(table, 'CurrentCellRow', -1),
                    "current_column": getattr(table, 'CurrentCellColumn', ''),
                }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not read current cell",
            )

    # =========================================================================
    # Multi-Row Selection
    # =========================================================================

    def select_multiple_rows(
        self, table_id: str, rows: List[int],
    ) -> Dict[str, Any]:
        """
        Select multiple rows in an ALV grid or table control.

        For ALV grids, uses the SelectedRows property which accepts
        comma-separated row indices (e.g., "0,2,5").

        For GuiTableControl, iterates and sets Selected on each
        absolute row.

        Args:
            table_id: SAP GUI table/grid ID
            rows: List of row indices to select

        Returns:
            Dict with selection result
        """
        self._require_session()

        try:
            table = self._find_element(table_id)

            if getattr(table, 'Type', '') == "GuiTableControl":
                for row in rows:
                    self._scroll_table_control_to_row(table, row)
                    table.GetAbsoluteRow(row).Selected = True
            else:
                # ALV: comma-separated row list
                table.SelectedRows = ",".join(str(r) for r in rows)

            return {
                "table_id": table_id,
                "selected_rows": rows,
                "count": len(rows),
                "status": "success",
            }
        except Exception as e:
            return self._error_result(
                {"table_id": table_id},
                e,
                "Could not select multiple rows",
            )
