"""
Discovery Mixin - popups, toolbars, shell content, screen elements, screenshots.

Provides screen discovery, popup handling, toolbar enumeration, and
screenshot capabilities for the SAP GUI controller.
"""

import logging
from typing import Any, Dict, List

from .models import ScreenElement

logger = logging.getLogger(__name__)


class DiscoveryMixin:
    """Mixin for discovery and inspection operations on SAP GUI screens."""

    # =========================================================================
    # Popup Window Handling
    # =========================================================================

    def get_popup_window(self) -> Dict[str, Any]:
        """
        Check if a popup window (modal dialog) is currently open.

        SAP popups appear as wnd[1], wnd[2], etc.  Returns the topmost
        popup's title, text content, and available buttons so the AI can
        decide how to respond.

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
        try:
            usr = self._session.findById(f"{popup_id}/usr")
            self._collect_popup_contents(usr, texts, buttons)
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
                            "id": btn.Id,
                            "text": getattr(btn, 'Text', '').strip(),
                            "tooltip": getattr(btn, 'Tooltip', '').strip(),
                        })
            except Exception:
                pass

        if texts:
            result["texts"] = texts
        if buttons:
            result["buttons"] = buttons

        return result

    def _collect_popup_contents(
        self, container, texts: list, buttons: list, depth: int = 0,
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
                        "id": child.Id,
                        "text": text,
                        "tooltip": getattr(child, 'Tooltip', '').strip(),
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
                                child, texts, buttons, depth + 1,
                            )
                    except Exception:
                        pass
        except Exception:
            pass

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

        toolbars: Dict[str, list] = {}
        for tbar_idx, tbar_name in [(0, "system_toolbar"), (1, "application_toolbar")]:
            buttons = []
            try:
                tbar = self._session.findById(f"{window_id}/tbar[{tbar_idx}]")
                for i in range(tbar.Children.Count):
                    btn = tbar.Children(i)
                    btype = getattr(btn, 'Type', '')
                    if btype in ('GuiButton',):
                        buttons.append({
                            "id": btn.Id,
                            "text": getattr(btn, 'Text', '').strip(),
                            "tooltip": getattr(btn, 'Tooltip', '').strip(),
                            "enabled": getattr(btn, 'Changeable', True) is not False,
                        })
            except Exception:
                pass
            if buttons:
                toolbars[tbar_name] = buttons

        return {
            "window_id": window_id,
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
            shell = self._session.findById(shell_id)
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
            return {"shell_id": shell_id, "error": str(e)}

    # =========================================================================
    # Screen Element Discovery
    # =========================================================================

    def get_screen_elements(self, container_id: str = "wnd[0]/usr",
                            max_depth: int = 3,
                            type_filter: str = "",
                            changeable_only: bool = False) -> List[ScreenElement]:
        """
        Enumerate all elements on the current screen.

        Useful for discovering field IDs when automating a new transaction.

        Args:
            container_id: Starting container (default: main user area)
            max_depth: Maximum recursion depth
            type_filter: Comma-separated SAP element types to include
                (e.g. "GuiTextField,GuiCTextField"). Empty = all types.
            changeable_only: If True, only return editable/input elements

        Returns:
            List of ScreenElement objects
        """
        self._require_session()

        type_filter_set = None
        if type_filter:
            type_filter_set = {t.strip() for t in type_filter.split(",") if t.strip()}

        try:
            container = self._session.findById(container_id)
            elements = self._enumerate_elements(
                container, max_depth,
                type_filter_set=type_filter_set,
                changeable_only=changeable_only,
            )
            return elements
        except Exception as e:
            logger.error(f"Failed to enumerate elements: {e}")
            return []

    def _enumerate_elements(self, container, max_depth: int,
                            current_depth: int = 0,
                            type_filter_set: set = None,
                            changeable_only: bool = False) -> List[ScreenElement]:
        """Recursively enumerate screen elements."""
        elements = []

        if current_depth >= max_depth:
            return elements

        try:
            for i in range(container.Children.Count):
                child = container.Children(i)

                element = ScreenElement(
                    id=child.Id,
                    type=child.Type,
                    name=getattr(child, 'Name', ''),
                    text=str(getattr(child, 'Text', ''))[:200],
                    changeable=getattr(child, 'Changeable', False),
                    visible=getattr(child, 'Visible', True),
                )

                # Apply filters — but always recurse into containers
                include = True
                if type_filter_set and element.type not in type_filter_set:
                    include = False
                if changeable_only and not element.changeable:
                    include = False
                if include:
                    elements.append(element)

                # Recurse into containers regardless of filters
                if hasattr(child, 'Children') and child.Children.Count > 0:
                    child_elements = self._enumerate_elements(
                        child, max_depth, current_depth + 1,
                        type_filter_set=type_filter_set,
                        changeable_only=changeable_only,
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
                return active.Id
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

        try:
            if filepath is None:
                filepath = os.path.join(tempfile.gettempdir(), "sap_screenshot.png")
                return_base64 = True
            else:
                return_base64 = False

            # Find the topmost window (popups are wnd[1], wnd[2], etc.)
            window_id = self._find_topmost_window()
            window = self._session.findById(window_id)
            window.HardCopy(filepath, "PNG")

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
            return {"error": str(e)}

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
