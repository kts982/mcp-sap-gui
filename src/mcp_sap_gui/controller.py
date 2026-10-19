"""
SAP GUI Controller Base - connection, navigation, and screen info.

This module provides the core controller class with connection management,
transaction execution, and screen information retrieval.
"""

import json
import logging
import os
import re
import time
from datetime import datetime
from functools import lru_cache
from typing import Any, Dict, FrozenSet, Iterator, List, Optional, Tuple

from .models import (
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    SessionInfo,
    VKey,
    _strip_tcode_prefix,
)

logger = logging.getLogger(__name__)
# GetObjectTree calls and their failures, on a logger of their own so that
# they can be raised to DEBUG (or silenced) without the rest of the package.
_tree_logger = logging.getLogger("mcp_sap_gui.object_tree")


_WINDOW_ID_RE = re.compile(r"^wnd\[(\d+)\]$")
_NORMALIZED_WND_PATH_RE = re.compile(r"(?:^|/)(wnd\[\d+\].*)$")
# SAP GUI 8.x runs sessions in sapguiserver.exe processes. Each registers a
# process-specific ROT entry "SAPGUISERVER-<hexPID>" (SAP Business Client,
# ABAP in Eclipse, direct sapgui.exe) plus an unsuffixed "SAPGUISERVER" that
# only ever resolves to the first process. SAP Logon registers "SAPGUI".
_SAPGUISERVER_ROT_RE = re.compile(r"^SAPGUISERVER-[0-9A-F]+$", re.IGNORECASE)

_NOT_AVAILABLE_MESSAGE = (
    "Cannot connect to SAP GUI. Ensure SAP Logon Pad, SAP Business Client "
    "or another SAP GUI session is running"
)
_SCRIPTING_DISABLED_NOTE = (
    "SAP GUI Scripting is disabled on this server (profile parameter "
    "sapgui/user_scripting); no session can be read."
)
# RPC_E_DISCONNECTED: the session window/tab was closed under us.
_HRESULT_DISCONNECTED = -2147417848
# RPC_S_SERVER_UNAVAILABLE: the whole GUI server process is gone.
_HRESULT_SERVER_UNAVAILABLE = -2147023174
_DEAD_SESSION_HRESULTS = frozenset({_HRESULT_DISCONNECTED, _HRESULT_SERVER_UNAVAILABLE})

# Top-level window children: user area, menu/status/title bars, toolbars and
# docking containers. The latter (wnd[0]/shellcont, wnd[0]/shellcont[1]) sit
# BESIDE usr and hold the dialog-structure tree of view clusters (SM34, most
# IMG activities) and the SE80 repository tree.
# Below them, a path segment is a type prefix plus the element's screen name,
# and a screen name is not just letters and digits: %#AUTOTEXT001 (unnamed
# elements, e.g. SO23's btn%#AUTOTEXT005), *DY_1000-CLIF (SE24),
# DYNPROLOGIC=>DYNPFIELDS-SELECTED_CHECKS (SLIN), <FILTER>-FIELD_LABEL, and
# + & ' ! ? " ( ) { in other standard screens. So a segment is anything but
# whitespace and the separators / [ ]; discovery returns the IDs as they are.
_ELEMENT_ID_RE = re.compile(
    r"^wnd\[(\d+)\]"
    r"(?:/(?:usr|mbar|sbar|titl|tbar\[\d+\]|shellcont(?:\[\d+\])?)"
    r"(?:/[^\s/\[\]]+(?:\[[A-Za-z0-9_,]+\])*)*)$"
)

# The callers of GuiSession.GetObjectTree, as named in
# MCP_SAP_GUI_DISABLE_OBJECT_TREE: table control cells (sap_read_table),
# popup windows (sap_get_popup_window and the popup digest of every action
# response), classic list screens (sap_read_list), sap_get_screen_elements
# and the sub-screen IDs in the screen info of every action response.
OBJECT_TREE_SITES = ("table", "popup", "list", "elements", "subscreens")
OBJECT_TREE_DISABLE_ENV = "MCP_SAP_GUI_DISABLE_OBJECT_TREE"
OBJECT_TREE_TRACE_ENV = "MCP_SAP_GUI_TRACE_FILE"
_ALL_SITES_VALUES = frozenset({"1", "true", "yes", "on", "all"})
# GetObjectTree is missing (SAP GUI before 7.70 PL3): the walk is by design.
_DISP_E_UNKNOWNNAME = -2147352570
_DISP_E_MEMBERNOTFOUND = -2147352573
_NOT_SUPPORTED_HRESULTS = frozenset({_DISP_E_UNKNOWNNAME, _DISP_E_MEMBERNOTFOUND})
# Never asked of GetObjectTree. On SAP GUI 8.00 (PL14 and PL16 seen) a call
# that requests Selected fails, and SAP Logon often dies with every session
# in it (issue #22); every other property read fine on the same screens.
# Read Selected over COM, on the checkboxes that have it.
_OBJECT_TREE_UNSAFE_PROPS = frozenset({"Selected"})


@lru_cache(maxsize=8)
def _parse_disabled_sites(value: str) -> FrozenSet[str]:
    value = value.strip().lower()
    if not value:
        return frozenset()
    if value in _ALL_SITES_VALUES:
        return frozenset(OBJECT_TREE_SITES)
    sites = {part.strip() for part in value.split(",") if part.strip()}
    unknown = sorted(sites - set(OBJECT_TREE_SITES))
    if unknown:
        raise ValueError(
            f"{OBJECT_TREE_DISABLE_ENV}: unknown site(s) {', '.join(unknown)}; "
            f"use 'all' or a comma list of {', '.join(OBJECT_TREE_SITES)}"
        )
    return frozenset(sites)


def object_tree_disabled_sites(value: Optional[str] = None) -> FrozenSet[str]:
    """The GetObjectTree call sites turned off by MCP_SAP_GUI_DISABLE_OBJECT_TREE.

    Unset or empty: none. ``all`` (or ``1``/``true``/``yes``/``on``): every
    site. Otherwise a comma list of OBJECT_TREE_SITES, e.g. ``table,popup``.
    A disabled site reads its elements one by one over COM, as on SAP GUI
    before 7.70 PL3. Raises ValueError for an unknown site name; the server
    checks this at start-up and refuses to run.
    """
    if value is None:
        value = os.environ.get(OBJECT_TREE_DISABLE_ENV, "")
    return _parse_disabled_sites(value)


def _hresult(exc: BaseException) -> Optional[int]:
    """The HRESULT of a pywin32 com_error, else None."""
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        return hresult
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def _trace_object_tree(phase: str, site: str, root_id: str, detail: str = "") -> None:
    """Append one line to MCP_SAP_GUI_TRACE_FILE and fsync it.

    The ``start`` line is on disk before the COM call is made. If SAP GUI
    dies inside the call, the last ``start`` with no ``done`` or ``failed``
    after it names the call that was in flight. Timestamps are local time
    with the UTC offset, to line up with the Windows event log.
    """
    path = os.environ.get(OBJECT_TREE_TRACE_ENV, "").strip()
    if not path:
        return
    stamp = datetime.now().astimezone().isoformat(timespec="milliseconds")
    line = f"{stamp} {os.getpid()} {phase} {site} {root_id}"
    if detail:
        line += f" {detail}"
    try:
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
            handle.flush()
            os.fsync(handle.fileno())
    except OSError as e:
        _tree_logger.debug("Cannot write %s %s: %s", OBJECT_TREE_TRACE_ENV, path, e)


# A sub-screen container: sub<area>:<program>:<dynnr>, ssub inside a tab strip.
# A namespaced program brings its slashes along: subX:/SCWM/SAPLRF_SSCR:0001.
_SUBSCREEN_ID_RE = re.compile(
    r"/s?sub[A-Za-z0-9_]*:((?:/[A-Za-z0-9_]+/)?[A-Za-z0-9_]+):(\d{4})$"
)


class SAPGUIControllerBase:
    """
    Base controller for SAP GUI Scripting API via COM automation.

    Provides connection management, transaction navigation, and screen
    information retrieval.  Extended by mixins for fields, tables, trees,
    and discovery operations.
    """

    def __init__(self):
        """Initialize the SAP GUI controller."""
        self._win32com = None
        self._sap_gui_auto = None
        self._saplogon_engine = None
        # Scripting engine the bound connection came from (SAP Logon or a
        # SAP GUI server process); session-lookup helpers walk this one.
        self._application = None
        self._connection = None
        self._session = None
        self._owns_session = False
        # Read once per bind (_read_session_traits): the server's write gate
        # checks scripting_read_only on the event loop, where COM is off limits.
        self.sap_gui_version = ""
        self.scripting_read_only = False
        self._check_dependencies()

    def _check_dependencies(self):
        """Check if required dependencies are available."""
        try:
            import win32com.client
            self._win32com = win32com.client
        except ImportError:
            raise SAPGUINotAvailableError(
                "pywin32 is required but not installed. "
                "Install with: pip install pywin32"
            )

    @property
    def is_connected(self) -> bool:
        """Check if connected to an SAP system."""
        return self._session is not None

    def _ensure_com_initialized(self):
        """Ensure COM is initialized on the current thread (needed for thread pool workers)."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass  # Already initialized or not needed

    def _get_sap_gui(self):
        """Get the SAP Logon (``SAPGUI``) automation object."""
        if self._sap_gui_auto is None:
            self._ensure_com_initialized()
            try:
                self._sap_gui_auto = self._win32com.GetObject("SAPGUI")
            except Exception as e:
                logger.warning(
                    "Failed to get SAP GUI automation object: %s",
                    e,
                    exc_info=logger.isEnabledFor(logging.DEBUG),
                )
                raise SAPGUINotAvailableError(_NOT_AVAILABLE_MESSAGE + ".")
        return self._sap_gui_auto

    def _get_saplogon_engine(self):
        """Get the SAP Logon scripting engine (needed for OpenConnection)."""
        if self._saplogon_engine is None:
            sap_gui = self._get_sap_gui()
            # Use property access (no parentheses) - works more reliably
            engine = sap_gui.GetScriptingEngine
            if engine is None:
                raise SAPGUINotAvailableError(
                    "Could not get SAP GUI Scripting Engine. "
                    "Is SAP GUI Scripting enabled?"
                )
            self._saplogon_engine = engine
        return self._saplogon_engine

    def _get_application(self):
        """Get the scripting engine of the bound session.

        Falls back to the SAP Logon engine when nothing is bound yet.
        """
        if self._application is None:
            self._application = self._get_saplogon_engine()
        return self._application

    def _list_rot_names(self) -> List[str]:
        """Return the display names of all COM Running Object Table entries.

        Other applications register there too: one whose name cannot be read
        is skipped, so it cannot hide the SAP entries.
        """
        import pythoncom

        rot = pythoncom.GetRunningObjectTable()
        bind_ctx = pythoncom.CreateBindCtx(0)
        names = []
        for moniker in rot.EnumRunning():
            try:
                names.append(str(moniker.GetDisplayName(bind_ctx, None)))
            except Exception as e:
                logger.debug("Skipping unreadable ROT entry: %s", e)
        return names

    def _get_engine_from_rot(self, rot_name: str):
        """Resolve a ROT entry to its scripting engine, or None on failure."""
        try:
            gui = self._win32com.GetObject(rot_name)
            # Property, not a method: calling it raises "Member not found".
            engine = gui.GetScriptingEngine
        except Exception as e:
            logger.debug(
                "Skipping scripting engine %s: %s",
                rot_name,
                e,
                exc_info=logger.isEnabledFor(logging.DEBUG),
            )
            return None
        if engine is None:
            logger.debug("ROT entry %s returned no scripting engine", rot_name)
        return engine

    def _iter_scripting_engines(self) -> Iterator[Tuple[str, str, Any]]:
        """Yield ``(host, rot_name, engine)`` for every reachable SAP GUI.

        Order: SAP Logon (``SAPGUI``) first, then every process-specific
        ``SAPGUISERVER-<hexPID>`` entry sorted by name. The unsuffixed
        ``SAPGUISERVER`` is used only when no suffixed entry exists, since
        it always points at the first GUI server process. Failures on single
        entries are skipped; if the ROT cannot be enumerated, only SAP Logon
        is tried.
        """
        self._ensure_com_initialized()

        engine = self._get_engine_from_rot("SAPGUI")
        if engine is not None:
            yield "saplogon", "SAPGUI", engine

        try:
            names = self._list_rot_names()
        except Exception as e:
            logger.debug(
                "ROT enumeration failed, using SAP Logon only: %s",
                e,
                exc_info=logger.isEnabledFor(logging.DEBUG),
            )
            return

        server_names = sorted({n for n in names if _SAPGUISERVER_ROT_RE.fullmatch(n)})
        if not server_names and "SAPGUISERVER" in names:
            server_names = ["SAPGUISERVER"]

        for rot_name in server_names:
            engine = self._get_engine_from_rot(rot_name)
            if engine is not None:
                yield "sapguiserver", rot_name, engine

    def _discover_connections(self) -> List[Tuple[str, str, Any, Any]]:
        """Return ``(host, rot_name, engine, connection)`` across all engines.

        The list position is the global connection index used by
        list_connections() and connect_to_existing_session().

        Raises:
            SAPGUINotAvailableError: If no scripting engine is reachable.
        """
        engines = list(self._iter_scripting_engines())
        if not engines:
            raise SAPGUINotAvailableError(_NOT_AVAILABLE_MESSAGE + ".")

        found: List[Tuple[str, str, Any, Any]] = []
        for host, rot_name, engine in engines:
            try:
                for i in range(engine.Children.Count):
                    found.append((host, rot_name, engine, engine.Children(i)))
            except Exception as e:
                logger.warning(
                    "Could not list connections of %s: %s",
                    rot_name,
                    e,
                    exc_info=logger.isEnabledFor(logging.DEBUG),
                )
        return found

    @staticmethod
    def _is_scripting_disabled(conn) -> bool:
        """Return True when the server disabled scripting for this connection."""
        try:
            value = conn.DisabledByServer
        except Exception:
            return False
        return isinstance(value, (bool, int)) and bool(value)

    def _normalize_window_id(self, window_id: Any) -> str:
        """Normalize SAP window IDs to the short form used by findById()."""
        if not isinstance(window_id, str):
            raise ValueError("SAP window ID must be a string like 'wnd[0]'")
        window_id = window_id.strip()
        match = _NORMALIZED_WND_PATH_RE.search(window_id)
        if match:
            return match.group(1)
        return window_id

    def _normalize_element_id(self, element_id: Any) -> str:
        """Normalize SAP element IDs to the short form used by findById()."""
        if not isinstance(element_id, str):
            raise ValueError("SAP element ID must be a string like 'wnd[0]/usr/...'")
        element_id = element_id.strip()
        match = _NORMALIZED_WND_PATH_RE.search(element_id)
        if match:
            return match.group(1)
        return element_id

    def _validate_window_id(self, window_id: Any) -> str:
        """Validate a user-supplied SAP window ID."""
        normalized = self._normalize_window_id(window_id)
        if not _WINDOW_ID_RE.fullmatch(normalized):
            raise ValueError(
                f"Invalid SAP window ID: {window_id!r}. Expected format like 'wnd[0]'."
            )
        return normalized

    def _validate_element_id(self, element_id: Any) -> str:
        """Validate a user-supplied SAP element/container ID."""
        normalized = self._normalize_element_id(element_id)
        if not _ELEMENT_ID_RE.fullmatch(normalized):
            raise ValueError(
                "Invalid SAP element ID: "
                f"{element_id!r}. Expected format like 'wnd[0]/usr/...', "
                "'wnd[0]/mbar/...', 'wnd[0]/sbar', 'wnd[0]/tbar[0]/...', "
                "or 'wnd[0]/shellcont/...' (docking container)."
            )
        return normalized

    def _validate_container_id(self, container_id: Any) -> str:
        """Validate a discovery container: an element ID or a bare window.

        A bare ``wnd[n]`` is accepted here only, so the window's top-level
        children (incl. docking containers) can be listed; every other tool
        keeps the stricter element grammar.
        """
        normalized = self._normalize_element_id(container_id)
        if _WINDOW_ID_RE.fullmatch(normalized):
            return normalized
        return self._validate_element_id(container_id)

    def _find_window(self, window_id: Any):
        """Validate and resolve an SAP window by short ID."""
        normalized = self._validate_window_id(window_id)
        return self._session.findById(normalized)

    def _find_element(self, element_id: Any):
        """Validate and resolve an SAP element by short ID."""
        normalized = self._validate_element_id(element_id)
        return self._session.findById(normalized)

    def _object_tree(self, root_id: str, props: List[str], *,
                     site: str) -> Optional[Dict[str, Any]]:
        """An element and all its descendants in ONE call, or None.

        GuiSession.GetObjectTree (SAP GUI 7.70 PL3+) returns the subtree as
        JSON nodes ``{"properties": {...}, "children": [...]}`` holding the
        requested properties, every value a string ("" where the element
        lacks the property). Reading each element over COM instead costs one
        round trip per property: 4.7 s against 0.2 s for a 1,139-cell list.
        None on an older SAP GUI or any failure, so callers walk instead.

        *props* must not include Selected (_OBJECT_TREE_UNSAFE_PROPS).
        *site* names the caller (one of OBJECT_TREE_SITES). A site listed in
        MCP_SAP_GUI_DISABLE_OBJECT_TREE returns None without making the call,
        and MCP_SAP_GUI_TRACE_FILE records each call before it is made.
        """
        if site not in OBJECT_TREE_SITES:
            raise ValueError(f"unknown GetObjectTree site {site!r}")
        unsafe = _OBJECT_TREE_UNSAFE_PROPS.intersection(props)
        if unsafe:
            raise ValueError(
                f"GetObjectTree must not read {', '.join(sorted(unsafe))}: "
                "it crashes SAP GUI 8.00 (issue #22)"
            )
        if site in object_tree_disabled_sites():
            return None
        _tree_logger.debug("GetObjectTree start site=%s root=%s props=%d",
                           site, root_id, len(props))
        _trace_object_tree("start", site, root_id, f"props={len(props)}")
        started = time.perf_counter()
        try:
            raw = self._session.GetObjectTree(root_id, props)
            ms = (time.perf_counter() - started) * 1000
            size = len(raw) if isinstance(raw, str) else -1
            _trace_object_tree("done", site, root_id, f"ms={ms:.0f} bytes={size}")
            _tree_logger.debug("GetObjectTree done site=%s root=%s ms=%.0f bytes=%d",
                               site, root_id, ms, size)
            if not isinstance(raw, str):
                return None
            nodes = json.loads(raw).get("children") or []
            return nodes[0] if nodes else None
        except Exception as e:
            ms = (time.perf_counter() - started) * 1000
            hresult = _hresult(e)
            _trace_object_tree("failed", site, root_id,
                               f"ms={ms:.0f} error={type(e).__name__} hresult={hresult}")
            if isinstance(e, AttributeError) or hresult in _NOT_SUPPORTED_HRESULTS:
                _tree_logger.debug("GetObjectTree(%s) not available: %s", root_id, e)
            else:
                _tree_logger.warning(
                    "GetObjectTree failed site=%s root=%s props=%d hresult=%s "
                    "error=%s: %s; reading element by element instead",
                    site, root_id, len(props), hresult, type(e).__name__, e,
                )
            return None

    # The elements whose Selected state the read tools report.
    _SELECTABLE_TYPES = ("GuiRadioButton", "GuiCheckBox")

    def _selected(self, element_id: str = "", element=None) -> Optional[bool]:
        """Selected of one radio button or checkbox over COM, or None.

        GetObjectTree must not be asked for it (_OBJECT_TREE_UNSAFE_PROPS),
        so a tree read looks the element up by *element_id*; a walk passes
        the *element* it holds. None when it cannot be read.
        """
        try:
            if element is None:
                element = self._session.findById(self._normalize_element_id(element_id))
            return bool(element.Selected)
        except Exception:
            return None

    def _is_sensitive_field_id(self, field_id: str) -> bool:
        """Return True when a field ID likely refers to a secret input."""
        normalized = (field_id or "").upper()
        return any(token in normalized for token in ("PWD", "BCODE", "PASSWORD"))

    def _mask_field_value(self, field_id: str, value: Any) -> str:
        """Mask sensitive field values before logging them."""
        if self._is_sensitive_field_id(field_id):
            return "***"
        return str(value)

    def _sanitize_error_message(self, exc: Exception, fallback: str) -> str:
        """Return a client-safe error message without raw COM details."""
        message = str(exc).strip()
        if isinstance(exc, SAPGUINotAvailableError):
            return _NOT_AVAILABLE_MESSAGE + " and SAP GUI Scripting is enabled."
        if isinstance(exc, (SAPGUINotConnectedError, SAPGUIError, ValueError)):
            return message or fallback
        if message:
            lower = message.lower()
            blocked_tokens = (
                "host=",
                "server=",
                "path=",
                "traceback",
                "clsid",
                "progid",
                "c:\\",
                "\\",
                "/app/",
                "/con[",
                "/ses[",
                ".dll",
                ".exe",
                "0x",
            )
            if not any(token in lower for token in blocked_tokens):
                if re.fullmatch(r"[A-Za-z0-9 _.,'()/-]{1,120}", message):
                    return message
        return fallback

    def _error_result(
        self, context: Dict[str, Any], exc: Exception, fallback: str,
    ) -> Dict[str, Any]:
        """Log full error details server-side and return a safe client payload."""
        logger.warning(
            "%s: %s",
            fallback,
            exc,
            exc_info=logger.isEnabledFor(logging.DEBUG),
        )
        result = dict(context)
        result["error"] = self._sanitize_error_message(exc, fallback)
        return result

    def _require_session(self):
        """Ensure we have an active session that is not busy."""
        if not self.is_connected:
            raise SAPGUINotConnectedError(
                "Not connected to SAP. Call connect() first."
            )
        try:
            if self._session.Busy:
                raise SAPGUIError(
                    "SAP session is busy processing a previous request. "
                    "Wait for it to complete before sending another command."
                )
        except AttributeError:
            # Busy unavailable: either an older GUI version, or the session
            # window/tab was closed (dynamic dispatch then loses type info
            # and Busy raises AttributeError). Probe liveness via Id.
            self._check_session_alive()

    def _check_session_alive(self) -> None:
        """Raise SAPGUINotConnectedError if the bound session was closed.

        Only the "disconnected" and "server unavailable" HRESULTs count as a
        dead session (pywintypes.com_error carries them as ``hresult``); any
        other failure keeps the binding so a transient error never forces a
        reconnect.
        """
        try:
            self._session.Id
        except Exception as e:
            if getattr(e, "hresult", None) not in _DEAD_SESSION_HRESULTS:
                return  # Id also unavailable, but not a dead session
            logger.info("Bound SAP session was closed: %s", e)
            self._session = None
            self._connection = None
            raise SAPGUINotConnectedError(
                "SAP session was closed (window or tab closed). "
                "Reconnect with sap_connect_existing."
            )

    # =========================================================================
    # Connection Management
    # =========================================================================

    def connect(self, system_description: str,
                client: str = None,
                user: str = None,
                password: str = None,
                language: str = None) -> SessionInfo:
        """
        Connect to an SAP system.

        Opens a connection exactly like double-clicking in SAP Logon Pad.
        Optionally fills login credentials.

        Args:
            system_description: Exact system name as shown in SAP Logon
            client: SAP client number (optional)
            user: SAP username (optional)
            password: SAP password (optional)
            language: Login language (optional, e.g., "EN")

        Returns:
            SessionInfo with connection details

        Raises:
            SAPGUIError: If connection fails
        """
        try:
            # OpenConnection only exists on the SAP Logon engine.
            app = self._get_saplogon_engine()

            logger.info("Opening connection to: %s", system_description)
            self._connection = app.OpenConnection(system_description, True)
            self._application = app
            self._owns_session = True

            if self._connection is None:
                raise SAPGUIError(f"Failed to open connection to '{system_description}'")

            self._session = self._connection.Children(0)

            if self._session is None:
                raise SAPGUIError("No session available on the connection")

            # Fill login credentials if provided
            if client:
                self._safe_set_field("wnd[0]/usr/txtRSYST-MANDT", str(client))
            if user:
                self._safe_set_field("wnd[0]/usr/txtRSYST-BNAME", user)
            if password:
                self._safe_set_field("wnd[0]/usr/pwdRSYST-BCODE", password)
            if language:
                self._safe_set_field("wnd[0]/usr/txtRSYST-LANGU", language)

            # Press Enter to login if credentials were provided
            has_credentials = password is not None
            if has_credentials:
                self.send_vkey(VKey.ENTER)

            logger.info("Connected successfully to %s as %s",
                         system_description, user or "(existing credentials)")
            self._read_session_traits()
            return self.get_session_info()

        except SAPGUIError:
            raise
        except Exception as e:
            logger.warning(
                "Connection failed for '%s': %s",
                system_description,
                e,
                exc_info=logger.isEnabledFor(logging.DEBUG),
            )
            raise SAPGUIError(
                "Connection failed. Verify the SAP Logon entry and login screen state."
            )

    def connect_to_existing_session(self, connection_index: int = 0,
                                     session_index: int = 0) -> SessionInfo:
        """
        Connect to an already open SAP session.

        Args:
            connection_index: Index of the connection (0 = first)
            session_index: Index of the session within the connection (0 = first)

        Returns:
            SessionInfo with session details
        """
        try:
            found = self._discover_connections()

            if not found:
                raise SAPGUIError("No SAP connections found")

            if connection_index < 0 or connection_index >= len(found):
                summary = "; ".join(
                    f"{i}: {self._connection_label(conn, host)}"
                    for i, (host, _rot, _engine, conn) in enumerate(found)
                )
                raise SAPGUIError(
                    f"Connection index {connection_index} out of range. "
                    f"Found {len(found)} connection(s): {summary}"
                )

            _host, rot_name, engine, connection = found[connection_index]

            if self._is_scripting_disabled(connection):
                raise SAPGUIError(
                    f"Connection {connection_index}: SAP GUI Scripting is disabled "
                    "on this server (profile parameter sapgui/user_scripting). "
                    "No session can be attached."
                )

            session_count = connection.Children.Count
            if session_count == 0:
                raise SAPGUIError(
                    f"Connection {connection_index} has no open sessions."
                )
            if session_index < 0 or session_index >= session_count:
                raise SAPGUIError(
                    f"Session index {session_index} out of range. "
                    f"Available: 0-{session_count - 1}"
                )

            self._connection = connection
            self._application = engine
            self._session = connection.Children(session_index)
            self._owns_session = False

            logger.info(
                "Connected to existing session %s/%s via %s",
                connection_index, session_index, rot_name,
            )
            self._read_session_traits()
            return self.get_session_info()

        except SAPGUIError:
            raise
        except Exception as e:
            logger.warning(
                "Failed to connect to existing session %s/%s: %s",
                connection_index,
                session_index,
                e,
                exc_info=logger.isEnabledFor(logging.DEBUG),
            )
            raise SAPGUIError("Failed to connect to the requested SAP session.")

    def disconnect(self):
        """Detach from the current SAP session.

        Sessions opened by this controller are closed on disconnect. Sessions
        attached via connect_to_existing_session() are left open.
        """
        if self._session and self._owns_session and self._connection:
            try:
                self._connection.CloseSession(self._session.Id)
            except Exception:
                pass
        self._session = None
        self._connection = None
        self._application = None
        self._owns_session = False
        self.sap_gui_version = ""
        self.scripting_read_only = False
        logger.info("Disconnected")

    def _read_session_traits(self) -> None:
        """Cache the SAP GUI release and the server's scripting mode.

        Both are fixed for a bound session (a /o rebind stays on the same
        server and engine). MajorVersion is a number: 8100 for 8.10, 7700
        for 7.70.
        """
        try:
            major = int(self._application.MajorVersion)
            patch = int(self._application.Patchlevel)
            release = (f"{major // 1000}.{major % 1000 // 10:02d}"
                       if major >= 1000 else str(major))
            self.sap_gui_version = f"{release} PL{patch}"
        except Exception as e:
            logger.debug("SAP GUI version not readable: %s", e)
            self.sap_gui_version = ""
        try:
            self.scripting_read_only = bool(
                self._session.Info.ScriptingModeReadOnly
            )
        except Exception as e:
            logger.debug("Scripting mode not readable: %s", e)
            self.scripting_read_only = False

    def _find_session_by_id(self, session_id: str):
        """Locate a session by ID and return its connection plus session object."""
        app = self._get_application()
        for connection_index in range(app.Children.Count):
            connection = app.Children(connection_index)
            for session_index in range(connection.Children.Count):
                session = connection.Children(session_index)
                if getattr(session, "Id", None) == session_id:
                    return connection, session
        return None, None

    def _rebind_after_new_session(
        self, previous_session_id: str, previous_session_count: int | None,
    ) -> None:
        """Rebind to the newest session after a /o command opens a new window."""
        try:
            active_session = self._get_application().ActiveSession
            active_session_id = getattr(active_session, "Id", None)
            if active_session_id and active_session_id != previous_session_id:
                connection, session = self._find_session_by_id(active_session_id)
                if session is not None:
                    self._connection = connection
                    self._session = session
                    return
        except Exception:
            pass

        if self._connection is None or previous_session_count is None:
            return

        try:
            current_session_count = self._connection.Children.Count
            if current_session_count > previous_session_count:
                self._session = self._connection.Children(current_session_count - 1)
        except Exception:
            pass

    def get_session_info(self) -> SessionInfo:
        """Get information about the current session."""
        self._require_session()

        info = self._session.Info
        return SessionInfo(
            system_name=info.SystemName,
            system_number=str(info.SystemNumber),
            client=info.Client,
            user=info.User,
            language=info.Language,
            transaction=info.Transaction,
            program=info.Program,
            screen_number=info.ScreenNumber,
            session_number=info.SessionNumber,
            sap_gui_version=self.sap_gui_version,
            scripting_read_only=self.scripting_read_only,
        )

    def _connection_label(self, conn, host: str) -> str:
        """Short human label for a connection, used in error messages."""
        description = ""
        try:
            description = str(conn.Description or "")
        except Exception:
            pass
        if not description:
            try:
                description = str(conn.Children(0).Info.SystemName or "")
            except Exception:
                description = ""
        return f"{description or '(no description)'} [{host}]"

    def list_connections(self) -> List[Dict[str, Any]]:
        """List open SAP connections and sessions across all SAP GUI engines.

        Covers SAP Logon as well as SAP GUI server processes (SAP Business
        Client, embedded SAP GUI). ``index`` is the global connection index
        accepted by connect_to_existing_session().
        """
        connections = []
        for i, (host, rot_name, _engine, conn) in enumerate(
            self._discover_connections()
        ):
            sessions = []

            conn_desc = ""
            try:
                conn_desc = str(conn.Description or "")
            except Exception:
                conn_desc = ""

            scripting_disabled = self._is_scripting_disabled(conn)

            try:
                session_count = conn.Children.Count
            except Exception:
                session_count = 0

            for j in range(session_count):
                try:
                    sess = conn.Children(j)
                    info = sess.Info
                    entry_s: Dict[str, Any] = {
                        "index": j,
                        "id": sess.Id,
                        "user": info.User,
                        "transaction": info.Transaction,
                        "system": info.SystemName,
                        "client": info.Client,
                    }
                    # The number SAP shows in the title, "DEV(2)", and the
                    # window title: together they tell several sessions on
                    # one system apart.
                    try:
                        entry_s["session_number"] = int(info.SessionNumber)
                    except Exception:
                        pass
                    try:
                        entry_s["title"] = str(sess.findById("wnd[0]").Text)
                    except Exception:
                        pass
                    sessions.append(entry_s)
                except Exception as e:
                    sessions.append(
                        self._error_result(
                            {"index": j},
                            e,
                            "Could not inspect SAP session",
                        )
                    )

            if not conn_desc:
                # Business Client connections have an empty Description.
                first = next((x for x in sessions if x.get("system")), None)
                if first:
                    conn_desc = str(first["system"])
                    if first.get("client"):
                        conn_desc += f" ({first['client']})"

            entry: Dict[str, Any] = {
                "index": i,
                "id": getattr(conn, 'Id', f"conn_{i}"),
                "description": conn_desc,
                "session_count": session_count,
                "sessions": sessions,
                "host": host,
                "rot_entry": rot_name,
                "scripting_disabled": scripting_disabled,
            }
            if scripting_disabled:
                entry["note"] = _SCRIPTING_DISABLED_NOTE
            connections.append(entry)

        return connections

    # =========================================================================
    # Transaction & Navigation
    # =========================================================================

    def execute_transaction(self, tcode: str) -> Dict[str, Any]:
        """
        Execute a transaction code.

        Uses Session.StartTransaction() when possible (cleaner API),
        falls back to okcd + sendVKey for /o (new window) prefix.

        Args:
            tcode: Transaction code (e.g., "MM03", "VA01", "SE80")
                   Can include /n prefix for new session or /o for new window

        Returns:
            Dict with transaction and screen info
        """
        self._require_session()
        try:
            # Ensure proper format
            tcode = tcode.strip()
            if not tcode.startswith("/"):
                tcode = f"/n{tcode}"

            logger.info("Executing transaction: %s", tcode)

            # /o and /* prefixes must use okcd approach (not StartTransaction)
            upper = tcode.upper()
            opens_new_session = upper.startswith("/O")
            previous_session_id = getattr(self._session, "Id", "")
            previous_session_count = None
            if opens_new_session and self._connection is not None:
                try:
                    previous_session_count = self._connection.Children.Count
                except Exception:
                    previous_session_count = None

            # A bare /n leaves the current transaction; StartTransaction("")
            # has nothing to start, so it goes through the command field too.
            leaves_transaction = upper.strip() == "/N"

            if opens_new_session or leaves_transaction or upper.startswith("/*"):
                self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
                self._session.findById("wnd[0]").sendVKey(VKey.ENTER)
                if opens_new_session and previous_session_id:
                    self._rebind_after_new_session(
                        previous_session_id, previous_session_count,
                    )
            else:
                # Use StartTransaction for /n prefix (preferred API)
                plain_tcode = tcode.removeprefix("/n").removeprefix("/N")
                try:
                    self._session.StartTransaction(plain_tcode)
                except Exception:
                    # Fallback to okcd approach
                    self._session.findById("wnd[0]/tbar[0]/okcd").text = tcode
                    self._session.findById("wnd[0]").sendVKey(VKey.ENTER)

            return {
                "transaction": _strip_tcode_prefix(tcode),
                "screen": self.get_screen_info(),
            }
        except Exception as e:
            return self._error_result(
                {"transaction": _strip_tcode_prefix(tcode)},
                e,
                "Could not execute transaction",
            )

    def send_vkey(self, vkey: int, window: str = "wnd[0]") -> Dict[str, Any]:
        """
        Send a virtual key to SAP.

        Args:
            vkey: Virtual key code (use VKey enum)
            window: Target window ID (default: main window)

        Returns:
            Dict with screen info after key press
        """
        self._require_session()

        try:
            window = self._validate_window_id(window)
            target = self._find_window(window)
            refusal = self._vkey_refusal(target, window, vkey)
            if refusal is not None:
                return refusal
            logger.debug("Sending VKey %s to %s", vkey, window)
            target.sendVKey(vkey)
            return {"vkey": vkey, "screen": self.get_screen_info()}
        except Exception as e:
            return self._error_result(
                {"vkey": vkey, "window": window},
                e,
                "Could not send SAP key",
            )

    def check_vkey(self, vkey: int, window: str = "wnd[0]") -> Optional[Dict[str, Any]]:
        """Return send_vkey's refusal for *vkey*, or None if it would be sent.

        Sends nothing. Lets a caller that asks the user first (the Save key)
        skip the question for a key the screen would refuse anyway.
        """
        self._require_session()

        try:
            window = self._validate_window_id(window)
            return self._vkey_refusal(self._find_window(window), window, vkey)
        except Exception as e:
            return self._error_result(
                {"vkey": vkey, "window": window},
                e,
                "Could not send SAP key",
            )

    def _vkey_refusal(self, target, window: str, vkey: int) -> Optional[Dict[str, Any]]:
        """The error result for a key *target* does not accept, else None."""
        if self._vkey_allowed(target, vkey):
            return None
        # Sending it would raise COM error 617 "The virtual key is
        # not enabled": say so, and which keys the screen does take.
        return {
            "vkey": vkey,
            "window": window,
            "error": f"{self._vkey_name(vkey)} is not active on this screen",
            "available_keys": [
                self._vkey_name(key) for key in self._SENDABLE_VKEYS
                if self._vkey_allowed(target, key)
            ],
            "screen": self.get_screen_info(),
        }

    # The keys sap_send_key offers: Enter, F1-F12, Shift+F1-F9, Ctrl+F/G/P.
    _SENDABLE_VKEYS = (*range(0, 22), 71, 84, 86)

    @staticmethod
    def _vkey_allowed(window, vkey: int) -> bool:
        """GuiFrameWindow.IsVKeyAllowed; True when it cannot be asked."""
        try:
            return bool(window.IsVKeyAllowed(vkey))
        except Exception:
            return True

    @staticmethod
    def _vkey_name(vkey: int) -> str:
        """The sap_send_key name of a virtual key code."""
        if vkey == 0:
            return "Enter"
        if 1 <= vkey <= 12:
            return f"F{vkey}"
        if 13 <= vkey <= 24:
            return f"Shift+F{vkey - 12}"
        return {71: "Ctrl+F", 84: "Ctrl+G", 86: "Ctrl+P"}.get(vkey, str(vkey))

    def press_enter(self) -> Dict[str, Any]:
        """Press Enter key."""
        return self.send_vkey(VKey.ENTER)

    def press_back(self) -> Dict[str, Any]:
        """Press Back (F3)."""
        return self.send_vkey(VKey.F3)

    def press_cancel(self) -> Dict[str, Any]:
        """Press Cancel (F12/ESC)."""
        return self.send_vkey(VKey.F12)

    def press_save(self) -> Dict[str, Any]:
        """Press Save (Ctrl+S/F11)."""
        return self.send_vkey(VKey.F11)

    def press_execute(self) -> Dict[str, Any]:
        """Press Execute (F8)."""
        return self.send_vkey(VKey.F8)

    def get_screen_info(self, include_popup: bool = True) -> Dict[str, Any]:
        """Get information about the current screen.

        When a popup is open the result carries a compact ``popup`` digest
        (classification, texts, button labels and any pre-filled inputs with a
        notice), unless *include_popup* is False.

        Reads from ``session.ActiveWindow`` so the title always reflects
        what the user actually sees.  When ActiveWindow is a popup
        (wnd[1], wnd[2], ...) the response ``active_window`` field tells
        the caller which window is in focus -- no separate tool call
        needed.

        ``session.Info`` (transaction, program, screen_number) already
        reflects the active screen regardless of which window is focused.
        """
        self._require_session()

        try:
            info = self._session.Info

            # Determine the active window -- use ActiveWindow when
            # available, fall back to wnd[0].
            active_wnd_id = "wnd[0]"
            active = None
            try:
                active = self._session.ActiveWindow
                if active is not None:
                    active_wnd_id = self._normalize_window_id(active.Id)
            except Exception:
                pass

            # Read title from the active window
            try:
                title = self._session.findById(active_wnd_id).Text
            except Exception:
                title = ""

            # Status bar -- try active window first, fall back to wnd[0]
            status = self._get_status_bar_info(active_wnd_id)

            result: Dict[str, Any] = {
                "active_window": active_wnd_id,
                "transaction": info.Transaction,
                "program": info.Program,
                "screen_number": info.ScreenNumber,
                "title": title,
                "message": status.get("text"),
                "message_type": status.get("message_type", ""),
                "message_id": status.get("message_id", ""),
                "message_number": status.get("message_number", ""),
            }
            # Only when set: screen info rides on every action response.
            if status.get("has_long_text"):
                result["message_has_long_text"] = True
            if status.get("as_popup"):
                result["message_as_popup"] = True
            subscreens = self._subscreens(active_wnd_id)
            if subscreens:
                result["subscreens"] = subscreens
            cursor = self._cursor_field(active)
            if cursor:
                result["cursor"] = cursor

            # A popup opened: say what it is in THIS response, so the caller
            # sees pre-filled values before deciding to confirm it.
            if include_popup and active_wnd_id != "wnd[0]":
                digest = getattr(self, "_popup_digest", lambda: {})()
                if digest:
                    result["popup"] = digest

            return result
        except Exception as e:
            return self._error_result({}, e, "Could not read screen information")

    def _subscreens(self, window_id: str) -> List[str]:
        """The sub-screens on the window, outer first, as ``program:dynnr``.

        Info.Program / ScreenNumber name only the main screen: an RF step,
        a tab page or a header/item area is a sub-screen, and when only it
        changes the screen info looks the same. One GetObjectTree call
        (1-9 ms live, against up to 180 ms walking the elements over COM);
        an older SAP GUI without it reports none.
        """
        tree = self._object_tree(f"{window_id}/usr", ["Id"], site="subscreens")
        if tree is None:
            return []
        found: List[str] = []
        stack = [tree]
        while stack:
            node = stack.pop()
            node_id = (node.get("properties") or {}).get("Id", "")
            match = _SUBSCREEN_ID_RE.search(node_id)
            if match:
                found.append(f"{match.group(1)}:{match.group(2)}")
            stack.extend(reversed(node.get("children") or []))
        return list(dict.fromkeys(found))

    def _cursor_field(self, window) -> str:
        """ID of the field the cursor is in, '' when there is none.

        SystemFocus is where the application put the cursor. GuiFocus is
        not: after a transaction start it stays on the command field while
        the cursor sits in the first input field. A redisplay that moves the
        cursor to a required field without a message is visible only here.
        """
        try:
            return self._normalize_element_id(window.SystemFocus.Id)
        except Exception:
            return ""

    def _get_status_bar_info(
        self, window_id: str = "wnd[0]",
    ) -> Dict[str, Any]:
        """Get structured information from the status bar.

        Tries the given *window_id*'s status bar first (e.g. a popup's
        ``wnd[1]/sbar``).  Falls back to ``wnd[0]/sbar`` when the
        active window doesn't have one.

        Returns a dict with:
        - text: The full message text
        - message_type: S=Success, W=Warning, E=Error, A=Abort, I=Info
        - message_id: SAP message class (e.g., 'MM', 'SD')
        - message_number: Three-digit message number
        - message_parameters: List of up to 4 message parameters
        """
        sbar = None
        for wnd in dict.fromkeys([window_id, "wnd[0]"]):
            try:
                sbar = self._session.findById(f"{wnd}/sbar")
                break
            except Exception:
                continue
        if sbar is None:
            return {"text": None}
        try:
            info: Dict[str, Any] = {"text": sbar.Text}
            for attr, key in [
                ("MessageType", "message_type"),
                ("MessageId", "message_id"),
                ("MessageNumber", "message_number"),
            ]:
                try:
                    value = getattr(sbar, attr)
                    # SAP pads MessageId to 20 characters ("00          ...")
                    info[key] = value.strip() if isinstance(value, str) else value
                except Exception:
                    info[key] = ""
            # Message parameters (up to 4)
            params = []
            for attr in ["MessageParameter", "MessageParameter1",
                         "MessageParameter2", "MessageParameter3"]:
                try:
                    val = getattr(sbar, attr, None)
                    if val is not None:
                        params.append(str(val))
                except Exception:
                    pass
            if params:
                info["message_parameters"] = params
            # MessageHasLongText (7.60 PL2+): -1 no message, 0 no, 1 yes.
            # MessageAsPopup: the message also sits in a popup that must be
            # closed before anything else works.
            try:
                if int(sbar.MessageHasLongText) == 1:
                    info["has_long_text"] = True
            except Exception:
                pass
            try:
                if sbar.MessageAsPopup:
                    info["as_popup"] = True
            except Exception:
                pass
            return info
        except Exception:
            return {"text": None}

    def _get_status_bar_message(self) -> Optional[str]:
        """Get the message text from the status bar (legacy helper)."""
        return self._get_status_bar_info().get("text")

    def _safe_set_field(self, field_id: str, value: str) -> bool:
        """Set field value, returning False on error instead of raising."""
        try:
            self._session.findById(field_id).text = value
            return True
        except Exception:
            return False
