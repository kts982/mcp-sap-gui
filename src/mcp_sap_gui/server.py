"""
MCP Server for SAP GUI Scripting.

This module implements a Model Context Protocol (MCP) server that exposes
SAP GUI automation capabilities to AI assistants like Claude.

Uses FastMCP from the official MCP Python SDK for decorator-based tool
definitions with automatic JSON schema generation from type hints.

Security Note:
    This server provides powerful automation capabilities. Production deployments
    should implement:
    - Transaction whitelisting
    - Read-only modes
    - Audit logging
    - Rate limiting
    - User confirmation for write operations
"""

import asyncio
import json
import logging
import os
import re
import time
from dataclasses import dataclass, field
from typing import List, Literal, Optional

from dotenv import load_dotenv
from fastmcp import Context, FastMCP
from fastmcp.apps import UI_EXTENSION_ID
from fastmcp.exceptions import ToolError
from fastmcp.server.lifespan import lifespan
from fastmcp.tools import ToolResult
from mcp.shared.exceptions import McpError
from mcp.types import ImageContent, TextContent

from .audit import AuditMiddleware
from .confirmation import (
    ATTACH_POINT,
    SAVE_POINT,
    SETTABLE_POINTS,
    ConfirmationMiddleware,
    ask_user,
    build_relax_prompt,
    effective_points,
    log_attach_event,
    log_confirmation_event,
    points_with_provenance,
)
from .controller import (
    _DEAD_SESSION_HRESULTS,
    _HRESULT_DISCONNECTED,
    _HRESULT_SERVER_UNAVAILABLE,
    OBJECT_TREE_DISABLE_ENV,
    OBJECT_TREE_TRACE_ENV,
    object_tree_disabled_sites,
)
from .models import SAPGUINotConnectedError
from .policy import (
    DEFAULT_PRESET,
    POLICY_FILE_ENV,
    PRESETS,
    PolicyError,
    TransactionPolicy,
    load_policy,
    normalize_pattern,
    resolve_policy_file,
)
from .preview import (
    SCREENSHOT_INCLUDED,
    SCREENSHOT_OMITTED,
    SCREENSHOT_UNAVAILABLE,
    build_preview_card,
    build_preview_text,
    downscale_png_b64,
    prefab_available,
)
from .prompts import (
    WORKFLOW_TARGET_PARAMETERS,
    WorkflowName,
    normalize_transaction,
    register_prompts,
    render_transaction_guide,
    render_workflow_guide,
)
from .sap_controller import VKey
from .session_manager import SessionManager
from .strict import STRICT_ERRORS_ENV, StrictErrorMiddleware, strict_errors_from_env

logger = logging.getLogger(__name__)


# Kept under this name for callers and tests; the presets live in policy.py.
_DEFAULT_BLOCKED_TRANSACTIONS = list(PRESETS[DEFAULT_PRESET]["block"])

_TRANSACTION_PREFIX_RE = re.compile(r"^(?:/(?:N|O|\*)\s*)+", re.IGNORECASE)
_TRANSACTION_CODE_RE = re.compile(r"^/?[A-Z0-9_]+(?:/[A-Z0-9_]+)*$")
_COMMAND_FIELD_EXACT_NAMES = {
    "okcd",
    "okcode",
    "okcodefield",
    "okcodeinput",
    "ok_code",
    "tcode",
    "transaction",
    "transactioncode",
    "transaction_code",
    "command",
    "commandcode",
    "command_code",
}
_COMMAND_FIELD_PREFIXES = ("txt", "ctxt", "pwd", "cmb")


def _normalize_transaction_code(raw: str, *, strict: bool = True) -> str | None:
    """Normalize an SAP transaction code for policy checks.

    Strips command prefixes such as ``/n`` or ``/o``, preserves slash-style
    transactions like ``/SCWM/MON``, uppercases the result, and optionally
    validates the remaining shape.
    """
    if not isinstance(raw, str):
        if strict:
            raise ValueError("Transaction code must be a string")
        return None

    value = raw.strip().upper()
    if not value:
        if strict:
            raise ValueError("Transaction code cannot be empty")
        return None

    while value.startswith("="):
        value = value[1:].lstrip()
    value = _TRANSACTION_PREFIX_RE.sub("", value).strip()
    if not value:
        if strict:
            raise ValueError("Transaction code cannot be empty")
        return None

    if not _TRANSACTION_CODE_RE.fullmatch(value):
        if strict:
            raise ValueError(
                f"Invalid SAP transaction code: {raw!r}. "
                "Expected forms like 'MM03', 'VA03', or '/SCWM/MON'."
            )
        return None
    return value


def _normalize_transaction_list(codes: Optional[List[str]]) -> Optional[List[str]]:
    """Normalize and de-duplicate a list of transaction codes."""
    if codes is None:
        return None

    normalized: list[str] = []
    seen: set[str] = set()
    for code in codes:
        canonical = _normalize_transaction_code(code)
        if canonical not in seen:
            seen.add(canonical)
            normalized.append(canonical)
    return normalized


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Profile definitions: which tag sets are visible in each profile.
# Declared here (not next to the profile tool) because ServerConfig validates
# against it at import time.
_PROFILES: dict[str, set[str]] = {
    "exploration": {"read"},
    "operator": {"read", "write"},
    "full": {"read", "write", "destructive"},
}


@dataclass
class ServerConfig:
    """Configuration for the MCP SAP GUI server."""

    # Security settings
    read_only: bool = False
    allowed_transactions: Optional[List[str]] = None  # None = all allowed
    blocked_transactions: List[str] = field(
        default_factory=lambda: list(_DEFAULT_BLOCKED_TRANSACTIONS)
    )
    # Allow exceptions: codes/patterns that win over blocked_transactions.
    allowed_exceptions: List[str] = field(default_factory=list)
    # Where the two lists came from (--policy-preset / --policy-file); only
    # used to describe the effective policy in logs and error texts.
    policy_preset: str = DEFAULT_PRESET
    policy_source: str = "preset"
    policy_file_sha256: Optional[str] = None
    # Server-level policy floor (--profile). A session may restrict below this
    # but never above it. Modelled on --read-only: no runtime mutator exists.
    profile: str = "full"
    # Server-level confirmation points (--confirm). Always active, never
    # removable by sap_set_confirmation_points.
    confirmation_floor: List[str] = field(default_factory=list)
    # --strict-errors / MCP_SAP_GUI_STRICT_ERRORS: return a failed result
    # (strict.result_error) as an MCP error (isError true). Off by default.
    strict_errors: bool = False

    # Behavior settings
    auto_connect_existing: bool = True
    default_language: str = "EN"
    max_table_rows: int = 500

    def __post_init__(self):
        # Entries may be glob patterns ("SU*", "*"), so they go through the
        # pattern normalizer; plain codes still get the /n-stripping one.
        self.blocked_transactions = list(dict.fromkeys(
            normalize_pattern(entry, _normalize_transaction_code)
            for entry in self.blocked_transactions
        ))
        self.allowed_exceptions = list(dict.fromkeys(
            normalize_pattern(entry, _normalize_transaction_code)
            for entry in self.allowed_exceptions
        ))
        self.policy = TransactionPolicy(
            block=self.blocked_transactions,
            allow=self.allowed_exceptions,
            preset=self.policy_preset,
            source=self.policy_source,
            file_sha256=self.policy_file_sha256,
        )
        self.allowed_transactions = _normalize_transaction_list(self.allowed_transactions)
        unknown = sorted(set(self.confirmation_floor) - set(SETTABLE_POINTS))
        if unknown:
            raise ValueError(
                f"Unknown confirmation point(s): {', '.join(unknown)}. "
                f"Valid points: {', '.join(SETTABLE_POINTS)}"
            )
        self.confirmation_floor = sorted(set(self.confirmation_floor))
        # Validated here so _effective_profile can trust it: an unrecognised
        # profile must not silently degrade into "no floor at all".
        if self.profile not in _PROFILES:
            raise ValueError(
                f"Unknown policy profile: {self.profile!r}. "
                f"Valid profiles: {', '.join(sorted(_PROFILES))}"
            )


# ---------------------------------------------------------------------------
# Tool annotation presets
# ---------------------------------------------------------------------------

_READ_ONLY = {"readOnlyHint": True, "destructiveHint": False}
_WRITE = {"readOnlyHint": False, "destructiveHint": False}
_DESTRUCTIVE = {"readOnlyHint": False, "destructiveHint": True}

# Tag sets for policy profiles (fastmcp visibility system)
_TAGS_READ = {"read"}
_TAGS_WRITE = {"write"}
_TAGS_DESTRUCTIVE = {"destructive"}


# ---------------------------------------------------------------------------
# Lifespan
# ---------------------------------------------------------------------------

@lifespan
async def _lifespan(server: FastMCP):
    """Create the session manager on startup, clean up on shutdown."""
    global _session_mgr
    _session_mgr = SessionManager()
    try:
        yield {"session_mgr": _session_mgr}
    finally:
        if _session_mgr is not None:
            try:
                _session_mgr.executor.submit(
                    _session_mgr.release_all
                ).result(timeout=10)
            except Exception:
                pass
            _session_mgr.shutdown()
            _session_mgr = None


# ---------------------------------------------------------------------------
# MCP Instructions (injected into every client's system prompt)
# ---------------------------------------------------------------------------

_INSTRUCTIONS = """\
SAP GUI automation server. Controls SAP GUI for Windows via COM scripting.

## Rules (read first)

- **Never guess element IDs.** Discover them on every new screen \
(`sap_get_screen_elements`, `sap_get_toolbar_buttons`): IDs differ between \
systems and customizations.
- **Check `screen.active_window` after every action.** `wnd[1]` or higher is a \
popup: read it with `sap_get_popup_window` before you go on.
- **Confirmations.** F11 / Save always asks the user first; \
`sap_set_confirmation_points` can add `transactions`, `batch_fields`, \
`field_writes`, `all_writes` and `attach`. A "declined" error is a stop: do not \
retry or route around it, tell the user. Offer `sap_preview` before significant \
writes (batch fills, Save) so the user sees the screen and the pending values.

## Getting Started

1. SAP already logged on: `sap_connect_existing`; pass `system` (e.g. 'DEV') \
when the user named the system. SAP Business Client sessions work too \
(experimental).
2. SAP not open: `sap_connect` with the SAP Logon entry.
3. `sap_get_session_info` shows the current transaction and screen.

## Screen Discovery

`sap_get_screen_elements` with `type_filter` (e.g. "GuiTextField,GuiCTextField") \
and `changeable_only=true` keeps the output small; `sap_read_field` reads one \
field. A tree docked beside the screen (SM34 view clusters, IMG activities) is \
not under `wnd[0]/usr`: discovery lists it under `docking_containers` \
(e.g. `wnd[0]/shellcont`).

## Popup Handling

Popups often follow button presses, Enter, F8 and double-clicks. \
`sap_get_popup_window` gives the text, the buttons, a classification and a safe \
next step; `sap_handle_popup` acts on it.

## Tables

`sap_read_table` detects the type (`table_type`). Start with \
`columns_only=true`, then fetch only `columns="COL_A,COL_B"`.
- ALV grid (reports): page with `start_row` until `total_rows`; its toolbar is \
`sap_get_alv_toolbar`.
- Table control (SM30, SPRO): one read returns the visible rows; page with \
`start_row=next_start_row` while `has_more` is true. `total_rows` counts empty \
lines; `empty_rows_from` is where the data ends. To jump to an entry use the \
"Position..." toolbar button, not manual scrolling.

## SPRO and Trees

Never `sap_read_tree` SPRO (1000+ nodes). `sap_search_tree_nodes` searches \
expanded nodes only: expand parents with \
`sap_get_tree_node_children(expand=true)`. Run an activity with \
`sap_click_tree_link(tree_id, node_key, "2")`; a double-click opens its \
documentation. In a view cluster (SM34) use `double_click_tree_item` with the \
column name (usually "Column1").

## Keys and Common Mistakes

Enter confirms, F3 goes back, F8 executes, F11 saves, F12 cancels. F4 is value \
help on the focused field (`sap_set_focus` first): the popup's classification \
says what opened (`list`, `date_picker`, `time_picker`, `input_required`), and \
`sap_get_workflow_guide("search_help", field_id)` covers each. Dropdowns (`cmb`) \
have no F4: `sap_set_field` takes the entry's key or text. In table maintenance \
F5 means "New Entries", not refresh: check the toolbar tooltips first.
"""


# ---------------------------------------------------------------------------
# MCP Resource: detailed SAP GUI reference guide
# ---------------------------------------------------------------------------

_SAP_GUI_GUIDE = """\
# SAP GUI Navigation Reference for AI Agents

Detailed reference guide for automating SAP GUI via MCP tools.

## Element Types and ID Patterns

| Type | Description | Typical ID Pattern |
|------|-------------|-------------------|
| GuiTextField | Text input | `wnd[0]/usr/txtFIELD_NAME` |
| GuiCTextField | Text with search help (F4) | `wnd[0]/usr/ctxtFIELD_NAME` |
| GuiPasswordField | Password input | `wnd[0]/usr/pwdFIELD_NAME` |
| GuiButton | Pushbutton | `wnd[0]/tbar[1]/btn[8]` or `wnd[0]/usr/btnBUTTON_NAME` |
| GuiCheckBox | Checkbox | `wnd[0]/usr/chkFIELD_NAME` |
| GuiRadioButton | Radio button | `wnd[0]/usr/radFIELD_NAME` |
| GuiComboBox | Dropdown | `wnd[0]/usr/cmbFIELD_NAME` |
| GuiTab | Tab in strip | `wnd[0]/usr/tabsTABSTRIP/tabpTAB01` |
| GuiMenu | Menu bar item | `wnd[0]/mbar/menu[0]/menu[1]` |
| GuiGridView | ALV Grid | `wnd[0]/usr/cntlGRID/shellcont/shell` |
| GuiTableControl | Classic table | `wnd[0]/usr/tblSAPLBD41TCTRL_V_TBDLS` |
| GuiTree | Tree control | `wnd[0]/usr/cntlTREE/shellcont/shell` |
| GuiStatusbar | Status bar | `wnd[0]/sbar` |
| GuiOkCodeField | Command/OK-code field | `wnd[0]/tbar[0]/okcd` |

## ID Naming Conventions

- `wnd[0]` = main window, `wnd[1]` = first popup, `wnd[2]` = second popup
- `usr` = user area (screen body)
- `tbar[0]` = system toolbar (top), `tbar[1]` = application toolbar
- `mbar` = menu bar
- `btn[N]` = toolbar button by position
- `shell/shellcont/shell` = container for ALV grids and tree controls
- `txt` prefix = GuiTextField, `ctxt` = GuiCTextField, `chk` = GuiCheckBox
- `rad` = GuiRadioButton, `cmb` = GuiComboBox, `pwd` = GuiPasswordField

## Transaction Code Conventions

- Standard: `MM03` (display material), `VA03` (display sales order)
- With /n prefix: `/nMM03` — closes current transaction, opens new one
- With /o prefix: `/oMM03` — opens in new session (new window)
- SCWM transactions need /n: `/n/SCWM/MON`, `/n/SCWM/PRDO`

## Table Type Comparison

### GuiGridView (ALV)
- Used in: Reports, list displays, transaction results
- ID contains: `shellcont/shell`
- All rows accessible directly (internal scrolling)
- Has built-in toolbar: sort, filter, export, layout
- Use `sap_get_alv_toolbar` for toolbar button discovery
- Use `sap_press_alv_toolbar_button` to execute toolbar actions
- Pagination: `start_row` parameter in `sap_read_table`

### GuiTableControl
- Used in: SPRO customizing, SM30 table maintenance, config screens
- ID starts with: `tbl`
- Only visible rows accessible at a time
- Scrolling via `VerticalScrollbar.Position`
- Pagination: `start_row` in `sap_read_table` (handles scrolling internally)
- "Position..." button for jumping to specific entries
- F5 = "New Entries" (NOT refresh!)

## Reading Status Bar Messages

The status bar (`wnd[0]/sbar`) shows messages after actions:
- **S** (Success): Operation completed successfully
- **E** (Error): Operation failed — read the message for details
- **W** (Warning): Warning that may need acknowledgment
- **I** (Info): Informational message

Use `sap_read_field("wnd[0]/sbar")` to read the status bar after actions.

## SPRO Navigation Pattern (Step-by-Step)

SPRO (transaction SPRO) is SAP's customizing tree with 1000+ nodes.

1. Execute transaction: `sap_execute_transaction("SPRO")`
2. Click "SAP Reference IMG" button (check toolbar with `sap_get_toolbar_buttons`)
3. Search for your target: `sap_search_tree_nodes(tree_id, "your search text")`
   - **Limitation**: only searches already-expanded/loaded nodes
   - If no results, expand parent nodes with `sap_get_tree_node_children(expand=true)` first
4. Navigate to the node using `sap_get_tree_node_children` with `expand=true`
5. Execute the activity: `sap_click_tree_link(tree_id, node_key, "2")`
   - Column "2" is the execute/activity icon in SPRO trees
   - Do NOT use `double_click_tree_node` — that opens documentation
6. Handle the resulting screen (often a popup selection screen or table maintenance view)

## Table Maintenance Views (SM30-style)

These appear when you execute SPRO activities or use SM30:

1. Usually shows a selection screen first (popup on wnd[1])
2. After execution, shows a GuiTableControl with configuration entries
3. Toolbar has: "New Entries", "Position...", "Delete", "Save"
4. To find specific entries, use "Position..." button (opens search dialog)
5. To read all entries, use `sap_read_table` with `start_row` pagination: \
pass the previous page's `next_start_row` and stop when `has_more` is false
6. `total_rows` counts empty lines too, so it is not the entry count: the data \
ends at `empty_rows_from` when the response has it

## Classic Lists (WRITE Output)

A report that prints a plain list (program `SAPMSSY0`, screen 120) has no table \
object — every word is a `GuiLabel` named `lbl[col,row]`. Discovery leaves \
these cells out and reports the list under `lists` instead:
- Read it with `sap_read_list`: the visible page as lines of text, plus the list \
colour of a line (`negative` = red / error, `positive`, `total`, `heading`)
- Page through a long list with `scroll_to` (see `scroll` in the response)
- F4 hit lists in a popup are the same kind of list: `window_id="wnd[1]"`
- To act on a line, read with `with_ids=true`, `sap_set_focus` that label, then F2 \
(Enter to apply a line of an F4 hit list)

## Web Dynpro Screens

Some newer SAP screens use Web Dynpro technology:
- `sap_get_screen_elements` may return 0 elements on wnd[0]/usr
- Use `sap_screenshot` as fallback to see what's displayed
- Element IDs follow different patterns (deeper nesting)
- Try increasing `max_depth` in `sap_get_screen_elements`

## Handling Splitter Layouts

Some transactions use split-screen layouts:
- Elements are nested inside `shellcont[0]`, `shellcont[1]`, etc.
- Use `sap_get_screen_elements` with `max_depth=3` or higher
- Example: Warehouse Monitor `/SCWM/MON` uses splitter with tree in `shellcont[0]`

## Docking Containers (View Clusters)

A docking container is attached to the window BESIDE the user area, so its ID \
starts with `wnd[0]/shellcont` instead of `wnd[0]/usr`:
- The dialog-structure tree of a view cluster (SM34, most IMG activities) is \
typically `wnd[0]/shellcont/shell`; the tables it opens stay under `wnd[0]/usr`
- `sap_get_screen_elements` on the user area reports them as `docking_containers`
- `sap_get_screen_elements(container_id="wnd[0]", max_depth=1)` lists every \
top-level area of a window (title, menu bar, toolbars, user area, status bar, \
docking containers)
- Read the structure with `sap_read_tree`; switch to a node's view with \
`sap_double_click_tree_item(tree_id, node_key, "Column1")` (take the column name \
from `column_names`). `sap_double_click_tree_node` reports success there but does \
NOT switch the view — check that the window title changed
- A dependent node (child level) needs its parent entry selected in the table first

## Known Limits of SAP GUI Scripting

- **Drag and drop cannot be scripted.** Screens that only work by dragging \
(e.g. the RF menu manager `/SCWM/RFMENU`) cannot be driven; stop early and tell the user.
- **A running program keeps its loaded version.** After a program was changed and \
activated elsewhere, leave it (`/n`) and start it again before testing.
"""


# ---------------------------------------------------------------------------
# Module-level state
# ---------------------------------------------------------------------------

mcp = FastMCP("mcp-sap-gui", instructions=_INSTRUCTIONS, lifespan=_lifespan)
mcp.add_middleware(AuditMiddleware())
# Next to audit, outside the confirmation gate: audit sees what it returns,
# and the gate's confirmed_by_user (added only to a result that is not an
# error) is already in the result it flags. The lambda reads the running
# module's config at call time, because main() rebinds it.
mcp.add_middleware(StrictErrorMiddleware(enabled=lambda: config.strict_errors))
# ConfirmationMiddleware is registered further down, after its injected
# helpers are defined. add_middleware is outermost-first, so audit (added
# here) stays outside the confirmation gate and blocked calls keep their
# audit line.
register_prompts(mcp)
_session_mgr: Optional[SessionManager] = None
config = ServerConfig()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# COM errors that are safe to retry: the callee refused the call before
# running it, so re-running the tool body cannot repeat any work. Values are
# pywintypes.com_error.hresult, i.e. signed 32-bit.
_TRANSIENT_COM_HRESULTS = frozenset({
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED: "Call was rejected by callee."
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER: "The message filter
                  # indicated that the application is busy."
})
# Deliberately not retried:
# - -2147417851, 0x80010105 RPC_E_SERVERFAULT ("The server threw an
#   exception."): SAP GUI raises it for an element reference that went stale
#   after a server round trip (live, 8.10: a field found before Enter, read
#   after it, even on the same screen). That fails the same way every time,
#   and the call may have half-run: a retry would re-run the whole tool body,
#   its presses and keys included, only to fail again.
# - _DEAD_SESSION_HRESULTS (0x80010108 RPC_E_DISCONNECTED, 0x800706BA
#   RPC_S_SERVER_UNAVAILABLE): the session object or the whole SAP GUI
#   process is gone and does not come back. Raised at once as
#   SAPGUINotConnectedError instead.

_COM_MAX_RETRIES = 3
_COM_BASE_DELAY = 0.3  # seconds

_DEAD_SESSION_NAMES = {
    _HRESULT_DISCONNECTED: "RPC_E_DISCONNECTED",
    _HRESULT_SERVER_UNAVAILABLE: "RPC_S_SERVER_UNAVAILABLE",
}


def _com_hresult(exc: BaseException) -> int | None:
    """Return the HRESULT a COM exception carries, or None."""
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        return hresult
    # pywintypes.com_error stores hresult in args[0]
    if exc.args and isinstance(exc.args[0], int):
        return exc.args[0]
    return None


def _is_transient_com_error(exc: Exception) -> bool:
    """Check if a COM exception is a known transient error safe to retry."""
    return _com_hresult(exc) in _TRANSIENT_COM_HRESULTS


def _com_with_retry(fn):
    """Run a COM operation with retry for transient errors.

    Runs synchronously in the COM thread. Retries up to _COM_MAX_RETRIES
    times with exponential backoff for known transient COM errors. A COM
    error saying the SAP GUI session or process is gone is raised at once as
    SAPGUINotConnectedError rather than as the raw COM tuple.
    """
    last_exc = None
    for attempt in range(_COM_MAX_RETRIES + 1):
        try:
            return fn()
        except Exception as e:
            hresult = _com_hresult(e)
            if hresult in _DEAD_SESSION_HRESULTS:
                raise SAPGUINotConnectedError(
                    "SAP GUI is no longer reachable "
                    f"({_DEAD_SESSION_NAMES[hresult]} "
                    f"0x{hresult & 0xFFFFFFFF:08X}): the session window was "
                    "closed or SAP Logon exited. Start SAP Logon if needed, "
                    "then reconnect with sap_connect_existing."
                ) from e
            if attempt < _COM_MAX_RETRIES and hresult in _TRANSIENT_COM_HRESULTS:
                delay = _COM_BASE_DELAY * (2 ** attempt)
                logger.warning(
                    "Transient COM error (attempt %d/%d), retrying in %.1fs: %s",
                    attempt + 1, _COM_MAX_RETRIES + 1, delay, e,
                )
                time.sleep(delay)
                last_exc = e
            else:
                raise
    raise last_exc  # pragma: no cover — unreachable, loop always raises


async def _com(fn):
    """Run a synchronous COM operation in the dedicated thread.

    Automatically retries known transient COM errors (RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER) up to 3 times with exponential backoff.
    """
    if _session_mgr is None:
        raise RuntimeError("Server not initialised – lifespan has not started")
    return await asyncio.get_running_loop().run_in_executor(
        _session_mgr.executor, lambda: _com_with_retry(fn)
    )


def _note_cap(result, requested: int, capped: int):
    """Say so when a tool lowered a row/node limit to config.max_table_rows."""
    if isinstance(result, dict) and capped < requested:
        result["capped_to"] = capped
    return result


def _session_key(ctx: Context) -> str:
    """Derive a stable key for the current MCP client session.

    ``ctx.session_id``, not ``id(ctx.session)``: with the MCP Python SDK v2
    (FastMCP 4) every request gets a new ServerSession object, so ``id()``
    changes on every call, and confirmation points and the SAP binding
    silently vanish between calls. ``session_id`` is the same for all calls
    of one client on FastMCP 3.x and 4.
    """
    return ctx.session_id


def _ctrl(ctx: Context):
    """Get the SAPGUIController bound to the current MCP session."""
    return _session_mgr.get_or_create(_session_key(ctx)).controller


def _check_write(ctx: Context | None = None):
    """Raise if writes are off: server read-only mode, or an SAP session whose
    server only allows read-only scripting (every set would fail in COM)."""
    if config.read_only:
        raise ValueError("Write operations disabled in read-only mode")
    if ctx is None or _session_mgr is None:
        return
    managed = _session_mgr.get_existing(_session_key(ctx))
    if managed is not None and managed.controller.scripting_read_only:
        raise ValueError(
            "This SAP session only allows read-only scripting (server profile "
            "parameter sapgui/user_scripting_set_readonly, or "
            "sapgui/nwbc_scripting in SAP Business Client): nothing can be "
            "entered, pressed or navigated through scripting. Reading works "
            "for what is on screen; grids and tables cannot be scrolled."
        )


def active_confirmation_points(ctx: Context) -> set[str]:
    """Return the confirmation points active for the calling session.

    Effective set = server floor (``--confirm``) + session-added points +
    the always-on ``save`` point.
    """
    session_points: set[str] = set()
    if _session_mgr is not None:
        try:
            session_points = _session_mgr.get_confirmation_points(_session_key(ctx))
        except Exception:  # pragma: no cover — defensive, never block a call
            session_points = set()
    return effective_points(config.confirmation_floor, session_points)


def precheck_before_confirmation(tool_name: str, args,
                                 ctx: Context | None = None) -> None:
    """Re-run the cheap in-body policy checks before eliciting.

    The confirmation middleware runs upstream of the tool body, so without
    this a rejected call would still prompt the user first — and a blocked
    transaction would reach the approval dialog. Only called for calls the
    gate is about to block on, which are writes by construction.

    Rejections are re-raised in fastmcp's own "Error calling tool 'x': ..."
    shape so a gated rejection is byte-identical to an ungated one; otherwise
    the error text is an oracle for which points are active.
    """
    get = args.get if hasattr(args, "get") else (lambda _k, _d=None: None)
    try:
        _check_write(ctx)
        if tool_name == "sap_execute_transaction":
            tcode = get("tcode")
            if isinstance(tcode, str):
                _enforce_transaction_policy(tcode)
        elif tool_name == "sap_set_field":
            _check_okcode_bypass(get("field_id"), get("value"))
        elif tool_name == "sap_set_batch_fields":
            fields = get("fields")
            if isinstance(fields, dict):
                for fid, val in fields.items():
                    _check_okcode_bypass(fid, str(val))
    except ValueError as exc:
        raise ToolError(f"Error calling tool '{tool_name}': {exc}") from exc


# Registered here, not next to AuditMiddleware: the gate needs the RUNNING
# module's helpers injected as callables. Importing the server module by name
# from the middleware instead would load a second module copy under
# `python -m mcp_sap_gui.server` and read empty globals (see
# ConfirmationMiddleware's docstring). Audit was added earlier, so it stays
# outermost.
mcp.add_middleware(ConfirmationMiddleware(
    active_points=active_confirmation_points,
    precheck=precheck_before_confirmation,
))


def _is_transaction_allowed(tcode: str) -> tuple[bool, str]:
    """Return allow/deny decision plus the canonical transaction code."""
    canonical = _normalize_transaction_code(tcode)
    if not config.policy.permits(canonical):
        return False, canonical
    if config.allowed_transactions is not None and canonical not in config.allowed_transactions:
        return False, canonical
    return True, canonical


def _is_leave_transaction_command(tcode: object) -> bool:
    """Return True for a bare ``/n``: leave the current transaction."""
    return isinstance(tcode, str) and tcode.strip().upper() == "/N"


def _enforce_transaction_policy(tcode: str, *, source: str = "transaction") -> str:
    """Raise when a transaction is not permitted by the active policy."""
    # A bare /n starts nothing, so there is no code to check against the
    # block/allow lists; it returns to the session's start screen.
    if _is_leave_transaction_command(tcode):
        return "/N"

    allowed, canonical = _is_transaction_allowed(tcode)
    if allowed:
        return canonical

    if not config.policy.permits(canonical):
        raise ValueError(
            f"Transaction {canonical} is blocked by security policy"
            f"{f' (attempted via {source})' if source != 'transaction' else ''}"
            f"{config.policy.how_to_allow(canonical)}"
        )
    raise ValueError(
        f"Transaction {canonical} is not in the allowed transaction list"
        f"{f' (attempted via {source})' if source != 'transaction' else ''}"
    )


def _compact_command_field_name(field_id: str) -> str:
    """Return a normalized field leaf name for OK-code detection."""
    if not isinstance(field_id, str):
        return ""

    normalized = field_id.replace("\\", "/").strip()
    if "wnd[" in normalized:
        normalized = "wnd[" + normalized.split("wnd[", 1)[1]
    leaf = normalized.rsplit("/", 1)[-1].lower()
    for prefix in _COMMAND_FIELD_PREFIXES:
        if leaf.startswith(prefix) and len(leaf) > len(prefix):
            leaf = leaf[len(prefix):]
            break
    return re.sub(r"[^a-z0-9]", "", leaf)


def _is_okcode_field(field_id: str) -> bool:
    """Return True when a field ID likely targets an SAP command field."""
    compact = _compact_command_field_name(field_id)
    if not compact:
        return False
    if compact in {name.replace("_", "") for name in _COMMAND_FIELD_EXACT_NAMES}:
        return True
    return (
        compact.endswith("okcd")
        or compact.endswith("okcode")
        or compact.endswith("tcode")
        or "commandcode" in compact
        or "transactioncode" in compact
    )


def _check_okcode_bypass(field_id: str, value: str):
    """Block attempts to execute blocked transactions via the OK-code field."""
    if not _is_okcode_field(field_id):
        return
    candidate = _normalize_transaction_code(value, strict=False)
    if candidate is None:
        return
    _enforce_transaction_policy(candidate, source="command field")


def _is_transaction_blocked(tcode: str) -> bool:
    """Check if a transaction is blocked by security policy."""
    allowed, _ = _is_transaction_allowed(tcode)
    return not allowed


_KEY_MAP = {
    "Enter": VKey.ENTER,
    "F1": VKey.F1, "F2": VKey.F2, "F3": VKey.F3, "Back": VKey.F3,
    "F4": VKey.F4, "F5": VKey.F5, "Refresh": VKey.F5,
    "F6": VKey.F6, "F7": VKey.F7, "F8": VKey.F8, "Execute": VKey.F8,
    "F9": VKey.F9, "F10": VKey.F10, "F11": VKey.F11, "Save": VKey.F11,
    "F12": VKey.F12, "Cancel": VKey.F12,
    # Shift+F keys (common in SAP: Shift+F1=WhereUsed, Shift+F4=CloseAll, etc.)
    "Shift+F1": VKey.SHIFT_F1, "Shift+F2": VKey.SHIFT_F2,
    "Shift+F3": VKey.SHIFT_F3, "Shift+F4": VKey.SHIFT_F4,
    "Shift+F5": VKey.SHIFT_F5, "Shift+F6": VKey.SHIFT_F6,
    "Shift+F7": VKey.SHIFT_F7, "Shift+F8": VKey.SHIFT_F8,
    "Shift+F9": VKey.SHIFT_F9,
    # Ctrl combinations
    "Ctrl+F": VKey.CTRL_F, "Ctrl+G": VKey.CTRL_G, "Ctrl+P": VKey.CTRL_P,
}


def _parse_key(key: str) -> int:
    """Parse key name to VKey code."""
    if key not in _KEY_MAP:
        raise ValueError(
            f"Unknown key: '{key}'. Valid keys: {', '.join(_KEY_MAP.keys())}"
        )
    return _KEY_MAP[key]


# Keys that trigger a save confirmation via elicitation
_SAVE_KEYS = {"F11", "Save"}


async def _elicit_save(ctx: Context, message: str,
                       tool: str = "sap_send_key") -> tuple[str, str] | None:
    """Ask the user to confirm a save.

    Returns ``None`` if a person accepted (caller should proceed), else the
    action (``decline`` / ``cancel``) and a note that is empty unless the
    client answered without showing a dialog (confirmation.ask_user).
    Raises ``ValueError`` if the client does not support elicitation.
    """
    try:
        # Shown as a checkbox (CLI): tick it and Accept to save.
        answer = await ask_user(ctx, points=SAVE_POINT, tool=tool,
                                message=message, response_title="Yes, save")
    except McpError as exc:
        log_confirmation_event(SAVE_POINT, tool, "unsupported_client")
        raise ValueError(
            f"Save action requires confirmation but the client does not "
            f"support elicitation: {exc}"
        ) from exc

    if answer.yes:
        return None  # proceed
    return answer.action, answer.note


async def _confirm_save(ctx: Context, key: str) -> dict | None:
    """Request user confirmation before sending a save key.

    Returns ``None`` if the user accepted (caller should proceed).
    Returns a cancellation dict if declined/cancelled.
    Raises ``ValueError`` if the client does not support elicitation.
    """
    declined = await _elicit_save(ctx, (
        f"You are about to send '{key}' which triggers Save (F11) "
        "in SAP. This can persist changes to the database. "
        "Do you want to proceed?"
    ))
    if declined is None:
        return None
    action, note = declined
    return {
        "status": "cancelled",
        "action": action,
        "key": key,
        "reason": f"User did not confirm save ({action})",
        # Same wording as a decline at a confirmation point, which the
        # instructions tell agents to stop on.
        "error": f"Save declined by user ({action}). Nothing was sent.{note}",
    }


async def _confirm_popup_save(ctx: Context, plan: dict,
                              tool: str = "sap_handle_popup") -> dict | None:
    """The Save key's confirmation, for a press on a save-shaped popup or
    an answer to a popup question the tool cannot read.

    Returns ``None`` if the user accepted, else the popup read with
    ``status: cancelled`` (nothing was pressed).
    """
    popup = plan["popup"]
    btn = plan.get("button") or {}
    label = btn.get("text") or btn.get("tooltip") or "?"
    title = popup.get("title") or popup.get("window_id", "")
    question = " ".join(popup.get("texts", []))[:200]
    if plan.get("unread_question"):
        message = (
            f"You are about to answer the SAP popup '{title}' with '{label}'. "
            "Its question is shown in a browser control the tool cannot "
            "read, so this answer may save, delete or discard data. Check "
            "the question in the SAP window. Do you want to proceed?"
        )
    else:
        message = (
            f"You are about to press '{label}' on the SAP popup '{title}'"
            + (f" ({question})" if question else "")
            + ", which may save changes to the database. "
            "Do you want to proceed?"
        )
    declined = await _elicit_save(ctx, message, tool)
    if declined is None:
        return None
    action, note = declined
    # The popup is still open and unchanged: keep its buttons for the next
    # step (e.g. press "No").
    what = "the answer" if plan.get("unread_question") else "save"
    popup.update({
        "status": "cancelled",
        "action": "none",
        "button": label,
        "reason": f"User did not confirm {what} ({action})",
        "error": (f"Popup answer declined by user ({action}). "
                  f"Nothing was pressed.{note}"),
    })
    return popup


def _to_dict(obj):
    """Convert dataclass to dict, pass dicts through unchanged."""
    return obj.__dict__ if hasattr(obj, '__dict__') and not isinstance(obj, dict) else obj


# ===========================================================================
# Connection tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_connect(
    system_description: str,
    ctx: Context,
    client: str | None = None,
    user: str | None = None,
    language: str | None = None,
) -> dict:
    """Connect to an SAP system by its name in SAP Logon Pad.

    Credentials are resolved from a .env file (SAP_USER, SAP_PASSWORD,
    SAP_CLIENT, SAP_LANGUAGE) so passwords never appear in MCP tool calls.
    Parameters provided here override the .env values (except password).
    If SAP is already open and logged in, use sap_connect_existing instead."""
    c = _ctrl(ctx)
    kwargs: dict[str, str] = {"system_description": system_description}
    # Resolve credentials: explicit param > env var > omit
    kwargs["user"] = user or os.environ.get("SAP_USER") or None
    kwargs["client"] = client or os.environ.get("SAP_CLIENT") or None
    kwargs["language"] = language or os.environ.get("SAP_LANGUAGE") or None
    kwargs["password"] = os.environ.get("SAP_PASSWORD") or None
    # Strip None values so controller uses its own defaults
    kwargs = {k: v for k, v in kwargs.items() if v is not None}
    return _to_dict(await _com(lambda: c.connect(**kwargs)))


# "DEV(2)/100 " in front of a window title: the system, session number and
# client, which the attach dialog shows in their own columns already.
_TITLE_PREFIX_RE = re.compile(r"^\S+\(\d+\)/\d+\s+")
_SYSTEM_ID_RE = re.compile(r"^[A-Z0-9]{3}$")
_CLIENT_RE = re.compile(r"^\d{3}$")


def _parse_system(system: str) -> tuple[str, str | None]:
    """'DEV' or 'DEV/100' -> ('DEV', None) or ('DEV', '100')."""
    sid, _, client = str(system).strip().upper().partition("/")
    sid, client = sid.strip(), client.strip()
    if not _SYSTEM_ID_RE.fullmatch(sid) or (client and not _CLIENT_RE.fullmatch(client)):
        raise ValueError(
            f"system={system!r}: expected a system ID such as 'DEV', or "
            "system and client such as 'DEV/100'"
        )
    return sid, client or None


def _attach_candidates(connections: list) -> list[dict]:
    """Every open session that can be attached, with its connection index."""
    candidates = []
    for conn in connections:
        if conn.get("scripting_disabled"):
            continue
        for sess in conn.get("sessions", []):
            if "error" in sess or not sess.get("system"):
                continue
            candidates.append({**sess, "connection_index": conn["index"]})
    return candidates


def _attach_label(sess: dict) -> str:
    """'DEV (2) / 100 / TESTUSER - SE16 - Data Browser: Initial Screen'."""
    number = sess.get("session_number")
    head = f"{sess.get('system', '')}" + (f" ({number})" if number else "")
    title = _TITLE_PREFIX_RE.sub("", str(sess.get("title") or "")).strip()
    parts = [f"{head} / {sess.get('client', '')} / {sess.get('user', '')}",
             str(sess.get("transaction") or ""), title]
    return " - ".join(part for part in parts if part)


def _candidate_summary(candidates: list[dict]) -> str:
    return "; ".join(
        f"{_attach_label(s)} (connection_index={s['connection_index']}, "
        f"session_index={s['index']})"
        for s in candidates
    ) or "no session that can be attached"


_NOT_AN_OPTION = (
    "The answer to the session picker was not one of the listed sessions. "
    "Nothing was attached. Ask the user which system to use, then call "
    "sap_connect_existing(system='SID') (or system='SID/client')."
)


async def _pick_session(ctx: Context, candidates: list[dict],
                        message: str) -> dict:
    """Ask the user which session to attach to (the 'attach' point).

    Raises ValueError when the user declines, gives an answer that is not
    one of the sessions, or when the client cannot show the dialog (fail
    closed: the agent asks in the chat instead).

    The options go out as titled values (oneOf const/title): the user sees
    the label, the answer is the session's "connection/session" index pair.
    A plain enum of labels was shown as a free-text box by the Claude Code
    CLI (2.1.302), so only an exact retyped label could match.
    """
    options: dict[str, dict[str, str]] = {}
    titles: set[str] = set()
    for sess in candidates:
        title = _attach_label(sess)
        if title in titles:  # same system and number on two connections
            title += f" [connection {sess['connection_index']}]"
        titles.add(title)
        options[f"{sess['connection_index']}/{sess['index']}"] = {"title": title}
    def is_option(choice) -> bool:
        return choice in options

    try:
        # A client that shows the options as a text box lets the user type
        # anything, and fastmcp passes it through for titled options: an
        # accepted answer that is no option is refused, not taken as a no.
        answer = await ask_user(
            ctx, points=ATTACH_POINT, tool="sap_connect_existing",
            message=message, response_type=options,
            response_title="SAP session", is_yes=is_option, is_valid=is_option,
        )
    except McpError as exc:
        log_confirmation_event(ATTACH_POINT, "sap_connect_existing",
                               "unsupported_client")
        raise ValueError(
            "The 'attach' confirmation point is on, but this client cannot "
            "show the session picker. Ask the user in the chat which SAP "
            "system to use, then call sap_connect_existing(system='SID') "
            "(or system='SID/client')."
        ) from exc
    except Exception as exc:  # an answer the schema rejects (pydantic)
        log_confirmation_event(ATTACH_POINT, "sap_connect_existing",
                               "invalid_answer")
        raise ValueError(_NOT_AN_OPTION) from exc
    if answer.action == "invalid":
        raise ValueError(_NOT_AN_OPTION)
    if not answer.yes:
        raise ValueError(
            f"Attach declined by user ({answer.action}). Nothing was "
            f"attached.{answer.note}"
        )
    return candidates[list(options).index(answer.data)]


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_connect_existing(
    ctx: Context,
    connection_index: int = 0,
    session_index: int = 0,
    system: str | None = None,
) -> dict:
    """Connect to an already open SAP session. Use this when SAP is already logged in.

    This is the most common starting point. connection_index=0 and session_index=0
    connect to the first open session. Also attaches to SAP Business Client /
    embedded SAP GUI sessions. Indexes come from sap_list_connections; a
    connection with scripting_disabled=true cannot be attached.

    system: pass it only when the user named the SAP system, as 'DEV' or
    'DEV/100' (system/client). One open session on it is attached; several,
    or none, return an error that lists what is open. The indexes are then
    ignored.

    With the 'attach' confirmation point on, the user picks the session in a
    dialog unless `system` names exactly one open session. The response
    says how it was chosen (`chosen_by`: named, dialog or index)."""
    c = _ctrl(ctx)
    attach_point = ATTACH_POINT in active_confirmation_points(ctx)

    if system is None and not attach_point:
        info = _to_dict(await _com(
            lambda: c.connect_to_existing_session(connection_index, session_index)
        ))
        session_id = await _com(lambda: getattr(getattr(c, "_session", None),
                                                "Id", None))
        chosen = {"system": info.get("system_name"), "client": info.get("client"),
                  "user": info.get("user"), "id": session_id}
        log_attach_event(chosen, "index")
        return {**info, "chosen_by": "index"}

    candidates = _attach_candidates(await _com(c.list_connections))
    chosen_by = "dialog"
    if system is not None:
        sid, client = _parse_system(system)
        matches = [s for s in candidates
                   if str(s.get("system", "")).upper() == sid
                   and (client is None or str(s.get("client", "")) == client)]
        if not matches:
            raise ValueError(
                f"No open session on {system}. Open sessions: "
                f"{_candidate_summary(candidates)}"
            )
        if len(matches) == 1:
            pick, chosen_by = matches[0], "named"
        elif attach_point:
            pick = await _pick_session(
                ctx, matches,
                f"Several sessions are open on {system}. Attach the agent to "
                "which one?",
            )
        else:
            raise ValueError(
                f"{len(matches)} sessions are open on {system}: "
                f"{_candidate_summary(matches)}. Ask the user which one, then "
                "call sap_connect_existing with its connection_index and "
                "session_index (without system)."
            )
    else:
        if not candidates:
            raise ValueError(
                "No SAP session is open that can be attached (none open, or "
                "SAP GUI Scripting is disabled on the server)."
            )
        pick = await _pick_session(
            ctx, candidates, "Attach the agent to which SAP session?",
        )

    # The pick is a position from the listing made before the dialog (or the
    # name lookup). If sessions opened or closed meanwhile, that position can
    # now hold another session: check what was attached before using it.
    # Attaching, the check and the detach are one COM call, so no other tool
    # call (a parallel one) can run on the session before it is checked.
    def attach_checked():
        info = _to_dict(c.connect_to_existing_session(pick["connection_index"],
                                                      pick["index"]))
        attached = {
            "id": getattr(getattr(c, "_session", None), "Id", None),
            "system": info.get("system_name"),
            "client": info.get("client"),
            "user": info.get("user"),
            "session_number": info.get("session_number"),
        }
        changed = [
            key for key in ("id", "system", "client", "user", "session_number")
            if pick.get(key) is not None and attached.get(key) is not None
            and str(pick[key]) != str(attached[key])
        ]
        if changed:
            c.disconnect()  # attached, not opened: nothing is closed
        return info, attached, changed

    info, attached, changed = await _com(attach_checked)
    if changed:
        raise ValueError(
            "The open SAP sessions changed while the session was being chosen: "
            f"{_attach_label(pick)} was picked, but that position now holds "
            f"{_attach_label({**pick, **attached})}. Nothing was attached; "
            "call sap_connect_existing again."
        )
    log_attach_event(attached, chosen_by)
    return {**info, "chosen_by": chosen_by}


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_list_connections(ctx: Context) -> dict:
    """List all open SAP connections and sessions.

    Covers SAP Logon plus SAP Business Client / embedded SAP GUI sessions
    (`host`: saplogon | sapguiserver). A connection with no sessions and
    `scripting_disabled: true` means the server parameter sapgui/user_scripting
    is off, so no session can be read. Each session has its system, client,
    user, transaction, `session_number` (the "(2)" in "DEV(2)") and window
    `title`."""
    c = _ctrl(ctx)
    connections = await _com(c.list_connections)
    return {"connections": connections}


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_session_info(ctx: Context) -> dict:
    """Get information about the current SAP session (system, client, user, transaction, screen).

    Also the SAP GUI release (sap_gui_version) and scripting_read_only: when
    true, the SAP server allows read-only scripting and write tools refuse."""
    c = _ctrl(ctx)
    return _to_dict(await _com(c.get_session_info))


# ===========================================================================
# Navigation tools
# ===========================================================================

@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_execute_transaction(tcode: str, ctx: Context) -> dict:
    """Execute (start/run) an SAP transaction code (tcode), e.g. MM03, VA01, SM30.

    Navigates to the transaction's initial screen. Always check the
    screen info in the response to understand what screen you landed on.
    Some transactions require /n prefix for SCWM (e.g., /n/SCWM/MON).
    A bare /n leaves the current transaction (unsaved data is lost).

    Subject to transaction blocklist/allowlist. Use sap_get_session_info
    to see the current transaction before navigating away."""
    _check_write(ctx)
    _enforce_transaction_policy(tcode)
    c = _ctrl(ctx)
    return await _com(lambda: c.execute_transaction(tcode))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_send_key(
    key: Literal[
        "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8",
        "F9", "F10", "F11", "F12", "Back", "Save", "Cancel",
        "Execute", "Refresh",
        "Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5",
        "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9",
        "Ctrl+F", "Ctrl+G", "Ctrl+P",
    ],
    ctx: Context,
) -> dict:
    """Press a keyboard key / function key in the SAP window (sendVKey).

    Common keys: Enter (confirm), F1 (Help), F3 (Back), F4 (search help /
    value help), F5 (Refresh), F8 (Execute), F11 (Save), F12 (Cancel/Escape).
    Also supports Shift+F1..F9 and Ctrl+F, Ctrl+G, Ctrl+P.

    A key that is not active on the screen is not sent: the response has an
    error and available_keys, the keys this screen accepts.

    F11 / Save requires user confirmation via elicitation before proceeding;
    the response then has confirmed_by_user: true. A declined save
    returns status 'cancelled' with an error."""
    _check_write(ctx)
    vkey = _parse_key(key)
    c = _ctrl(ctx)
    confirmed = False
    if key in _SAVE_KEYS:
        # A key the screen refuses is never sent: do not ask the user to
        # approve a save that cannot happen.
        refusal = await _com(lambda: c.check_vkey(vkey))
        if refusal is not None:
            return refusal
        cancellation = await _confirm_save(ctx, key)
        if cancellation is not None:
            return cancellation
        confirmed = True
    result = await _com(lambda: c.send_vkey(vkey))
    if confirmed:
        result["confirmed_by_user"] = True
    return result


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_screen_info(ctx: Context) -> dict:
    """Get current SAP screen info: transaction, program, screen number, window
    title, and status bar message (success/error/warning text).

    Reads from ``session.ActiveWindow`` so the response always reflects
    what the user sees.  The ``active_window`` field tells you which
    window is in focus (e.g. ``wnd[0]`` for the main screen, ``wnd[1]``
    for a popup).  The ``title`` comes from the active window.

    Every action tool (press_button, send_key, select_menu, etc.) returns
    this same screen info, so you always know when a popup appears.
    Use sap_get_popup_window for full popup content (texts, buttons).

    Present only when true: message_has_long_text (sap_press_button on
    'wnd[0]/sbar' opens it) and message_as_popup (the message also sits in a
    popup that must be closed first).

    Present only when set: subscreens, the sub-screens on the active window
    as "program:dynnr", outer first (program / screen_number name only the
    main screen, so an RF step, a tab page or a detail area that changes
    shows only here), and cursor, the ID of the field the cursor is in. A
    screen redisplayed without a message, e.g. for an empty required
    field, still moves the cursor to that field."""
    c = _ctrl(ctx)
    return await _com(c.get_screen_info)


# ===========================================================================
# Field tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_field(field_id: str, ctx: Context) -> dict:
    """Read the value of a field on the current SAP screen.

    Returns value, type, changeable status, and labels (left/right).
    For a dropdown, value is the selected entry's text and key its key.
    For a radio button or checkbox, value is its label and selected says
    whether it is selected.
    Use sap_get_screen_elements to discover field IDs on unknown screens."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_field(field_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_set_field(field_id: str, value: str, ctx: Context) -> dict:
    """Set (type/enter) a value into an input field on the current SAP screen.

    Works on GuiTextField and GuiCTextField input fields, and on dropdowns
    (GuiComboBox, ID prefix cmb): there value picks the entry by its key or
    its text, and the response reports the key and text selected. For
    filling in multiple fields of a form at once, use sap_set_batch_fields
    instead. After setting a field, you may need to press Enter to trigger
    validation.

    A value longer than the field (MaxLength) or a field that is not
    changeable is refused before writing: status 'rejected' and an error
    that says which, and the field keeps its content."""
    _check_write(ctx)
    _check_okcode_bypass(field_id, value)
    c = _ctrl(ctx)
    return await _com(lambda: c.set_field(field_id, value))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_press_button(button_id: str, ctx: Context) -> dict:
    """Press (click) a pushbutton or toolbar button on the current SAP screen.

    Returns screen info after the press so you can detect navigation or popups.
    Use sap_get_toolbar_buttons to discover toolbar button IDs.
    Use sap_get_screen_elements to find on-screen button IDs.
    When screen info has message_has_long_text, pressing 'wnd[0]/sbar' opens
    that long text (Diagnosis / Procedure), usually as a list popup: read it
    with sap_read_list(window_id='wnd[1]').

    A popup button is checked like sap_handle_popup's press: a press that
    may save, or any answer but Cancel to a popup question the tool cannot
    read, asks the user first (confirmed_by_user: true when accepted)."""
    _check_write(ctx)
    c = _ctrl(ctx)

    def plan_or_press():
        # One COM job, as before, unless the button answers a popup that
        # needs the user: then stop after the read.
        plan = c.plan_popup_button(button_id)
        if plan is not None:
            return plan, None
        return None, c.press_button(button_id)

    plan, result = await _com(plan_or_press)
    if result is not None:
        return result
    cancellation = await _confirm_popup_save(ctx, plan, "sap_press_button")
    if cancellation is not None:
        return cancellation

    def recheck_and_press():
        changed = c.popup_changed(plan["popup"])
        if changed is not None:
            return {"button_id": button_id, "error": changed}
        return c.press_button(button_id)

    result = await _com(recheck_and_press)
    if "error" not in result:
        result["confirmed_by_user"] = True
    return result


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_menu(menu_id: str, ctx: Context) -> dict:
    """Select a menu item from the menu bar or a submenu.

    Example: 'wnd[0]/mbar/menu[1]/menu[0]'.
    Use sap_get_screen_elements on 'wnd[0]/mbar' to discover menu structure.
    Returns screen info after selection so you can detect navigation."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_menu(menu_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_checkbox(checkbox_id: str, ctx: Context, selected: bool = True) -> dict:
    """Select or deselect a checkbox on the current SAP screen.

    Set selected=false to uncheck. Use sap_get_screen_elements with
    type_filter='GuiCheckBox' to find checkbox IDs."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_checkbox(checkbox_id, selected))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_radio_button(radio_id: str, ctx: Context) -> dict:
    """Select a radio button on the current SAP screen.

    Use sap_get_screen_elements with type_filter='GuiRadioButton'
    to find radio button IDs on the current screen."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_radio_button(radio_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_combobox_entry(combobox_id: str, key_or_value: str, ctx: Context) -> dict:
    """Select an entry in a combobox/dropdown by its key or display value text.

    Accepts either the technical key or the visible display text.
    Use sap_get_combobox_entries first to see all valid options.
    F4 opens nothing on a dropdown; sap_set_field selects entries the same
    way."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_combobox_entry(combobox_id, key_or_value))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_tab(tab_id: str, ctx: Context) -> dict:
    """Select a tab in a tab strip control.

    Returns screen info after selection (tab content changes).
    Tab IDs typically look like 'wnd[0]/usr/tabsTABSTRIP/tabpTAB01'."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_tab(tab_id))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_combobox_entries(combobox_id: str, ctx: Context) -> dict:
    """List all entries in a combobox/dropdown.

    Returns key-value pairs so you know which values are valid."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_combobox_entries(combobox_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_set_batch_fields(
    fields: dict,
    ctx: Context,
    validate: bool = False,
    skip_readonly: bool = False,
    verbose: bool = False,
    validate_partial: bool = False,
) -> dict:
    """Fill in multiple input fields at once (dict of field_id to value).

    Use this to fill a form or selection screen in one call — more
    efficient than repeated sap_set_field calls. Table-control cells work
    too: take the cell_id template of each column from
    sap_read_table(columns_only=true) and replace {row}. A dropdown (cmb)
    takes an entry's key or text, as in sap_set_field.

    The result gives counts; `results` lists only the fields that did NOT
    succeed (all of them with verbose=true). Check failed == 0: a field
    that was not set does not stop the others. With --strict-errors such a
    batch is returned as an error.

    Args:
        fields: Dict mapping field_id -> value.
        validate: Press Enter after setting fields and return status-bar
            feedback. Skipped when no fields were actually set, and when
            any field failed (`failed` > 0) unless validate_partial=true.
            `validation.performed` is false when Enter was not pressed or
            SAP refused it (then with `error`).
        skip_readonly: Silently skip fields whose element reports
            Changeable == False instead of counting them as failures.
        verbose: List every field in `results`, not only the failures.
        validate_partial: With validate=true, press Enter even when some
            fields failed.
    """
    _check_write(ctx)
    for fid, val in fields.items():
        _check_okcode_bypass(fid, str(val))
    c = _ctrl(ctx)
    return await _com(
        lambda: c.set_batch_fields(
            fields, skip_readonly=skip_readonly, validate=validate,
            verbose=verbose, validate_partial=validate_partial,
        )
    )


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_textedit(textedit_id: str, ctx: Context, max_lines: int = 0) -> dict:
    """Read the content of a multiline text editor (GuiTextedit) —
    long texts, notes, comments, document text.

    Returns full text and line count. Use max_lines to cap output for
    large text editors (0 = all lines)."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_textedit(textedit_id, max_lines))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_set_textedit(textedit_id: str, text: str, ctx: Context) -> dict:
    """Write text into a multiline text editor (GuiTextedit) —
    long texts, notes, comments, document text."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.set_textedit(textedit_id, text))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_set_focus(element_id: str, ctx: Context) -> dict:
    """Set focus to any screen element by its ID.

    Some SAP actions require focus on a specific element before they work
    (e.g., F4 search help on a field). Use this to set focus before
    sending keys with sap_send_key."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.set_focus(element_id))


# ===========================================================================
# Table tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_table(
    table_id: str,
    ctx: Context,
    max_rows: int = 100,
    columns: str = "",
    columns_only: bool = False,
    start_row: Optional[int] = None,
) -> dict:
    """Read rows from a table on the current screen — ALV grid (report/list
    output) or TableControl (SM30 maintenance view, customizing screens).

    Auto-detects the table type. The response includes a 'table_type' field
    ('GuiGridView' for ALV or 'GuiTableControl') so you know which
    type-specific tools to use next (e.g., sap_get_alv_toolbar for ALV,
    sap_scroll_table_control for TableControl).

    Use columns_only=true for schema discovery (returns column metadata
    only, no data). For a TableControl it also gives each column's cell_type
    and a cell_id template ({row} = zero-based visible row) for
    sap_set_batch_fields. Use columns to fetch only specific columns (CSV).
    Use start_row to paginate through large tables. On a TableControl a
    start_row (0 included) scrolls there first; without one the read starts
    where the table is scrolled now (e.g. after Position...).
    first_visible_row says which row is on top.

    A TableControl read returns at most the visible rows: page with
    start_row=next_start_row while has_more is true. start_row_note
    explains a clamped or failed scroll. total_rows counts its empty lines
    too (SM30 pads the view). When the read reaches the first empty row,
    the response has empty_rows_from, that row's absolute index: the data
    ends there."""
    c = _ctrl(ctx)
    capped = min(max_rows, config.max_table_rows)
    return _note_cap(await _com(lambda: c.read_table(
        table_id, capped, columns=columns,
        columns_only=columns_only, start_row=start_row,
    )), max_rows, capped)


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_alv_toolbar(grid_id: str, ctx: Context) -> dict:
    """Get all toolbar buttons from an ALV grid.

    Returns button IDs, texts, and types. Use this to discover available
    actions (sort, filter, export, etc.) before pressing them with
    sap_press_alv_toolbar_button. Only works on GuiGridView (ALV),
    not GuiTableControl."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_alv_toolbar(grid_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_press_alv_toolbar_button(grid_id: str, button_id: str, ctx: Context) -> dict:
    """Press a toolbar button on an ALV grid (e.g., sort, filter, export).

    Use sap_get_alv_toolbar to find button IDs. A menu button (Menu /
    ButtonAndMenu) opens its menu and the response lists menu_items (text and
    function_code); pick one with sap_select_alv_context_menu_item."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.press_alv_toolbar_button(grid_id, button_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_alv_context_menu_item(
    grid_id: str,
    menu_item_id: str,
    ctx: Context,
    toolbar_button_id: str | None = None,
    select_by: Literal["auto", "id", "text", "position"] = "auto",
) -> dict:
    """Select an item from an opened ALV context menu.

    First use sap_press_alv_toolbar_button on a Menu button to open it,
    then use this to select an item. Alternatively, pass toolbar_button_id
    to open the menu and select in one atomic call (recommended).

    `menu_item_id` can be a technical function code (for example `@M00006`),
    visible menu text (for example `Confirm WT in Foreground`), or a position
    descriptor. You can pass human-readable menu text directly. The
    menu_items of sap_press_alv_toolbar_button list what a menu offers.

    `select_by` controls how selection is performed:
    - `auto` (default): heuristic based on whether `menu_item_id` has spaces
    - `id`: function code
    - `text`: visible menu text
    - `position`: position descriptor
    """
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(
        lambda: c.select_alv_context_menu_item(
            grid_id, menu_item_id, toolbar_button_id, select_by
        )
    )


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_table_row(table_id: str, row: int, ctx: Context) -> dict:
    """Select a row in a table/grid.

    Works on both ALV grids and table controls. Row index is zero-based.
    For ALV: uses absolute row index. For TableControl: scrolls to make
    the row visible first if needed."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_table_row(table_id, row))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_double_click_cell(table_id: str, row: int, column: str, ctx: Context) -> dict:
    """Double-click a cell in a table/grid (often opens details or drills down).

    Row is zero-based. Column is the column name (from sap_read_table
    or sap_get_column_info). Works on both ALV and TableControl."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(
        lambda: c.double_click_table_cell(table_id, row, column)
    )


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_modify_cell(grid_id: str, row: int, column: str, value: str, ctx: Context) -> dict:
    """Modify the value of a cell in an ALV grid or table control.

    Only works on editable cells. Use sap_get_cell_info to check if
    a cell is changeable before attempting to modify it."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.modify_cell(grid_id, row, column, value))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_set_current_cell(grid_id: str, row: int, column: str, ctx: Context) -> dict:
    """Set the current (focused) cell in an ALV grid or table control.

    Useful before pressing toolbar buttons that act on the current cell."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.set_current_cell(grid_id, row, column))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_column_info(grid_id: str, ctx: Context) -> dict:
    """Get detailed column info from an ALV grid or table control.

    Returns column names, titles and tooltips; for an ALV grid also each
    column's data_type (string, date, decimal, ...) and key flag, for a table
    control each column's cell_type and cell_id template. For a lighter
    alternative, use sap_read_table with columns_only=true."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_column_info(grid_id))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_current_cell(table_id: str, ctx: Context) -> dict:
    """Get the currently focused cell position in an ALV grid or table control."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_current_cell(table_id))


# ---- TableControl-specific tools ----

@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_scroll_table_control(table_id: str, position: int, ctx: Context) -> dict:
    """Scroll a GuiTableControl to a specific row position.

    Does NOT work on ALV grids (they handle scrolling internally).
    For reading data at a specific offset, prefer sap_read_table with
    start_row parameter — it handles scrolling automatically.

    position in the response is the row SAP actually put at the top, which
    can differ from the one asked for (then requested_position and a note
    say so). Cell IDs ([col,row]) count visible rows from that row."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.scroll_table_control(table_id, position))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_table_control_row_info(
    table_id: str,
    ctx: Context,
    rows: list[int] | None = None,
) -> dict:
    """Get row metadata (selectable, selected) from a GuiTableControl.

    If rows is omitted, queries all visible rows.
    Does NOT work on ALV grids."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_table_control_row_info(table_id, rows))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_all_table_control_columns(
    table_id: str,
    ctx: Context,
    select: bool = True,
) -> dict:
    """Select or deselect all columns in a GuiTableControl. Does NOT work on ALV grids."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_all_table_control_columns(table_id, select))


# ---- ALV-specific tools ----

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_cell_info(grid_id: str, row: int, column: str, ctx: Context) -> dict:
    """Get detailed cell metadata from an ALV grid.

    Returns value, changeable, cell_type (Normal, CheckBox, ValueList,
    Button, ...), state (Normal, Error, Warning, Info: why a cell was
    rejected), f4_help, hotspot, color, tooltip, max_length; plus checked
    for a checkbox and list_box_count (entries) for a dropdown cell.
    Does NOT work on GuiTableControl."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_cell_info(grid_id, row, column))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_press_column_header(grid_id: str, column: str, ctx: Context) -> dict:
    """Click a column header in an ALV grid (triggers sort). Does NOT work on GuiTableControl."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.press_column_header(grid_id, column))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_all_rows(grid_id: str, ctx: Context) -> dict:
    """Select all rows in an ALV grid. Does NOT work on GuiTableControl."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_all_rows(grid_id))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_multiple_rows(table_id: str, rows: list[int], ctx: Context) -> dict:
    """Select multiple rows at once in an ALV grid or table control.

    Pass a list of row indices (e.g., [0, 2, 5])."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_multiple_rows(table_id, rows))


# ---- Popup & dialog tools ----

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_popup_window(ctx: Context) -> dict:
    """Check if a popup/modal dialog is open (wnd[1], wnd[2], etc.).

    Returns the popup's title, text content, and available buttons so
    you know how to respond. Also classifies the popup and suggests a
    safe next action. Radio buttons and checkboxes in interactive_elements
    carry selected. Returns {popup_exists: false} if no popup."""
    c = _ctrl(ctx)
    return await _com(c.get_popup_window)


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_handle_popup(
    ctx: Context,
    action: Literal["read", "confirm", "cancel", "press", "auto"] = "read",
    button_text: str = "",
) -> dict:
    """Read and optionally act on the current popup/modal dialog.

    Use when active_window shows wnd[1] or higher. Combines popup
    inspection and response in a single call.

    Actions:
    - read: return popup content without acting (default)
    - confirm: press OK/Yes/Continue/Enter on the popup
    - cancel: press Cancel (or F12) on the popup; never No, which on a
      save prompt drops the changes
    - press: press a specific button by its text or tooltip
    - auto: take only a clearly safe action; otherwise return the popup read-only

    A press that may save (a Save button, or Yes/OK on a popup about saving,
    e.g. "Save the changes first?") asks the user first, like F11 in
    sap_send_key. So does any answer but Cancel when the popup's question
    cannot be read (question_readable: false; SAP GUI's Edge browser
    control). On such a popup, confirm never falls back to Enter, and auto
    only reads. An accepted press reports confirmed_by_user: true.

    Returns the popup contents plus classification, requested action,
    and post-action screen/popup state when something was pressed."""
    _check_write(ctx)
    c = _ctrl(ctx)

    def plan_or_handle():
        # One COM job, as before, unless the press may save: then stop
        # after the read and ask the user before pressing.
        plan = c.plan_popup_action(action, button_text)
        if plan.get("save_risk"):
            return plan, None
        return plan, c.apply_popup_plan(plan)

    plan, result = await _com(plan_or_handle)
    if result is not None:
        return result
    cancellation = await _confirm_popup_save(ctx, plan)
    if cancellation is not None:
        return cancellation
    result = await _com(lambda: c.apply_popup_plan(plan, recheck=True))
    if result.get("action") != "error":
        result["confirmed_by_user"] = True
    return result


# ---- Toolbar discovery ----

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_toolbar_buttons(ctx: Context, window_id: str = "wnd[0]") -> dict:
    """List all buttons on the system toolbar (tbar[0]) and app toolbar (tbar[1]).

    Returns button IDs, text, tooltip, and enabled state. This is for
    standard SAP toolbars, NOT ALV (use sap_get_alv_toolbar for ALV)."""
    c = _ctrl(ctx)
    return await _com(lambda: c.get_toolbar_buttons(window_id))


# ---- Shell content ----

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_shell_content(shell_id: str, ctx: Context) -> dict:
    """Read content from a GuiShell subtype (e.g., HTMLViewer).

    An HTML viewer's page text and URL are readable only with the Internet
    Explorer browser control; with the Edge control (SAP GUI 7.70+ default,
    e.g. the F1 Performance Assistant) the response says so in `note`.
    Use sap_get_screen_elements first to find shell element IDs."""
    c = _ctrl(ctx)
    return await _com(lambda: c.read_shell_content(shell_id))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_list(
    ctx: Context,
    window_id: str = "wnd[0]",
    max_lines: int = 200,
    scroll_to: int = -1,
    with_ids: bool = False,
) -> dict:
    """Read a classic ABAP list (WRITE output of a report) as lines of text.

    Use it when a report shows a plain list rather than an ALV grid (program
    SAPMSSY0, screen 120), and for F4 hit lists in a popup (window_id=
    'wnd[1]'). Such a list has no table object: every word is its own label,
    so sap_get_screen_elements leaves the cells out and reports it under
    'lists'. This returns the page as the lines the user sees, with no
    screenshot needed.

    - colors: rows whose text has a semantic list colour, e.g. 'negative'
      (red, errors), 'positive' (green), 'total', 'heading'
    - scroll: present when the list is longer than the page; pass
      scroll_to=position+page_size to read the next page
    - with_ids=true adds the ID of each line's first label, to set the focus
      on a line (sap_set_focus) before F2 / double-click
    - is_list=false means this screen has no list: use sap_read_table"""
    c = _ctrl(ctx)
    return await _com(
        lambda: c.read_list(
            window_id, max_lines=max_lines, scroll_to=scroll_to, with_ids=with_ids,
        )
    )


# ===========================================================================
# Tree tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_read_tree(tree_id: str, ctx: Context, max_nodes: int = 200) -> dict:
    """Read data from a tree control (TableTreeControl, ColumnTreeControl, etc.).

    Returns node hierarchy with texts and column values. For large trees
    (e.g., SPRO with 1000+ nodes), use sap_get_tree_node_children for
    step-by-step navigation instead. Use sap_search_tree_nodes to find
    specific nodes by text.

    total_nodes is the number of nodes read; truncated=true means the tree
    has more loaded nodes (available_nodes) than max_nodes allowed."""
    c = _ctrl(ctx)
    capped = min(max_nodes, config.max_table_rows)
    return _note_cap(
        await _com(lambda: c.read_tree(tree_id, capped)), max_nodes, capped,
    )


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_expand_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Expand a folder node in a tree control to reveal its children.

    After expanding, use sap_get_tree_node_children or sap_read_tree
    to see the newly visible child nodes."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.expand_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_collapse_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Collapse a folder node in a tree control (e.g. SPRO/customizing tree)."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.collapse_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_select_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Select a node in a tree control.

    Highlights the node without opening it. For SPRO-style trees,
    use sap_click_tree_link on the execute icon column instead."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.select_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_double_click_tree_node(tree_id: str, node_key: str, ctx: Context) -> dict:
    """Double-click a node in a tree control (often opens details or drills down).

    In SPRO/customizing trees, this may open documentation (hypertext)
    rather than executing the activity. Use sap_click_tree_link on the
    execute column (typically column '2') for SPRO activities."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.double_click_tree_node(tree_id, node_key))


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_double_click_tree_item(
    tree_id: str, node_key: str, item_name: str, ctx: Context,
) -> dict:
    """Double-click a specific item (column cell) in a tree node row.

    item_name is the column name (e.g., 'Column1', 'Column2').
    Use sap_read_tree to discover column names for the tree."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(
        lambda: c.double_click_tree_item(tree_id, node_key, item_name)
    )


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_click_tree_link(tree_id: str, node_key: str, item_name: str, ctx: Context) -> dict:
    """Click a hyperlink in a tree node item.

    For SPRO/customizing trees, click on the execute icon column
    (typically item_name='2') to run an activity. Use sap_read_tree
    to see which columns have link-type items."""
    _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(
        lambda: c.click_tree_link(tree_id, node_key, item_name)
    )


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_find_tree_node_by_path(tree_id: str, path: str, ctx: Context) -> dict:
    """Find a tree node key by its path.

    E.g., '2\\1\\2' = 2nd child of root, then 1st child, then 2nd."""
    c = _ctrl(ctx)
    return await _com(lambda: c.find_tree_node_by_path(tree_id, path))


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_search_tree_nodes(tree_id: str, search_text: str, ctx: Context,
                                column: str = "", max_results: int = 20) -> dict:
    """Search for tree nodes by text. Returns matches with full ancestor paths.

    Useful for finding nodes in deep trees where the same label appears
    in multiple branches. Case-insensitive substring match.
    Optionally pass column to search in a specific column instead of node text.

    **Important limitation**: Only searches nodes that are already loaded
    (expanded) in the tree. Collapsed subtrees are not searched. If you
    don't find what you expect, expand parent nodes first using
    sap_get_tree_node_children with expand=true, then search again.

    total_matches counts the matches returned. stopped_early=true means the
    search reached max_results before the last loaded node, so more
    matches are possible."""
    c = _ctrl(ctx)
    capped = min(max_results, config.max_table_rows)
    return _note_cap(await _com(lambda: c.search_tree_nodes(
        tree_id, search_text, column, capped
    )), max_results, capped)


@mcp.tool(annotations=_WRITE, tags=_TAGS_WRITE)
async def sap_get_tree_node_children(tree_id: str, ctx: Context, node_key: str = "",
                                     expand: bool = False) -> dict:
    """Get direct children of a tree node. Much faster than read_tree for
    step-by-step navigation of deep trees (e.g., SPRO).

    Omit node_key or pass empty string for root-level nodes.
    Set expand=true to expand the node first (requires write permission)."""
    if expand:
        _check_write(ctx)
    c = _ctrl(ctx)
    return await _com(lambda: c.get_tree_node_children(
        tree_id, node_key, expand
    ))


# ===========================================================================
# Discovery tools
# ===========================================================================

_LIST_DISCOVERY_NOTE = (
    "Classic ABAP list: every word is its own cell, so the cells are left "
    "out. Read it with sap_read_list (window_id = the window in 'container'; "
    "with_ids=true gives each line's first cell). Any cell's ID is "
    "<container>/lbl[col,row] (chk[..] for a checkbox, txt[..] for an input "
    "field): row = first_row + line index, col = character offset in the "
    "line. expand_tables=true lists every cell."
)


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_screen_elements(
    ctx: Context,
    container_id: str = "wnd[0]/usr",
    max_depth: int = 2,
    type_filter: str = "",
    changeable_only: bool = False,
    expand_tables: bool = False,
) -> dict:
    """Discover all elements on the current SAP screen.

    Useful for finding field IDs when working with a new screen.

    Use type_filter and changeable_only to reduce response size on
    complex screens (e.g. type_filter="GuiTextField,GuiCTextField"
    to find only input fields).

    max_depth controls how deep to recurse into containers (default 2).
    Use max_depth=1 for a quick overview, max_depth=3+ for deeply nested
    layouts (splitter containers, tab strips with sub-containers).

    Pass container_id='wnd[0]/mbar' to discover the menu bar structure.

    Radio buttons and checkboxes report selected (true/false); a dropdown
    reports its selected key.

    A table control (SM30-style) is ONE element here, not one per cell: use
    sap_read_table(columns_only=true) for its columns and cell_id templates.
    A classic list (report WRITE output, SE16 standard list, F4 hit list) is
    left out and reported under 'lists': read it with sap_read_list.
    expand_tables=true lists every visible cell of both (large).

    Docking containers sit BESIDE the user area, not inside it: the
    dialog-structure tree of a view cluster (SM34, most IMG activities) and
    the SE80 tree are under 'wnd[0]/shellcont'. When the window has any, the
    response lists them as docking_containers; discover them with that id,
    or pass container_id='wnd[0]' (max_depth=1) for all top-level areas."""
    c = _ctrl(ctx)
    usr_window = re.search(r"(wnd\[\d+\])/usr$", container_id.strip())

    def _discover():
        lists: list = []
        found = c.get_screen_elements(
            container_id, max_depth=max_depth,
            type_filter=type_filter,
            changeable_only=changeable_only,
            expand_tables=expand_tables,
            lists=lists,
        )
        docking = c.get_docking_containers(usr_window.group(1)) if usr_window else []
        return found, docking, lists

    elements, docking, lists = await _com(_discover)
    result = {
        "element_count": len(elements),
        "elements": [
            {k: v for k, v in e.__dict__.items() if v is not None}
            for e in elements
        ],
    }
    if docking:
        result["docking_containers"] = docking
    if lists:
        result["lists"] = lists
        result["note"] = _LIST_DISCOVERY_NOTE
    return result


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_screenshot(
    ctx: Context,
    save_path: str = "",
    inline: bool = True,
) -> ToolResult:
    """Take a screenshot of the current SAP window.

    Use as a fallback when structured tools (sap_get_screen_elements,
    sap_read_field, sap_read_table) return empty or confusing results,
    e.g., on Web Dynpro screens where the element tree is non-standard.

    To put a screenshot into a document, pass save_path (must end in .png,
    directory must exist, existing files are never overwritten). The file
    keeps SAP GUI's full resolution; the inline image is capped at 1920px.
    The response then also reports the absolute path and pixel size. Set
    inline=False to skip the image when only the file is needed."""
    c = _ctrl(ctx)
    if not save_path.strip():
        result = await _com(c.take_screenshot)
        if "error" in result:
            raise ValueError(result["error"])
        return ToolResult(content=[
            ImageContent(type="image", data=result["data"], mimeType="image/png"),
        ])

    result = await _com(lambda: c.save_screenshot(save_path, inline=inline))
    if "error" in result:
        raise ValueError(result["error"])
    png_b64 = result.pop("data", None)
    content: list = [TextContent(type="text", text=json.dumps(result))]
    if png_b64:
        content.append(
            ImageContent(type="image", data=png_b64, mimeType="image/png")
        )
    return ToolResult(content=content)


# ===========================================================================
# Preview / approval card
# ===========================================================================

def _client_supports_ui(ctx: Context) -> bool:
    """Report whether the calling host negotiated the MCP Apps UI extension.

    Returns False outside a request context and for every SDK/in-memory
    client (the Python client cannot advertise the extension).
    """
    try:
        return bool(ctx.client_supports_extension(UI_EXTENSION_ID))
    except Exception:
        return False


_prefab_hint_logged = False


def _hint_missing_prefab_once() -> None:
    """Tell the operator once that the card would render if the extra existed."""
    global _prefab_hint_logged
    if not _prefab_hint_logged:
        _prefab_hint_logged = True
        logger.info(
            "client supports MCP Apps UI but prefab-ui is not installed; "
            "install the apps extra to enable the card"
        )


async def _preview_screenshot_b64(c) -> str | None:
    """Capture the current window as base64 PNG, or None on failure.

    A preview must never fail because the screenshot did; capture problems
    (no pillow, HardCopy error) degrade to a preview without the image.
    """
    try:
        result = await _com(c.take_screenshot)
    except Exception as exc:
        logger.warning("Preview screenshot unavailable: %s", exc)
        return None
    if not isinstance(result, dict) or result.get("error"):
        return None
    data = result.get("data")
    return data if isinstance(data, str) and data else None


# fastmcp stamps its internal prefab marker (_meta.ui = true) when app=True is
# used without prefab-ui installed, an untested shape for hosts that expect
# _meta.ui to be an object. Only claim app support when a card can be built.
_APP_KWARGS = {"app": True} if prefab_available() else {}


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ, **_APP_KWARGS)
async def sap_preview(
    ctx: Context,
    note: str = "",
    pending_fields: dict[str, str] | None = None,
    include_screenshot: bool = True,
) -> ToolResult:
    """Render a preview card of the current SAP screen for the user:
    screenshot, session/transaction info, status bar, and the values the
    agent is about to write (pending_fields). Call before significant
    writes (batch fills, F11 save) or whenever the user asks to see
    what's about to happen. Read-only: nothing is written or saved."""
    c = _ctrl(ctx)
    screen = await _com(c.get_screen_info)
    session = _to_dict(await _com(c.get_session_info))

    png_b64 = await _preview_screenshot_b64(c) if include_screenshot else None
    if not include_screenshot:
        screenshot_state = SCREENSHOT_OMITTED
    elif png_b64:
        screenshot_state = SCREENSHOT_INCLUDED
    else:
        screenshot_state = SCREENSHOT_UNAVAILABLE

    # Authored first and unconditionally: it is the result whenever the card
    # cannot be built. content must always be passed explicitly, else
    # ToolResult serializes the raw card dump as model-visible text.
    text = build_preview_text(
        screen=screen,
        session=session,
        note=note,
        pending_fields=pending_fields,
        screenshot=screenshot_state,
    )
    content: list = [TextContent(type="text", text=text)]

    wants_card = _client_supports_ui(ctx)
    if wants_card and not prefab_available():
        _hint_missing_prefab_once()
    if wants_card and prefab_available():
        try:
            card = build_preview_card(
                screen=screen,
                session=session,
                note=note,
                pending_fields=pending_fields,
                # Card copy only: sap_screenshot and the text branch's image
                # block keep full resolution.
                image_data_uri=(
                    "data:image/png;base64," + downscale_png_b64(png_b64)
                    if png_b64 else None
                ),
            )
        except Exception as exc:
            logger.warning("Preview card unavailable, falling back to text: %s", exc)
        else:
            # The card carries the image; a second copy as an ImageContent
            # block would double a ~300 KB payload.
            return ToolResult(content=content, structured_content=card)

    if png_b64:
        content.append(
            ImageContent(type="image", data=png_b64, mimeType="image/png")
        )
    return ToolResult(content=content)


# ===========================================================================
# Policy profile tools
# ===========================================================================

# _PROFILES is defined in the Configuration section above: ServerConfig
# validates its `profile` field against it at import time.


def _effective_profile(requested: str) -> tuple[str, set[str]]:
    """Cap a requested session profile at the server-level --profile floor.

    fastmcp session rules override global transforms, so a session could
    otherwise re-enable tools the server hid with ``mcp.disable(tags=...)``.
    Profiles are nested (exploration ⊂ operator ⊂ full), so intersecting the
    tag sets always lands exactly on one of them.
    """
    # ServerConfig.__post_init__ guarantees config.profile is a known profile.
    tags = _PROFILES[requested] & _PROFILES[config.profile]
    for name in ("exploration", "operator", "full"):
        if _PROFILES[name] == tags:
            return name, tags
    return requested, tags  # pragma: no cover — profiles are nested


@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_set_policy_profile(
    profile: Literal["exploration", "operator", "full"],
    ctx: Context,
) -> dict:
    """Switch the active policy profile for this session.

    Profiles control which tools are visible:
    - exploration: read-only tools (discover, inspect, screenshot)
    - operator: read + write tools (normal SAP interaction)
    - full: all tools including destructive (transaction execution)

    A session can only restrict, never widen: if the server was started with
    --profile, a request above that floor is capped to the floor.
    Default profile is 'full' unless the server was started with --profile."""
    effective, allowed_tags = _effective_profile(profile)
    # Subtractive semantics (same as the server-level --profile path): touch
    # only read/write/destructive-tagged tools. Untagged components — the
    # code-mode meta-tools (execute/search/get_schema/tags), prompts,
    # resources — must survive profile switches.
    all_tags = {"read", "write", "destructive"}
    await ctx.enable_components(tags=allowed_tags)
    disallowed = all_tags - allowed_tags
    if disallowed:
        await ctx.disable_components(tags=disallowed)
    if effective != profile:
        message = (
            f"Session profile '{profile}' capped to '{effective}' by the "
            f"server --profile floor"
        )
    else:
        message = f"Session switched to '{profile}' profile"
    return {
        "profile": effective,
        "requested_profile": profile,
        "server_profile_floor": config.profile,
        "enabled_tags": sorted(allowed_tags),
        "message": message,
    }


# ===========================================================================
# Confirmation point tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_set_confirmation_points(
    points: list[Literal["batch_fields", "field_writes", "transactions",
                         "all_writes", "attach"]],
    ctx: Context,
) -> dict:
    """Choose which SAP operations need the user's explicit approval first.

    Replaces this session's set of confirmation points. When a point is
    active, every matching tool call pauses and asks the user to approve it
    (MCP elicitation) before anything reaches SAP; declining returns an error
    and nothing is executed.

    Points:
    - transactions: sap_execute_transaction
    - batch_fields: sap_set_batch_fields
    - field_writes: sap_set_field, sap_modify_cell, sap_set_textedit,
      sap_select_checkbox, sap_select_radio_button, sap_select_combobox_entry
    - all_writes: every write-tagged tool. Strict and noisy — many write tools
      only select, scroll or navigate, and they will all prompt. Under
      all_writes, sap_send_key("F11") asks twice: once for the category and
      once for the unchanged save gate.
    - attach: sap_connect_existing lets the user pick the SAP session in a
      dialog, unless the call names a system with exactly one open session.

    'save' (F11 / Save key, Save toolbar button) is always on and cannot be
    turned off. Points set by the server's --confirm flag cannot be removed
    either. Adding a point is silent; *removing* one asks the user first, and
    a declined removal keeps the point.

    Confirmation prompts block until the user answers, and on clients without
    elicitation support the gated call fails instead of running unconfirmed
    (same fail-closed rule as the save gate). On such a client points also
    cannot be removed, so activating one there blocks that whole category for
    the rest of the session. Call sap_preview before a gated action so the
    user can see the screen and pending values first."""
    if _session_mgr is None:
        # Fail loudly: reporting success for points that are not stored
        # anywhere would be a silent safety failure.
        raise RuntimeError("Server not initialised – lifespan has not started")

    requested = set()
    for point in points:
        if point not in SETTABLE_POINTS:
            raise ValueError(
                f"Unknown confirmation point '{point}'. "
                f"Valid points: {', '.join(SETTABLE_POINTS)}"
            )
        requested.add(point)

    key = _session_key(ctx)
    floor = set(config.confirmation_floor)
    current = _session_mgr.get_confirmation_points(key)
    # Dropping a floor point (or 'save') is a no-op, not a relaxation, so it
    # never reaches the user.
    removals = sorted((current - requested) - floor)

    declined: list[str] = []
    note = ""
    if removals:
        try:
            answer = await ask_user(
                ctx,
                points=removals,
                tool="sap_set_confirmation_points",
                message=build_relax_prompt(removals),
                response_title="Yes, turn it off",
            )
        except McpError as exc:
            # Fail closed: the points stay active. Swallowing this would make
            # an unsupported client the easiest way to disable the gate.
            for point in removals:
                log_confirmation_event(point, "sap_set_confirmation_points",
                                       "unsupported_client")
            raise ValueError(
                f"Disabling a confirmation point requires confirmation but the "
                f"client does not support elicitation: {exc}"
            ) from exc

        if not answer.yes:
            declined = removals
            note = answer.note

    stored = (requested | set(declined)) - floor
    _session_mgr.set_confirmation_points(key, stored)

    active = sorted(effective_points(floor, stored))
    if declined:
        message = (
            f"User declined to disable: {', '.join(declined)}. "
            f"Active confirmation points: {', '.join(active)}"
            + (f".{note}" if note else "")
        )
    else:
        message = f"Active confirmation points: {', '.join(active)}"
    return {
        "points": points_with_provenance(floor, stored),
        "active": active,
        "removal_declined": declined,
        "always_on": [SAVE_POINT],
        "settable": list(SETTABLE_POINTS),
        "message": message,
    }


# ===========================================================================
# Workflow guidance tool
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_workflow_guide(
    workflow: WorkflowName,
    target: str,
) -> dict:
    """Return step-by-step guidance for a common SAP workflow.

    Available workflows:
    - **search_help**: F4 search help on a field. `target` = field ID.
    - **table_export**: Paginated table export. `target` = table element ID.
    - **spro_navigate**: SPRO customizing navigation. `target` = activity name.
    """
    guide_text = render_workflow_guide(workflow, target)
    return {
        "workflow": workflow,
        "target_parameter": WORKFLOW_TARGET_PARAMETERS[workflow],
        "target": target,
        "guide": guide_text,
    }


# ===========================================================================
# Transaction guidance tool
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_get_transaction_guide(
    transaction: str,
    task: str = "",
) -> dict:
    """Return a generic, read-first guide for a supported SAP transaction.

    Available transactions:
    - **/SCWM/MON**: EWM Warehouse Monitor with tree navigation and ALV results.
    - **SM30**: table/view maintenance and view clusters (SM34) - also what
      most IMG activities open: reading the table, the docked dialog-structure
      tree, new entries, saving and the customizing request prompt.

    Aliases accepted: `SCWM/MON`, `warehouse monitor`, `ewm warehouse monitor`,
    `SM34`, `table maintenance`, `view maintenance`, `view cluster`.
    """
    canonical = normalize_transaction(transaction)
    guide_text = render_transaction_guide(canonical, task)
    return {
        "transaction": canonical,
        "task": task,
        "mode": "read-first",
        "guide": guide_text,
    }


# ===========================================================================
# Session lifecycle tools
# ===========================================================================

@mcp.tool(annotations=_READ_ONLY, tags=_TAGS_READ)
async def sap_disconnect(ctx: Context) -> dict:
    """Disconnect from the current SAP session and release the binding.

    Sessions opened by this MCP session (via sap_connect) are closed.
    Sessions that were attached (via sap_connect_existing) are detached
    but left open so the user can continue working manually."""
    key = _session_key(ctx)
    return await _com(lambda: _session_mgr.release(key))


@mcp.resource("docs://sap-gui-guide")
def sap_gui_guide() -> str:
    """Detailed SAP GUI navigation reference for AI agents.

    Covers element types, ID patterns, transaction conventions,
    table type comparison, SPRO navigation, and troubleshooting."""
    return _SAP_GUI_GUIDE


# ===========================================================================
# Code mode (experimental, opt-in via --code-mode)
# ===========================================================================

_CODE_MODE_EXECUTE_DESCRIPTION = """\
Run a Python script that chains SAP GUI tools via `await call_tool(name, params)`.
Only the script's `return` value enters the conversation — use it for the final answer
(print output is discarded; every call_tool MUST be awaited).

SAP-specific rules:
- After every state-changing call (set_field, press_button, send_key,
  execute_transaction, cell/tree writes), check the returned
  screen.active_window: if it is not "wnd[0]", a popup appeared — call
  sap_get_popup_window and handle it before continuing.
- Read large tables in pages: loop sap_read_table with start_row until
  total_rows is reached (ALV) or has_more is false (TableControl, with
  start_row=next_start_row); aggregate locally and return only the aggregate.
- Never call sap_screenshot inside a script (image results are unusable in the
  sandbox) — call it directly as a normal tool instead.
- Sandbox limits: plain async code only; no classes; imports restricted to
  sys, os, typing, asyncio, re, datetime, json.
"""


def _build_code_mode_transform():
    """Construct the CodeMode transform with SAP-tuned settings.

    Kept as a separate function so tests can attach the identical transform
    to a test server. Imports lazily: fastmcp's code-mode extra
    (pydantic-monty) is an optional dependency group.
    """
    from fastmcp.experimental.transforms.code_mode import (
        CodeMode,
        GetSchemas,
        GetTags,
        MontySandboxProvider,
        Search,
    )

    return CodeMode(
        sandbox_provider=MontySandboxProvider(
            # Defaults (30 s / 50 calls) are too tight for table pagination
            # loops over slow SAP screens.
            limits={"max_duration_secs": 120, "max_memory": 100_000_000},
        ),
        discovery_tools=[GetTags(), Search(default_detail="detailed"), GetSchemas()],
        max_tool_calls=200,
        execute_description=_CODE_MODE_EXECUTE_DESCRIPTION,
    )


# ===========================================================================
# Main entry point
# ===========================================================================

def _object_tree_settings() -> dict:
    """The GetObjectTree switch and trace file, as read from the environment.

    Raises ValueError for an unknown site in MCP_SAP_GUI_DISABLE_OBJECT_TREE
    or a MCP_SAP_GUI_TRACE_FILE that cannot be opened for appending.
    """
    disabled = object_tree_disabled_sites()
    trace_file = os.environ.get(OBJECT_TREE_TRACE_ENV, "").strip() or None
    if trace_file:
        try:
            with open(trace_file, "a", encoding="utf-8"):
                pass
        except OSError as exc:
            raise ValueError(
                f"{OBJECT_TREE_TRACE_ENV}: cannot write {trace_file}: {exc}"
            ) from exc
    return {"disabled_sites": sorted(disabled), "trace_file": trace_file}


def main():
    """Main entry point."""
    import argparse
    global config

    # Load .env file (SAP_USER, SAP_PASSWORD, SAP_CLIENT, SAP_LANGUAGE)
    load_dotenv()

    parser = argparse.ArgumentParser(description="MCP Server for SAP GUI")
    parser.add_argument("--read-only", action="store_true",
                        help="Run in read-only mode (no write operations)")
    parser.add_argument("--allowed-transactions", nargs="*",
                        help="Whitelist of allowed transaction codes")
    parser.add_argument("--policy-preset", choices=sorted(PRESETS),
                        help="Transaction policy preset: 'default' blocks "
                             "basis, security and development transactions; "
                             "'abap-dev' additionally allows SA38, SE11 and "
                             "SE80; 'strict' blocks everything that the "
                             "policy file does not allow (default: the "
                             "preset named in the policy file, else 'default')")
    parser.add_argument("--policy-file", metavar="FILE",
                        help="JSON transaction policy (keys: preset, block, "
                             "allow; glob patterns such as SU*). Also read "
                             f"from ${POLICY_FILE_ENV}, else from "
                             "the per-user policy.json if it exists. Keep it "
                             "outside the agent's workspace")
    parser.add_argument("--transport", choices=["stdio", "http"],
                        default="stdio",
                        help="Transport mode (default: stdio)")
    parser.add_argument("--host", default="127.0.0.1",
                        help="HTTP bind address (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000,
                        help="HTTP port (default: 8000)")
    parser.add_argument("--profile", choices=["exploration", "operator", "full"],
                        default="full",
                        help="Default policy profile (default: full)")
    parser.add_argument("--confirm", nargs="+", metavar="POINT",
                        choices=list(SETTABLE_POINTS), default=[],
                        help="Require blocking user confirmation for these "
                             "operation categories in every session "
                             f"(choices: {', '.join(SETTABLE_POINTS)}). "
                             "Sessions can add points but never remove these")
    parser.add_argument("--strict-errors", action="store_true",
                        help="Return a failed tool result (an 'error' key, "
                             "or a batch with fields that were not set) as "
                             "an MCP error (isError true), keeping its "
                             "structured content. Also enabled by "
                             f"${STRICT_ERRORS_ENV}=1. Under --code-mode a "
                             "script's call_tool still returns the result "
                             "dict and does not raise: scripts check "
                             "'error' themselves")
    parser.add_argument("--audit-log", metavar="FILE",
                        help="Write audit log to FILE (JSON lines)")
    parser.add_argument("--code-mode", action="store_true",
                        help="EXPERIMENTAL: replace the tool catalog with "
                             "search/get_schema/tags/execute meta-tools; agents "
                             "script SAP flows in a sandbox (requires the "
                             "'code-mode' extra)")
    parser.add_argument("--debug", action="store_true",
                        help="Enable debug logging")
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    # Set up audit log file if requested
    if args.audit_log:
        audit_logger = logging.getLogger("mcp_sap_gui.audit")
        handler = logging.FileHandler(args.audit_log, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(message)s"))
        audit_logger.addHandler(handler)
        audit_logger.setLevel(logging.INFO)
        logger.info("Audit log: %s", args.audit_log)

    # Fail closed: a policy that cannot be read stops the server instead of
    # silently falling back to a more permissive default.
    try:
        policy = load_policy(
            _normalize_transaction_code,
            preset=args.policy_preset,
            policy_file=resolve_policy_file(args.policy_file),
        )
    except PolicyError as exc:
        parser.error(f"transaction policy: {exc}")

    # Same for the GetObjectTree switch: a mistyped site name would leave on
    # the very call it was meant to turn off.
    try:
        object_tree = _object_tree_settings()
    except ValueError as exc:
        parser.error(str(exc))

    config = ServerConfig(
        read_only=args.read_only,
        allowed_transactions=args.allowed_transactions,
        blocked_transactions=policy.block,
        allowed_exceptions=policy.allow,
        policy_preset=policy.preset,
        policy_source=policy.source,
        policy_file_sha256=policy.file_sha256,
        profile=args.profile,
        confirmation_floor=list(args.confirm or []),
        strict_errors=args.strict_errors or strict_errors_from_env(os.environ),
    )
    if config.strict_errors:
        logger.info("Strict errors: failed results are returned with isError")

    # The effective policy goes to the log and, as its own event, to the audit
    # log: a changed policy file is then visible after the fact.
    logger.info("Transaction policy: preset '%s' from %s (%d block, %d allow)",
                policy.preset, policy.source, len(policy.block), len(policy.allow))
    logging.getLogger("mcp_sap_gui.audit").info(
        json.dumps({"event": "policy", **config.policy.describe()})
    )
    if object_tree["disabled_sites"] or object_tree["trace_file"]:
        logger.info("GetObjectTree: disabled for %s (%s); trace file %s",
                    ", ".join(object_tree["disabled_sites"]) or "no site",
                    OBJECT_TREE_DISABLE_ENV, object_tree["trace_file"] or "off")
        logging.getLogger("mcp_sap_gui.audit").info(
            json.dumps({"event": "object_tree", **object_tree})
        )

    if config.confirmation_floor:
        logger.info("Confirmation floor: %s (plus always-on '%s')",
                    ", ".join(config.confirmation_floor), SAVE_POINT)

    # Apply server-level policy profile (hides tools globally)
    if args.profile != "full":
        allowed_tags = _PROFILES[args.profile]
        all_tags = {"read", "write", "destructive"}
        for tag in all_tags - allowed_tags:
            mcp.disable(tags={tag})
        logger.info("Server profile: %s (enabled tags: %s)",
                     args.profile, ", ".join(sorted(allowed_tags)))

    if args.code_mode:
        try:
            import pydantic_monty  # noqa: F401  # fail fast, not on first execute
        except ImportError:
            parser.error(
                "--code-mode requires the Monty sandbox. Install the extra: "
                "uv sync --extra code-mode  (or: pip install 'mcp-sap-gui[code-mode]')"
            )
        mcp.add_transform(_build_code_mode_transform())
        logger.info("Code mode: tool catalog replaced with meta-tools "
                    "(search/get_schema/tags/execute)")

    if args.transport == "http":
        logger.info("Starting HTTP transport on %s:%d", args.host, args.port)
        mcp.run(transport="http", host=args.host, port=args.port)
    else:
        mcp.run()


if __name__ == "__main__":
    main()
