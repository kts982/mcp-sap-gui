# AGENTS.md

Guide for coding agents (and people) working on this repository. Process rules (scope, what a pull request should contain, security reports) are in [CONTRIBUTING.md](CONTRIBUTING.md); this file adds the technical ground rules. Keep personal notes in a gitignored place (`.dev/`, `.claude/`). Claude Code users can import this file from their own `CLAUDE.md` with `@AGENTS.md`.

## What this is

An MCP server (FastMCP 3.x, Python 3.10+) that drives SAP GUI for Windows through the SAP GUI Scripting API over COM (pywin32). It runs only on Windows; the test suite runs anywhere, with COM mocked.

## Layout

- `src/mcp_sap_gui/server.py`: the FastMCP app: the 60 tools, the server instructions, the command-line flags and the middleware.
- `src/mcp_sap_gui/sap_controller.py`: `SAPGUIController`, composed from the mixins in `fields.py`, `tables.py`, `trees.py` and `discovery.py` on top of `controller.py` (`SAPGUIControllerBase`: connecting, sessions, screen info, GetObjectTree reads).
- `session_manager.py`: the SAP session bound to each MCP session, and the single COM thread.
- `confirmation.py`: confirmation points and the dialog helpers. `policy.py`: the transaction policy (presets, policy file). `audit.py`: the JSON audit log. `strict.py`: `--strict-errors`. `preview.py`: `sap_preview`. `prompts.py`: workflow guides.
- `tests/`: pytest, no SAP GUI needed. `docs/`: user documentation; `docs/TOOLS.md` lists every tool.

## Setup and checks

The same commands CI runs (Windows, Python 3.10 to 3.14):

```bash
uv sync --extra dev --extra screenshots --extra code-mode --extra apps
uv run ruff check src tests
uv run pytest -q
python scripts/check_docs.py
```

## Ground rules

### Safety

- Never weaken a safety gate to make something work: the transaction policy, read-only mode, the Save confirmation (F11, the Save toolbar button, popup answers that may save) and the confirmation points. They fail closed: when a client cannot show the dialog, the call is refused.
- A tool that changes SAP state is registered with `annotations=_WRITE` and `tags=_TAGS_WRITE`; a read tool with `_READ_ONLY` and `_TAGS_READ`. The tags drive the policy profiles and the `all_writes` confirmation point.
- Every tool that can write the command field (`wnd[0]/tbar[0]/okcd`) checks the value against the transaction policy. Strip command prefixes such as `/n` with `str.removeprefix`, never `lstrip`.

### SAP GUI Scripting (each of these was seen live)

- Never request `Selected` through `GetObjectTree`: on SAP GUI 8.00 it ends SAP Logon. Read it over COM for each element; `_OBJECT_TREE_UNSAFE_PROPS` in `controller.py` enforces this.
- Setting a table control's `VerticalScrollbar.Position` re-renders it and invalidates COM references: find the element again with `findById` after scrolling.
- Never guess element IDs, in code or in prompts: discover them on the screen.

### Python and FastMCP

- Never `from . import server` in a module that `server.py` imports. Under `python -m mcp_sap_gui.server` the running module is `__main__`, so that import loads a second copy with empty globals, and gates silently stop working. Pass callables in at registration instead. `tests/test_stdio_gate.py` (a real stdio subprocess) guards this.

### Tests

- No test touches SAP GUI. `tests/conftest.py` mocks `win32com` and `pythoncom`; use the fakes in `tests/fakes.py` and `tests/field_fakes.py`, or patch `server._ctrl` and `server._com`.
- For behaviour across the MCP boundary (middleware, confirmation dialogs), use fastmcp's in-memory `Client(server.mcp, elicitation_handler=...)`.
- Fake dialogs answer at once. `conftest.py` sets `confirmation.NO_DIALOG_MS` to 0 so those answers count as a person's; `tests/test_no_dialog.py` tests the real threshold.

### Docs and data

- A user-visible change gets an entry under `## Unreleased` in `CHANGELOG.md`. Adding, removing or renaming a tool also updates `docs/TOOLS.md` and the "60 MCP tools" counts in `README.md`, `docs/OVERVIEW.md` and `docs/TOOLS.md`.
- No real system IDs, client numbers, user names, host names or customer object names in committed files: code, tests, docs and commit messages. Use `DEV`, `100` and `TESTUSER`, as the tests do.
- Do not bump versions or edit release metadata (`pyproject.toml` version, `server.json`, `__init__.py`); the maintainer releases.

### Live testing

- Only against a test or trial system, never production. Probe scripts and notes belong in `.dev/` (gitignored), not in the repository.
