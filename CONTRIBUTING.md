# Contributing to MCP SAP GUI Server

Thanks for contributing. The project's focus is safe local MCP use, reliable SAP GUI automation, and clear documentation.

## Scope

Useful contributions right now include:

- bug fixes with regression tests
- documentation improvements
- safer or more reliable SAP GUI workflows
- better MCP client setup guidance
- test coverage for controller and server behavior

Larger ideas (new tool families, changes to the safety gates or the transaction policy, new transports) are welcome, but they should usually start as an issue or design discussion before a broad implementation.

## Development Setup

```bash
uv sync --extra dev --extra screenshots --extra code-mode --extra apps
```

These are the extras CI installs; with fewer, the tests for the missing ones are skipped.

Real SAP GUI integration work requires:

- Windows
- SAP GUI for Windows
- SAP GUI Scripting enabled

The unit test suite is mocked and can run in CI without a live SAP system.

Working with a coding agent? [AGENTS.md](AGENTS.md) has the technical ground rules (layout, checks, safety gates, SAP GUI Scripting pitfalls).

## Before Opening a PR

- Keep changes focused. Small, reviewable PRs are easier to merge.
- Use generic SAP examples such as system `DEV`, client `100`, user `TESTUSER`, `MAT-001`, and placeholder business values.
- Do not commit customer-specific system names, screenshots, or SAP data.
- Do not commit local-only files from `.dev/` or other ignored workspace artifacts.
- If you change user-facing behavior, update the relevant docs in `README.md` or `docs/`.

## Validation

Run the checks relevant to your change before opening a PR:

```bash
uv run ruff check src tests
uv run pytest -q
python scripts/check_docs.py
uv build
```

CI runs them on Windows with Python 3.11 to 3.14.

Notes:

- `pytest` and `ruff` are the current baseline checks.
- `check_docs.py` validates local Markdown links and anchors.
- `mypy` is not yet a release gate because the project still has a large pre-existing error baseline.

## Code and Documentation Guidelines

- Add or update tests for bug fixes and behavior changes where practical.
- Preserve the default safety posture around blocked transactions and read-only mode.
- If you add, remove, or rename MCP tools, update `docs/TOOLS.md` and any README counts or examples that mention them.
- Keep examples client-agnostic where possible; Claude can be used as an example, but the server should remain usable from Codex, Copilot, Gemini CLI, and other MCP clients.
- Prefer clear errors over silent fallbacks when SAP GUI state is ambiguous.

## Pull Requests

Please include:

- a short description of what changed
- why the change is needed
- how you validated it
- any docs updates that were made

If a change touches SAP interaction behavior, mention whether it was validated only with mocks or also against a real SAP GUI session.

## Security and Privacy

- For vulnerability reports, use the process in `SECURITY.md` rather than opening a public issue.
- Never include credentials, connection details, or real business data in commits, issues, or screenshots.
- Be careful with SAP screenshots: they often contain sensitive information.
- Changes that weaken the current safety checks should come with a strong justification.

By contributing, you agree that your contributions will be licensed under the project's MIT license.
