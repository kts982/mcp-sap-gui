# MCP SAP GUI Overview

## What This Project Is

`mcp-sap-gui` is a local MCP server for SAP GUI for Windows.

It gives MCP-capable assistants access to SAP GUI through the SAP GUI Scripting API, so they can:

- inspect the current SAP screen
- discover element IDs instead of guessing
- read and write fields
- work with tables, ALV grids, and trees
- handle popups
- capture screenshots

This project is intentionally focused on **real SAP GUI automation**, not browser automation, OCR, or generic desktop clicking.

## Current Release Shape

The current public release target is **`0.1.0` alpha**.

That means:

- Windows only
- SAP GUI for Windows only
- MCP `stdio` (default) and `streamable HTTP` transports
- interactive MCP client usage only

The server currently exposes **55 MCP tools** plus built-in MCP instructions and a `docs://sap-gui-guide` resource.

## Why This Exists

SAP work is full of repetitive GUI-heavy flows:

- checking configuration
- navigating SPRO trees
- reading tables and grids
- entering data into structured screens
- capturing screenshots and evidence

This project reduces that work by giving an assistant a structured interface to SAP GUI rather than forcing it to reason from screenshots alone.

## What It Does Well

The current strengths are:

- strong SAP GUI coverage across fields, tables, ALV grids, trees, menus, and popups
- discovery-first workflows that reduce brittle hardcoded IDs
- safety controls: read-only mode, transaction blocklists/allowlists, and tag-based policy profiles
- built-in MCP guidance so clients start with better navigation patterns
- unit-tested server and controller behavior

## What It Is Not Trying To Be Yet

`0.1.0` is not yet:

- a remote multi-user SAP automation service
- a Fiori or browser automation product
- a central admin console for SAP governance
- a workflow engine for unattended enterprise execution

Those are roadmap directions, not current release promises.

## Near-Term Direction

The next meaningful improvements are not just “more raw tools.” The higher-value areas are:

- better client and transport support
- stronger policy and approval controls
- workflow-level tools on top of the low-level primitives
- semantic helpers that reduce SAP ID fragility
- audit and admin tooling

In short: keep the SAP depth, but improve product shape.

## Suggested Reading

- [README.md](../README.md) for install and first use
- [CLIENTS.md](CLIENTS.md) for client-specific MCP setup
- [TOOLS.md](TOOLS.md) for the full tool catalog
