# Security Policy

## Reporting a Vulnerability

If you believe you have found a security issue in `mcp-sap-gui`, please do not open a public GitHub issue.

Use one of these paths instead:

- open a private GitHub security advisory for the repository
- contact the maintainer directly if you already have a private contact channel

Please include:

- a short description of the issue
- affected versions or commit range, if known
- reproduction steps or proof of concept
- impact assessment
- any suggested mitigation

## Scope Notes

`mcp-sap-gui` is currently an alpha project focused on local Windows use over MCP `stdio`.

Important current constraints:

- do not expose the server to untrusted users
- prefer `sap_connect_existing` for already authenticated SAP sessions
- avoid sharing screenshots, traces, or logs that contain SAP business data
- use transaction allowlists and read-only mode when appropriate

## Disclosure

I will try to acknowledge valid reports quickly and coordinate a fix before public disclosure when practical.
