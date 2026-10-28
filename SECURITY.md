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

`mcp-sap-gui` is an alpha project for local use on Windows. stdio, the default transport, has no network surface; the optional HTTP transport binds to 127.0.0.1 and refuses requests for other host names and from other web origins, but has no authentication of its own.

Important current constraints:

- do not expose the server, over HTTP or otherwise, to untrusted users
- prefer `sap_connect_existing` for already authenticated SAP sessions
- avoid sharing screenshots, traces, or logs that contain SAP business data
- use the transaction policy (an allowlist where you can), read-only mode and the audit log when appropriate

The safeguards and a recommended setup: [Security guide](docs/SECURITY_GUIDE.md).

## Disclosure

I will try to acknowledge valid reports quickly and coordinate a fix before public disclosure when practical.
