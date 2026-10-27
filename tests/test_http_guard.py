"""The HTTP transport refuses requests a web page could forge (DNS rebinding).

Built from the same settings main() passes to mcp.run(transport="http").
"""

import json

import pytest
from starlette.testclient import TestClient

from mcp_sap_gui import server

INITIALIZE = json.dumps({
    "jsonrpc": "2.0", "id": 1, "method": "initialize",
    "params": {"protocolVersion": "2025-06-18", "capabilities": {},
               "clientInfo": {"name": "test", "version": "0"}},
})
HEADERS = {"Content-Type": "application/json",
           "Accept": "application/json, text/event-stream"}


def _status(allowed_hosts=(), allowed_origins=(), **headers) -> int:
    app = server.mcp.http_app(
        **server.http_guard_kwargs(allowed_hosts, allowed_origins))
    with TestClient(app, base_url="http://127.0.0.1:8000") as client:
        response = client.post("/mcp", content=INITIALIZE,
                               headers={**HEADERS, **headers})
        return response.status_code


@pytest.mark.parametrize("headers", [
    {},                                        # an MCP client: no Origin
    {"Origin": "http://127.0.0.1:8000"},       # this server's own origin
])
def test_local_requests_pass(headers):
    assert _status(**headers) == 200


def test_foreign_origin_is_refused():
    assert _status(Origin="http://evil.example") == 403


def test_rebinding_host_is_refused():
    # A page on evil.example whose name now resolves to 127.0.0.1.
    assert _status(Host="evil.example:8000", Origin="http://evil.example:8000") == 421


def test_allowed_host_and_origin_pass():
    assert _status(["sap-box"], Host="sap-box:8000") == 200
    assert _status([], ["https://tools.example"], Origin="https://tools.example") == 200
