"""What a client sees as the server's identity, through a real in-memory client."""

from fastmcp import Client

from mcp_sap_gui import __version__, server


async def test_server_info_names_this_package_and_its_version():
    # server_info covers both eras: the initialize answer on legacy protocol
    # versions, server/discover on 2026-07-28 (which has no initialize).
    async with Client(server.mcp) as client:
        info = client.server_info

    assert info.name == "mcp-sap-gui"
    # Not FastMCP's version, which it reports when version= is not passed.
    assert info.version == __version__
