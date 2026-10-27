"""What a client sees in the initialize answer, through a real in-memory client."""

from fastmcp import Client

from mcp_sap_gui import __version__, server


async def test_server_info_names_this_package_and_its_version():
    async with Client(server.mcp) as client:
        info = client.initialize_result.serverInfo

    assert info.name == "mcp-sap-gui"
    # Not FastMCP's version, which it reports when version= is not passed.
    assert info.version == __version__
