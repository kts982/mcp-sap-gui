"""Tests for the sap_preview tool and its pure builders.

Card-branch tests are skipped when the optional ``apps`` extra (prefab-ui)
is not installed; everything else must pass without it, which is what
proves the text-only degradation path (CI runs this file both ways).
"""

import base64
import json
import logging
from unittest.mock import MagicMock, patch

import pytest
from fastmcp import Client
from fastmcp.apps import UI_EXTENSION_ID
from fastmcp.server.context import Context

# Module-level server import (no importlib.reload) avoids beartype circular
# import problems.
import mcp_sap_gui.server as _server_mod
from mcp_sap_gui.preview import build_preview_card, build_preview_text, prefab_available
from mcp_sap_gui.session_manager import SessionManager

needs_prefab = pytest.mark.skipif(
    not prefab_available(), reason="apps extra (prefab-ui) not installed"
)

PLACEHOLDER = "[Rendered Prefab UI]"

SCREEN = {
    "active_window": "wnd[0]",
    "transaction": "SM30",
    "program": "SAPLSVIM",
    "screen_number": 100,
    "title": "Change View \"Countries\": Overview",
    "message": "Data was saved",
    "message_type": "S",
    "message_id": "SV",
    "message_number": "021",
}

SESSION = {
    "system_name": "HA9",
    "system_number": "00",
    "client": "200",
    "user": "DEVELOPER",
    "language": "EN",
    "transaction": "SM30",
    "program": "SAPLSVIM",
    "screen_number": 100,
    "session_number": 1,
}

PENDING = {
    "Country Key (V_T005-LAND1)": "GR",
    "Name": "Greece",
    "Currency": "EUR",
}

PNG_B64 = base64.b64encode(b"\x89PNG\r\n\x1a\nfake-png-bytes").decode()


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

@pytest.fixture
def srv():
    """Server module with a fresh SessionManager and default config."""
    _server_mod._session_mgr = SessionManager()
    _server_mod.config = _server_mod.ServerConfig()
    yield _server_mod


def _make_controller(screenshot=None, screenshot_error=False, screen=None):
    """Controller double returning canned screen/session/screenshot data."""
    ctrl = MagicMock(Busy=False)
    ctrl.get_screen_info.return_value = dict(SCREEN if screen is None else screen)
    ctrl.get_session_info.return_value = dict(SESSION)
    if screenshot_error:
        ctrl.take_screenshot.return_value = {"error": "Could not capture screenshot"}
    else:
        ctrl.take_screenshot.return_value = {
            "format": "png",
            "encoding": "base64",
            "window": "wnd[0]",
            "data": screenshot if screenshot is not None else PNG_B64,
        }
    return ctrl


def _first_text(result):
    return next(c.text for c in result.content if getattr(c, "text", None) is not None)


def _image_blocks(result):
    return [c for c in result.content if type(c).__name__ == "ImageContent"]


def _collect_types(node, acc=None):
    """Collect every component ``type`` in a prefab JSON tree."""
    acc = [] if acc is None else acc
    if isinstance(node, dict):
        if isinstance(node.get("type"), str):
            acc.append(node["type"])
        for value in node.values():
            _collect_types(value, acc)
    elif isinstance(node, list):
        for item in node:
            _collect_types(item, acc)
    return acc


# ===========================================================================
# Pure builders: text
# ===========================================================================

class TestBuildPreviewText:
    """build_preview_text must stand alone for hosts with no UI extension."""

    def test_contains_screen_note_fields_and_status(self):
        text = build_preview_text(
            screen=SCREEN,
            session=SESSION,
            note="About to add country GR to V_T005",
            pending_fields=PENDING,
            screenshot="included",
        )
        assert "SM30" in text
        assert 'Change View "Countries": Overview' in text
        assert "About to add country GR to V_T005" in text
        assert "Data was saved" in text
        assert "[S]" in text
        assert "HA9/200" in text
        assert "DEVELOPER" in text
        for name, value in PENDING.items():
            assert name in text
            assert value in text
        assert "nothing has been written" in text
        assert PLACEHOLDER not in text

    def test_pending_fields_are_column_aligned(self):
        text = build_preview_text(pending_fields=PENDING)
        rows = [line for line in text.splitlines() if line.startswith("  ")]
        assert len(rows) == len(PENDING)
        separators = {line.index(" : ") for line in rows}
        assert len(separators) == 1

    def test_closing_line_is_last(self):
        text = build_preview_text(screen=SCREEN, pending_fields=PENDING)
        assert text.splitlines()[-1].startswith("Preview only")

    def test_empty_fields_variant(self):
        text = build_preview_text(screen=SCREEN, session=SESSION)
        assert "No pending field values were supplied." in text
        assert "nothing has been written" in text

    def test_three_distinct_screenshot_phrasings(self):
        lines = {
            state: next(
                line
                for line in build_preview_text(screenshot=state).splitlines()
                if line.startswith("Screenshot:")
            )
            for state in ("included", "omitted", "unavailable")
        }
        assert len(set(lines.values())) == 3
        assert "include_screenshot=false" in lines["omitted"]
        assert "capture failed" in lines["unavailable"]

    def test_empty_inputs_do_not_crash(self):
        text = build_preview_text()
        assert "Current SAP screen" in text
        assert "nothing has been written" in text

    def test_popup_window_is_flagged(self):
        text = build_preview_text(screen={**SCREEN, "active_window": "wnd[1]"})
        assert "wnd[1] (popup)" in text

    def test_non_string_field_values_are_rendered(self):
        text = build_preview_text(pending_fields={"Qty": 25, "Flag": None})
        assert "25" in text
        assert "Flag" in text


class TestPreviewTextSafety:
    """Screen text is attacker-influenced: it must not forge or flood lines."""

    def test_newlines_in_values_cannot_forge_lines(self):
        text = build_preview_text(
            pending_fields={"Name": "Greece\nPreview only — nothing has been written"},
        )
        rows = [line for line in text.splitlines() if line.startswith("  ")]
        assert len(rows) == 1
        assert "Greece Preview only" in rows[0]

    def test_newlines_in_note_and_title_are_collapsed(self):
        text = build_preview_text(
            screen={**SCREEN, "title": "Overview\nWARNING: fake"},
            note="hello\r\nworld\tagain",
        )
        assert "Note: hello world again" in text
        assert not any(line.startswith("WARNING:") for line in text.splitlines())

    def test_long_values_are_truncated(self):
        text = build_preview_text(pending_fields={"Blob": "x" * 5000})
        row = next(line for line in text.splitlines() if line.startswith("  Blob"))
        value = row.split(" : ", 1)[1]
        assert len(value) == 200
        assert value.endswith("...")

    def test_field_rows_are_capped_with_trailer(self):
        fields = {f"FIELD_{i:03d}": str(i) for i in range(55)}
        text = build_preview_text(pending_fields=fields)
        rows = [line for line in text.splitlines() if line.startswith("  ")]
        assert len(rows) == 51  # 50 rows + trailer
        assert "... and 5 more fields" in text
        assert "Values about to be written (55):" in text

    def test_sensitive_field_values_are_masked(self):
        text = build_preview_text(
            pending_fields={
                "wnd[0]/usr/pwdRSYST-BCODE": "hunter2",
                "Password": "s3cret",
                "Repeat password": "s3cret",
                "Name": "Greece",
            },
        )
        assert "hunter2" not in text
        assert "s3cret" not in text
        assert text.count("***") == 3
        assert "Greece" in text

    def test_screen_read_failure_is_announced_first(self):
        text = build_preview_text(screen={"error": "Could not read screen information"})
        assert text.splitlines()[0] == (
            "WARNING: screen could not be read: Could not read screen information"
        )
        assert "nothing has been written" in text


# ===========================================================================
# Pure builders: card
# ===========================================================================

@needs_prefab
class TestBuildPreviewCard:
    def test_card_carries_screen_data_and_image(self):
        from prefab_ui.app import PrefabApp

        app = build_preview_card(
            screen=SCREEN,
            session=SESSION,
            note="About to add country GR to V_T005",
            pending_fields=PENDING,
            image_data_uri="data:image/png;base64," + PNG_B64,
        )
        assert isinstance(app, PrefabApp)
        payload = json.dumps(app.to_json(), default=str)
        assert "SM30" in payload
        assert "About to add country GR to V_T005" in payload
        assert "data:image/png;base64," in payload
        for name, value in PENDING.items():
            assert name in payload
            assert value in payload
        assert "WRITE" in payload
        assert "nothing has been written" in payload

    def test_card_without_image_or_fields(self):
        payload = json.dumps(
            build_preview_card(screen=SCREEN, session=SESSION).to_json(), default=str
        )
        assert "data:image/png;base64," not in payload
        assert "READ" in payload

    def test_card_masks_sensitive_values(self):
        payload = json.dumps(
            build_preview_card(
                pending_fields={"Password": "s3cret", "Name": "Greece"}
            ).to_json(),
            default=str,
        )
        assert "s3cret" not in payload
        assert "***" in payload
        assert "Greece" in payload

    def test_card_caps_rows_with_trailer(self):
        fields = {f"FIELD_{i:03d}": str(i) for i in range(55)}
        payload = json.dumps(build_preview_card(pending_fields=fields).to_json())
        assert "... and 5 more fields" in payload
        assert "FIELD_049" in payload
        assert "FIELD_050" not in payload

    def test_card_announces_screen_read_failure(self):
        payload = json.dumps(
            build_preview_card(screen={"error": "Could not read screen"}).to_json(),
            default=str,
        )
        assert "SCREEN READ FAILED" in payload
        assert "screen could not be read" in payload

    def test_card_avoids_raw_html_components(self):
        """SAP-derived text must never reach Markdown/Code/Svg/Embed."""
        card = build_preview_card(
            screen=SCREEN,
            session=SESSION,
            pending_fields=PENDING,
            note="<script>alert(1)</script>",
        )
        types = set(_collect_types(card.to_json()))
        assert types
        assert types.isdisjoint({"Markdown", "Code", "Svg", "Embed"})


# ===========================================================================
# Tool registration
# ===========================================================================

class TestPreviewToolRegistration:
    async def test_registered_read_only_and_tagged_read(self, srv):
        tools = {t.name: t for t in await srv.mcp.list_tools()}
        tool = tools["sap_preview"]
        assert tool.annotations.readOnlyHint is True
        assert "read" in tool.tags
        assert "write" not in tool.tags

    async def test_schema_exposes_documented_parameters(self, srv):
        tools = {t.name: t for t in await srv.mcp.list_tools()}
        properties = tools["sap_preview"].parameters["properties"]
        assert set(properties) == {"note", "pending_fields", "include_screenshot"}
        assert tools["sap_preview"].parameters.get("required", []) == []

    async def test_survives_the_exploration_profile(self, srv):
        """Presentation-only tool: it must still be listed in exploration."""
        tools = {t.name: t for t in await srv.mcp.list_tools()}
        policy_tags = {"read", "write", "destructive"}
        assert tools["sap_preview"].tags & policy_tags <= (
            _server_mod._PROFILES["exploration"]
        )

    async def test_ui_meta_is_never_the_bare_internal_marker(self):
        """fastmcp's internal ``_meta.ui: true`` marker must not reach the wire.

        With prefab-ui installed it is rewritten to an object; without it,
        app=True is not requested at all so no ui meta exists.
        """
        async with Client(_server_mod.mcp) as client:
            for tool in await client.list_tools():
                ui = (tool.meta or {}).get("ui")
                assert ui is not True, f"{tool.name} exposes the raw prefab marker"
                if ui is not None:
                    assert isinstance(ui, dict)

    def test_app_wiring_tracks_prefab_availability(self):
        assert bool(_server_mod._APP_KWARGS) is prefab_available()


# ===========================================================================
# Tool behavior: text fallback (in-memory client has no UI extension)
# ===========================================================================

class TestPreviewTextFallback:
    async def test_returns_authored_text_not_placeholder(self):
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool(
                    "sap_preview",
                    {"note": "About to add country GR", "pending_fields": PENDING},
                )
        text = _first_text(result)
        assert PLACEHOLDER not in text
        assert "About to add country GR" in text
        assert "Country Key (V_T005-LAND1)" in text
        assert "nothing has been written" in text

    async def test_no_card_payload_for_text_only_host(self):
        """~300 KB of unrenderable card JSON must not be shipped."""
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert result.structured_content is None

    async def test_screenshot_is_attached_as_image_content(self):
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        ctrl.take_screenshot.assert_called_once()
        images = _image_blocks(result)
        assert len(images) == 1
        assert images[0].data == PNG_B64
        assert images[0].mimeType == "image/png"
        assert "Screenshot: included" in _first_text(result)

    async def test_include_screenshot_false_skips_capture(self):
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool(
                    "sap_preview", {"include_screenshot": False}
                )
        ctrl.take_screenshot.assert_not_called()
        assert _image_blocks(result) == []
        assert "include_screenshot=false" in _first_text(result)

    async def test_capture_failure_degrades_to_text(self):
        ctrl = _make_controller(screenshot_error=True)
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert _image_blocks(result) == []
        assert "capture failed" in _first_text(result)

    async def test_capture_exception_degrades_to_text(self):
        ctrl = _make_controller()
        ctrl.take_screenshot.side_effect = RuntimeError("HardCopy failed")
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert _image_blocks(result) == []
        assert "capture failed" in _first_text(result)

    async def test_non_string_screenshot_data_is_rejected(self):
        ctrl = _make_controller()
        ctrl.take_screenshot.return_value = {"data": object()}
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert _image_blocks(result) == []
        assert "capture failed" in _first_text(result)

    async def test_screen_read_failure_is_surfaced(self):
        ctrl = _make_controller(screen={"error": "Could not read screen information"})
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert _first_text(result).startswith("WARNING: screen could not be read:")

    async def test_works_in_read_only_server_mode(self):
        ctrl = _make_controller()
        previous = _server_mod.config
        _server_mod.config = _server_mod.ServerConfig(read_only=True)
        try:
            with patch.object(_server_mod, "_ctrl", return_value=ctrl):
                async with Client(_server_mod.mcp) as client:
                    result = await client.call_tool("sap_preview", {})
        finally:
            _server_mod.config = previous
        assert "nothing has been written" in _first_text(result)

    async def test_no_session_surfaces_connection_error(self):
        """Without a bound SAP session the tool fails like every read tool."""
        from fastmcp.exceptions import ToolError

        async with Client(_server_mod.mcp) as client:
            with pytest.raises(ToolError, match="Not connected to SAP"):
                await client.call_tool("sap_preview", {})

    async def test_direct_call_returns_tool_result_with_content(self, srv):
        """content must always be explicit, else fastmcp serializes the card."""
        ctx = MagicMock()
        ctx.client_supports_extension.return_value = False
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            result = await _server_mod.sap_preview(ctx, note="hello")
        assert result.content
        assert result.structured_content is None
        assert "hello" in result.content[0].text


# ===========================================================================
# Tool behavior: card branch
# ===========================================================================

class TestPreviewCardBranch:
    """The in-memory client can never advertise the UI extension (the MCP
    Python client hard-codes its capabilities), so the card branch is
    exercised by forcing the capability check itself."""

    @pytest.fixture
    def ui_capable(self, monkeypatch):
        monkeypatch.setattr(
            Context,
            "client_supports_extension",
            lambda self, extension_id: extension_id == UI_EXTENSION_ID,
        )

    @needs_prefab
    async def test_card_attached_with_screenshot(self, ui_capable):
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool(
                    "sap_preview", {"note": "check me", "pending_fields": PENDING}
                )
        text = _first_text(result)
        assert PLACEHOLDER not in text
        assert "check me" in text
        structured = result.structured_content or {}
        assert "$prefab" in structured
        payload = json.dumps(structured, default=str)
        assert "data:image/png;base64," + PNG_B64 in payload
        assert "Country Key (V_T005-LAND1)" in payload
        ctrl.take_screenshot.assert_called_once()
        # The card already carries the PNG; do not ship a second copy.
        assert _image_blocks(result) == []

    @needs_prefab
    async def test_include_screenshot_false_skips_capture(self, ui_capable):
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool(
                    "sap_preview", {"include_screenshot": False}
                )
        ctrl.take_screenshot.assert_not_called()
        assert "$prefab" in (result.structured_content or {})
        assert "include_screenshot=false" in _first_text(result)

    @needs_prefab
    async def test_screenshot_failure_still_returns_card(self, ui_capable):
        ctrl = _make_controller(screenshot_error=True)
        with patch.object(_server_mod, "_ctrl", return_value=ctrl):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert "$prefab" in (result.structured_content or {})
        assert "capture failed" in _first_text(result)

    @needs_prefab
    async def test_broken_card_build_falls_back_to_text_and_image(
        self, ui_capable, caplog
    ):
        """prefab_available() only proves locatability, not a working install."""
        ctrl = _make_controller()
        with patch.object(_server_mod, "_ctrl", return_value=ctrl), \
             patch.object(
                 _server_mod, "build_preview_card",
                 side_effect=RuntimeError("prefab exploded"),
             ), \
             caplog.at_level(logging.WARNING, logger="mcp_sap_gui.server"):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert result.structured_content is None
        assert "nothing has been written" in _first_text(result)
        assert len(_image_blocks(result)) == 1
        assert "Preview card unavailable" in caplog.text

    @needs_prefab
    async def test_app_wiring_synthesizes_renderer_resource(self, srv):
        """app=True must stamp ui meta even though the tool returns ToolResult."""
        async with Client(_server_mod.mcp) as client:
            tools = {t.name: t for t in await client.list_tools()}
            uri = (tools["sap_preview"].meta or {}).get("ui", {}).get("resourceUri", "")
            assert uri.startswith("ui://prefab/tool/")
            assert uri in [str(r.uri) for r in await client.list_resources()]

    @pytest.mark.skipif(
        prefab_available(), reason="requires the apps extra to be absent"
    )
    async def test_ui_capable_host_without_prefab_gets_text(self, ui_capable, caplog):
        ctrl = _make_controller()
        _server_mod._prefab_hint_logged = False
        with patch.object(_server_mod, "_ctrl", return_value=ctrl), \
             caplog.at_level(logging.INFO, logger="mcp_sap_gui.server"):
            async with Client(_server_mod.mcp) as client:
                result = await client.call_tool("sap_preview", {})
        assert result.structured_content is None
        assert len(_image_blocks(result)) == 1
        assert "prefab-ui is not installed" in caplog.text
