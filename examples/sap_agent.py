#!/usr/bin/env python3
"""
SAP Automation Agent - Document-driven SAP automation using Claude Code SDK.

Uses your existing Claude Max/Pro subscription (no API key needed).
Claude Code SDK spawns a Claude Code process that has access to the MCP SAP GUI
server via the project's .mcp.json — the same tools you use interactively.

This agent reads specification documents (Word, PDF, text) and executes
SAP transactions autonomously. It can:

  - Read a customizing spec and apply configuration in SAP
  - Execute test scenarios and capture results + screenshots
  - Analyze current SAP configuration and produce reports

Architecture:
    [Spec Document] --> [Claude Code SDK] --> [Claude Code + MCP SAP GUI] --> [SAP GUI]

Usage:
    # Interactive mode - ask Claude to do SAP tasks
    uv run python examples/sap_agent.py

    # With a spec document - Claude reads it and executes
    uv run python examples/sap_agent.py --doc "specs/putaway_strategy.docx"

    # Read-only analysis mode
    uv run python examples/sap_agent.py --read-only --doc "specs/current_config.txt"

    # With screenshot capture for documentation
    uv run python examples/sap_agent.py --doc "specs/test_scenario.docx" --output "results/"

Requirements:
    pip install claude-agent-sdk
    # Claude Code CLI must be installed and authenticated (claude --version)
    # The MCP SAP GUI server must be configured in .mcp.json (already done)
"""

import argparse
import asyncio
import sys
from pathlib import Path

from claude_agent_sdk import (
    ClaudeSDKClient,
    ClaudeAgentOptions,
    AssistantMessage,
    ResultMessage,
    TextBlock,
    ToolUseBlock,
)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# System prompt that makes Claude act as an SAP consultant
SAP_SYSTEM_PROMPT = """\
You are an SAP consultant automation agent. You interact with SAP GUI through \
the MCP SAP GUI tools (sap_connect, sap_execute_transaction, sap_set_field, etc.) \
to execute customizing, run transactions, and capture results.

Guidelines:
- Always connect to an existing SAP session first (mcp__sap-gui__sap_connect_existing)
- Use mcp__sap-gui__sap_get_screen_elements to discover field IDs on unfamiliar screens
- After setting fields, press Enter or Execute (F8) to confirm
- Every action tool returns screen info with active_window — if it says wnd[1],
  a popup appeared. Use mcp__sap-gui__sap_get_popup_window for full popup details.
- Take screenshots at key milestones for documentation
- When reading tables, start with a reasonable max_rows (50-100)
- For tree controls, expand nodes before trying to read children
- Use Read tool to read specification documents from disk

When given a specification document:
1. Read and understand the requirements
2. Plan the sequence of SAP transactions needed
3. Execute each step, verifying success before proceeding
4. Capture screenshots and document numbers as evidence
5. Summarize what was done and any issues encountered
"""

# All MCP SAP GUI tools + Read (for reading spec documents from disk)
ALLOWED_TOOLS = [
    # File reading (for spec documents)
    "Read",
    # MCP SAP GUI tools (prefixed with mcp__sap-gui__ by Claude Code)
    "mcp__sap-gui__sap_connect",
    "mcp__sap-gui__sap_connect_existing",
    "mcp__sap-gui__sap_list_connections",
    "mcp__sap-gui__sap_get_session_info",
    "mcp__sap-gui__sap_execute_transaction",
    "mcp__sap-gui__sap_send_key",
    "mcp__sap-gui__sap_get_screen_info",
    "mcp__sap-gui__sap_read_field",
    "mcp__sap-gui__sap_set_field",
    "mcp__sap-gui__sap_press_button",
    "mcp__sap-gui__sap_select_menu",
    "mcp__sap-gui__sap_select_checkbox",
    "mcp__sap-gui__sap_select_radio_button",
    "mcp__sap-gui__sap_select_combobox_entry",
    "mcp__sap-gui__sap_select_tab",
    "mcp__sap-gui__sap_read_table",
    "mcp__sap-gui__sap_get_alv_toolbar",
    "mcp__sap-gui__sap_press_alv_toolbar_button",
    "mcp__sap-gui__sap_select_alv_context_menu_item",
    "mcp__sap-gui__sap_select_table_row",
    "mcp__sap-gui__sap_double_click_cell",
    "mcp__sap-gui__sap_modify_cell",
    "mcp__sap-gui__sap_set_current_cell",
    "mcp__sap-gui__sap_get_column_info",
    "mcp__sap-gui__sap_read_tree",
    "mcp__sap-gui__sap_expand_tree_node",
    "mcp__sap-gui__sap_collapse_tree_node",
    "mcp__sap-gui__sap_select_tree_node",
    "mcp__sap-gui__sap_double_click_tree_node",
    "mcp__sap-gui__sap_double_click_tree_item",
    "mcp__sap-gui__sap_click_tree_link",
    "mcp__sap-gui__sap_find_tree_node_by_path",
    "mcp__sap-gui__sap_get_combobox_entries",
    "mcp__sap-gui__sap_set_batch_fields",
    "mcp__sap-gui__sap_read_textedit",
    "mcp__sap-gui__sap_set_textedit",
    "mcp__sap-gui__sap_set_focus",
    "mcp__sap-gui__sap_get_current_cell",
    "mcp__sap-gui__sap_scroll_table_control",
    "mcp__sap-gui__sap_get_table_control_row_info",
    "mcp__sap-gui__sap_select_all_table_control_columns",
    "mcp__sap-gui__sap_get_cell_info",
    "mcp__sap-gui__sap_press_column_header",
    "mcp__sap-gui__sap_select_all_rows",
    "mcp__sap-gui__sap_select_multiple_rows",
    "mcp__sap-gui__sap_get_popup_window",
    "mcp__sap-gui__sap_get_toolbar_buttons",
    "mcp__sap-gui__sap_read_shell_content",
    "mcp__sap-gui__sap_get_screen_elements",
    "mcp__sap-gui__sap_screenshot",
]


# ---------------------------------------------------------------------------
# Agent
# ---------------------------------------------------------------------------

async def run_agent(prompt: str, read_only: bool = False) -> None:
    """Run the SAP automation agent with the given prompt."""

    system = SAP_SYSTEM_PROMPT
    if read_only:
        system += "\n\nIMPORTANT: You are in READ-ONLY mode. Do not modify any SAP data."

    # Filter out write tools in read-only mode
    tools = list(ALLOWED_TOOLS)
    if read_only:
        write_tools = {
            "mcp__sap-gui__sap_execute_transaction",
            "mcp__sap-gui__sap_send_key",
            "mcp__sap-gui__sap_set_field",
            "mcp__sap-gui__sap_press_button",
            "mcp__sap-gui__sap_select_menu",
            "mcp__sap-gui__sap_select_checkbox",
            "mcp__sap-gui__sap_select_radio_button",
            "mcp__sap-gui__sap_select_combobox_entry",
            "mcp__sap-gui__sap_select_tab",
            "mcp__sap-gui__sap_set_batch_fields",
            "mcp__sap-gui__sap_set_textedit",
            "mcp__sap-gui__sap_set_focus",
            "mcp__sap-gui__sap_press_alv_toolbar_button",
            "mcp__sap-gui__sap_select_alv_context_menu_item",
            "mcp__sap-gui__sap_select_table_row",
            "mcp__sap-gui__sap_double_click_cell",
            "mcp__sap-gui__sap_modify_cell",
            "mcp__sap-gui__sap_set_current_cell",
            "mcp__sap-gui__sap_scroll_table_control",
            "mcp__sap-gui__sap_select_all_table_control_columns",
            "mcp__sap-gui__sap_press_column_header",
            "mcp__sap-gui__sap_select_all_rows",
            "mcp__sap-gui__sap_select_multiple_rows",
            "mcp__sap-gui__sap_expand_tree_node",
            "mcp__sap-gui__sap_collapse_tree_node",
            "mcp__sap-gui__sap_select_tree_node",
            "mcp__sap-gui__sap_double_click_tree_node",
            "mcp__sap-gui__sap_double_click_tree_item",
            "mcp__sap-gui__sap_click_tree_link",
        }
        tools = [t for t in tools if t not in write_tools]

    options = ClaudeAgentOptions(
        system_prompt=system,
        allowed_tools=tools,
        permission_mode="acceptEdits",
        cwd=str(Path(__file__).resolve().parent.parent),  # Project root for .mcp.json
    )

    async with ClaudeSDKClient(options=options) as client:
        await client.query(prompt)

        async for message in client.receive_response():
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        print(block.text)
                    elif isinstance(block, ToolUseBlock):
                        print(f"  -> {block.name}({str(block.input)[:100]})")

            elif isinstance(message, ResultMessage):
                print(f"\n--- Done ({message.num_turns} turns, "
                      f"{message.duration_ms / 1000:.1f}s) ---")
                if message.total_cost_usd:
                    print(f"Cost: ${message.total_cost_usd:.4f}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="SAP Automation Agent - Document-driven SAP automation using Claude Code SDK"
    )
    parser.add_argument(
        "--doc", type=str, default=None,
        help="Path to specification document (.txt, .md, .docx, .pdf)",
    )
    parser.add_argument(
        "--output", type=str, default=None,
        help="Directory to save screenshots and reports",
    )
    parser.add_argument(
        "--read-only", action="store_true",
        help="Run in read-only mode (no SAP modifications)",
    )
    args = parser.parse_args()

    if args.doc:
        doc_path = str(Path(args.doc).resolve())
        prompt = (
            f"I have a specification document at {doc_path}. "
            f"Please read it using the Read tool, then carry out the steps in SAP.\n"
        )
        if args.output:
            output_dir = str(Path(args.output).resolve())
            prompt += (
                f"\nTake screenshots at key milestones. "
                f"Save them to {output_dir} using appropriate filenames."
            )
        print(f"Loaded document: {args.doc}")
        print("Sending to Claude for execution...\n")
        asyncio.run(run_agent(prompt, read_only=args.read_only))
    else:
        # Interactive mode
        print("SAP Automation Agent (interactive mode)")
        print("Type your SAP task, or 'quit' to exit.\n")

        while True:
            try:
                user_input = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                break

            if not user_input or user_input.lower() in ("quit", "exit", "q"):
                break

            asyncio.run(run_agent(user_input, read_only=args.read_only))
            print()


if __name__ == "__main__":
    main()
