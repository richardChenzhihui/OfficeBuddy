"""Live MiniMax smoke test. Costs real API tokens — gated behind env vars:
    OFFICE_AGENT_LIVE_TEST=1 MINIMAX_API_KEY=... pytest -m live
"""
import os

import pytest

pytestmark = pytest.mark.live

requires_live = pytest.mark.skipif(
    not (os.getenv("OFFICE_AGENT_LIVE_TEST") and os.getenv("MINIMAX_API_KEY")),
    reason="live test disabled (set OFFICE_AGENT_LIVE_TEST=1 and MINIMAX_API_KEY)",
)


@requires_live
def test_live_tool_roundtrip(excel_doc_path):
    from office_agent.agent.loop import AgentSession
    from office_agent.config import Config

    config = Config(visual_verify=False, non_interactive=True)
    session = AgentSession(config)
    try:
        result = session.run_turn(
            "Open the spreadsheet, write the text 'LIVE_OK' into cell A1 of the "
            "first sheet, then save (do not overwrite the original).",
            str(excel_doc_path),
        )
        assert result.saved_paths, "model did not save"
        from openpyxl import load_workbook

        wb = load_workbook(result.saved_paths[0])
        assert wb.active["A1"].value == "LIVE_OK"
    finally:
        session.ctx.sessions.close_all()
