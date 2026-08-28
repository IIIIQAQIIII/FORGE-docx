"""Mission 05 — MCP assemble_documents E2E (calls the server tool function directly)."""

import json
from pathlib import Path

from docx import Document

import server


def test_mcp_assemble_documents_e2e(tmp_path):
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    out = tmp_path / "assembled.docx"
    for path, text in ((a, "MCP-A"), (b, "MCP-B")):
        doc = Document()
        doc.add_paragraph(text)
        doc.save(path)

    raw = server.assemble_documents(
        source_paths=[str(a), str(b)],
        output_path=str(out),
        explicit_format_hint="正式公文",
    )
    result = json.loads(raw)
    assert result["status"] == "ok"
    assert result["total"] == 2
    assert result["processed"] == 2
    assert result["failed"] == 0
    assert out.exists()
    assert result["output"] == str(out)
