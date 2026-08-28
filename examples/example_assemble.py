"""Synthetic example: assemble two generated DOCX files."""

from pathlib import Path

from docx import Document

from assembly_engine.service import assemble_documents

OUT = Path(__file__).resolve().parent.parent / "outputs"
OUT.mkdir(exist_ok=True)

a = OUT / "example_asm_a.docx"
b = OUT / "example_asm_b.docx"
for path, text in ((a, "第一份材料内容"), (b, "第二份材料内容")):
    d = Document()
    d.add_paragraph(text)
    d.save(path)

result = assemble_documents([str(a), str(b)], str(OUT / "example_assembled.docx"), explicit_format_hint="正式公文")
print(result["status"])
print(result["processed"], result["failed"])
print(result["assembly_payload_sha256"][:16])
