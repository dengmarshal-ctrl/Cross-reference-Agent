import io
import unittest
import zipfile

from app.ooxml_processor import analyze_docx, create_sample_docx, process_docx


class OoxmlProcessorTest(unittest.TestCase):
    def test_analyze_sample_docx_builds_caption_and_reference_plan(self) -> None:
        sample = create_sample_docx()

        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["tables"], 3)
        self.assertEqual(plan["document_summary"]["table_references"], 4)
        self.assertEqual(plan["document_summary"]["high_confidence_references"], 3)
        self.assertEqual(plan["document_summary"]["medium_confidence_references"], 1)
        self.assertEqual(
            plan["caption_actions"][0]["original_text"],
            "表 14.1.1.2 试验完成情况总结 意向治疗分析集(ITT)",
        )
        self.assertEqual(plan["caption_actions"][0]["display_text"], "表 1 试验完成情况总结 意向治疗分析集(ITT)")
        self.assertEqual(plan["reference_actions"][0]["proposed_text"], "表格1")

    def test_process_sample_docx_writes_seq_and_ref_fields(self) -> None:
        sample = create_sample_docx()

        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertIn('w:instr=" SEQ 表 \\* ARABIC "', document_xml)
        self.assertIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)
        self.assertIn("_CSR_Table_001", document_xml)
        self.assertIn("表格", document_xml)
        self.assertIn(">1<", document_xml)
        self.assertEqual(audit["summary"]["caption_actions"], 3)
        self.assertEqual(audit["summary"]["cross_references_created"], 3)

    def test_detects_caption_inside_first_table_row_and_matches_new_number(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>研究期间，重要方案偏离发生率相近（见表 1）。</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>表 14.1.1.3 重要方案偏离情况 － 意向治疗分析集(ITT)</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>重要方案偏离</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample)
        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertEqual(plan["caption_actions"][0]["source_kind"], "table_first_row")
        self.assertEqual(plan["reference_actions"][0]["confidence"], "high")
        self.assertEqual(plan["reference_actions"][0]["reason"], "正文编号匹配题注重排后的新编号")
        self.assertIn('w:instr=" SEQ 表 \\* ARABIC "', document_xml)
        self.assertIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)
        self.assertNotIn("表 14.1.1.3 重要方案偏离情况", document_xml)
        self.assertEqual(audit["summary"]["cross_references_created"], 1)

    def test_recommends_nearest_table_for_statistical_list_reference(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>方案偏离的详细情况参见统计列表 16.2.2。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.1.1.3 重要方案偏离情况 － 意向治疗分析集(ITT)</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>重要方案偏离</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["medium_confidence_references"], 1)
        self.assertEqual(plan["reference_actions"][0]["confidence"], "medium")
        self.assertEqual(plan["reference_actions"][0]["proposed_text"], "统计列表 1")
        self.assertIn("最近表格", plan["reference_actions"][0]["reason"])

    def _document_xml(self, docx_bytes: bytes) -> str:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as docx:
            return docx.read("word/document.xml").decode("utf-8")

    def _docx(self, body_xml: str) -> bytes:
        document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
{body_xml}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
        files = {
            "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
""",
            "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
""",
            "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
""",
            "word/document.xml": document_xml,
        }
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as docx:
            for name, content in files.items():
                docx.writestr(name, content)
        return buffer.getvalue()


if __name__ == "__main__":
    unittest.main()

