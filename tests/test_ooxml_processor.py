import io
import unittest
import zipfile

from app.ooxml_processor import analyze_docx, create_sample_docx, process_docx


class OoxmlProcessorTest(unittest.TestCase):
    def test_analyze_sample_docx_builds_caption_and_reference_plan(self) -> None:
        sample = create_sample_docx()

        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["tables"], 2)
        self.assertEqual(plan["document_summary"]["table_references"], 2)
        self.assertEqual(plan["document_summary"]["high_confidence_references"], 2)
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
        self.assertEqual(audit["summary"]["caption_actions"], 2)
        self.assertEqual(audit["summary"]["cross_references_created"], 2)

    def _document_xml(self, docx_bytes: bytes) -> str:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as docx:
            return docx.read("word/document.xml").decode("utf-8")


if __name__ == "__main__":
    unittest.main()

