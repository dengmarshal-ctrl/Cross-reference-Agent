import io
import unittest
import zipfile

from app.ooxml_processor import analyze_docx, create_sample_docx, process_docx


class OoxmlProcessorTest(unittest.TestCase):

    # ------------------------------------------------------------------
    # Basic table recognition (updated for new field names)
    # ------------------------------------------------------------------

    def test_analyze_sample_docx_builds_caption_and_reference_plan(self) -> None:
        sample = create_sample_docx()

        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["tables"], 3)
        self.assertEqual(plan["document_summary"]["images"], 2)
        self.assertEqual(plan["document_summary"]["table_references"], 4)
        self.assertEqual(plan["document_summary"]["image_references"], 2)
        self.assertEqual(plan["document_summary"]["high_confidence_references"], 5)
        self.assertEqual(plan["document_summary"]["medium_confidence_references"], 1)
        self.assertEqual(
            plan["caption_actions"][0]["original_text"],
            "表 14.1.1.2 试验完成情况总结 意向治疗分析集(ITT)",
        )
        self.assertEqual(
            plan["caption_actions"][0]["display_text"],
            "表 1 试验完成情况总结 意向治疗分析集(ITT)",
        )
        self.assertEqual(plan["caption_actions"][0]["object_type"], "table")
        # First table reference
        table_refs = [r for r in plan["reference_actions"] if r["object_type"] == "table"]
        self.assertEqual(table_refs[0]["proposed_text"], "表格1")

    def test_process_sample_docx_writes_seq_and_ref_fields(self) -> None:
        sample = create_sample_docx()

        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertIn('w:instr=" SEQ 表 \\* ARABIC "', document_xml)
        self.assertIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)
        self.assertIn("_CSR_Table_001", document_xml)
        self.assertIn("表格", document_xml)
        self.assertIn(">1<", document_xml)
        self.assertEqual(audit["summary"]["tables"], 3)
        self.assertEqual(audit["summary"]["images"], 2)
        self.assertEqual(audit["summary"]["caption_actions"], 5)
        self.assertGreaterEqual(audit["summary"]["cross_references_created"], 3)

    # ------------------------------------------------------------------
    # Image support
    # ------------------------------------------------------------------

    def test_detects_images_in_sample_docx(self) -> None:
        sample = create_sample_docx()
        plan = analyze_docx(sample)

        image_captions = [c for c in plan["caption_actions"] if c["object_type"] == "image"]
        self.assertEqual(len(image_captions), 2)
        self.assertEqual(image_captions[0]["display_text"], "图 1 研究流程图")
        self.assertEqual(image_captions[0]["action"], "ConvertTextToCaption")
        self.assertEqual(image_captions[0]["label"], "图")

    def test_image_references_detected_and_matched(self) -> None:
        sample = create_sample_docx()
        plan = analyze_docx(sample)

        image_refs = [r for r in plan["reference_actions"] if r["object_type"] == "image"]
        self.assertEqual(len(image_refs), 2)
        self.assertEqual(image_refs[0]["raw_text"], "图 1")
        self.assertEqual(image_refs[0]["confidence"], "high")
        self.assertIsNotNone(image_refs[0]["target_caption_id"])

    def test_process_writes_image_seq_and_ref_fields(self) -> None:
        sample = create_sample_docx()
        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertIn('w:instr=" SEQ 图 \\* ARABIC "', document_xml)
        self.assertIn("_CSR_Figure_001", document_xml)
        self.assertIn('w:instr=" REF _CSR_Figure_001 \\h "', document_xml)

    def test_image_only_document(self) -> None:
        sample = self._docx_with_image(
            """
    <w:p><w:r><w:t>如图 1 所示，研究流程如下。</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="5000000" cy="3000000"/>
            <wp:docPr id="1" name="Fig1"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>图 14.2.1 研究流程图</w:t></w:r></w:p>
            """
        )
        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["tables"], 0)
        self.assertEqual(plan["document_summary"]["images"], 1)
        self.assertEqual(plan["document_summary"]["image_references"], 1)
        image_caps = [c for c in plan["caption_actions"] if c["object_type"] == "image"]
        self.assertEqual(len(image_caps), 1)
        self.assertEqual(image_caps[0]["display_text"], "图 1 研究流程图")

    # ------------------------------------------------------------------
    # Table caption detection (existing tests updated for new fields)
    # ------------------------------------------------------------------

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
        self.assertEqual(plan["reference_actions"][0]["match_method"], "semantic_nearby_context")
        self.assertIn("语义匹配", plan["reference_actions"][0]["reason"])
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
        self.assertIn("最近对象", plan["reference_actions"][0]["reason"])

    def test_repeated_local_table_one_references_stay_in_context(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>1.1 研究人群</w:t></w:r></w:p>
    <w:p><w:r><w:t>按中心划分的入组受试者总数见表14.1.1.1。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.1.1.1 研究人群</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>筛选例数</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>1.2 方案偏离</w:t></w:r></w:p>
    <w:p><w:r><w:t>研究期间，两组发生率相近（见表1）。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表1 重要方案偏离情况 － 意向治疗分析集(ITT)</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>重要方案偏离</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>1.3 分析集</w:t></w:r></w:p>
    <w:p><w:r><w:t>各分析集的受试者分布情况见表1。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表1 分析数据集分布</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>分析数据集</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample)
        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertEqual(
            [reference["proposed_text"] for reference in plan["reference_actions"]],
            ["表1", "表2", "表3"],
        )
        self.assertEqual(plan["reference_actions"][1]["target_title"], "重要方案偏离情况 － 意向治疗分析集(ITT)")
        self.assertEqual(plan["reference_actions"][2]["target_title"], "分析数据集分布")
        self.assertIn("短编号引用", plan["reference_actions"][1]["reason"])
        self.assertEqual(audit["summary"]["cross_references_created"], 3)
        self.assertIn('w:instr=" REF _CSR_Table_002 \\h "', document_xml)
        self.assertIn('w:instr=" REF _CSR_Table_003 \\h "', document_xml)
        self.assertNotIn(">4<", document_xml)

    def test_short_reference_prefers_nearby_full_numbered_caption(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>1.1 已有局部表</w:t></w:r></w:p>
    <w:p><w:r><w:t>表1 非当前章节表格</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>旧章节数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>1.4.1 人口统计学和基线疾病特征</w:t></w:r></w:p>
    <w:p><w:r><w:t>基于意向治疗分析集（ITT），人口统计学和基线疾病特征见表1。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.1.2.1 人口学和基线特征总结 意向治疗分析集(ITT)</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>年龄（岁）</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample)
        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertEqual(plan["reference_actions"][0]["target_title"], "人口学和基线特征总结 意向治疗分析集(ITT)")
        self.assertEqual(plan["reference_actions"][0]["proposed_text"], "表2")
        self.assertEqual(plan["reference_actions"][0]["match_method"], "semantic_nearby_context")
        self.assertIn("语义匹配", plan["reference_actions"][0]["reason"])
        self.assertEqual(audit["summary"]["cross_references_created"], 1)
        self.assertIn('w:instr=" REF _CSR_Table_002 \\h "', document_xml)
        self.assertNotIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)

    def test_semantic_context_can_override_wrong_global_number(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>表4 其他章节无关表格</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>其他内容</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>表2 研究药物偏离情况</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>研究药物偏离</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>研究期间，在意向治疗分析集（ITT）中，共107例患者发生了重要方案偏离，两组发生率相近（见表4）。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表1 重要方案偏离情况 － 意向治疗分析集(ITT)</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>重要方案偏离</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample)
        output, audit = process_docx(sample)
        document_xml = self._document_xml(output)

        self.assertEqual(plan["reference_actions"][0]["target_title"], "重要方案偏离情况 － 意向治疗分析集(ITT)")
        self.assertEqual(plan["reference_actions"][0]["proposed_text"], "表3")
        self.assertEqual(plan["reference_actions"][0]["match_method"], "semantic_nearby_context")
        self.assertIn("重要方案偏离", plan["reference_actions"][0]["reason"])
        self.assertEqual(audit["summary"]["cross_references_created"], 1)
        self.assertIn('w:instr=" REF _CSR_Table_003 \\h "', document_xml)
        self.assertNotIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)

    # ------------------------------------------------------------------
    # Numbering strategy
    # ------------------------------------------------------------------

    def test_preserve_strategy_keeps_original_numbers(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>表 14.1.1.2 试验完成情况总结</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>表 14.1.1.3 筛选失败原因汇总</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample, strategy="preserve")

        self.assertEqual(plan["caption_actions"][0]["new_label"], "表 14.1.1.2")
        self.assertEqual(plan["caption_actions"][1]["new_label"], "表 14.1.1.3")

    def test_chapter_strategy_numbers_within_chapter(self) -> None:
        sample = self._docx(
            """
    <w:p><w:r><w:t>10 研究结果</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.1.1.2 试验完成情况</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>表 14.1.1.3 筛选失败原因</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>11 安全性评价</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.3.1.1 不良事件汇总</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
            """
        )

        plan = analyze_docx(sample, strategy="chapter")

        self.assertEqual(plan["caption_actions"][0]["new_label"], "表 10-1")
        self.assertEqual(plan["caption_actions"][1]["new_label"], "表 10-2")
        self.assertEqual(plan["caption_actions"][2]["new_label"], "表 11-1")

    # ------------------------------------------------------------------
    # Confirmation queue
    # ------------------------------------------------------------------

    def test_confirmed_medium_ref_gets_processed(self) -> None:
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
        medium_ref = plan["reference_actions"][0]
        self.assertEqual(medium_ref["confidence"], "medium")

        output_without, audit_without = process_docx(sample)
        self.assertEqual(audit_without["summary"]["cross_references_created"], 0)

        output_with, audit_with = process_docx(
            sample, confirmed_refs={medium_ref["id"]}
        )
        self.assertEqual(audit_with["summary"]["cross_references_created"], 1)
        self.assertEqual(audit_with["summary"]["user_confirmed"], 1)
        document_xml = self._document_xml(output_with)
        self.assertIn('w:instr=" REF _CSR_Table_001 \\h "', document_xml)

    # ------------------------------------------------------------------
    # Mixed table + image document
    # ------------------------------------------------------------------

    def test_mixed_table_and_image_processing(self) -> None:
        sample = self._docx_with_image(
            """
    <w:p><w:r><w:t>详见表 14.1.1.2 和图 1。</w:t></w:r></w:p>
    <w:p><w:r><w:t>表 14.1.1.2 试验完成情况</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="5000000" cy="3000000"/>
            <wp:docPr id="1" name="Fig1"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>图 14.2.1 研究流程图</w:t></w:r></w:p>
            """
        )

        plan = analyze_docx(sample)

        self.assertEqual(plan["document_summary"]["tables"], 1)
        self.assertEqual(plan["document_summary"]["images"], 1)
        table_refs = [r for r in plan["reference_actions"] if r["object_type"] == "table"]
        image_refs = [r for r in plan["reference_actions"] if r["object_type"] == "image"]
        self.assertEqual(len(table_refs), 1)
        self.assertEqual(len(image_refs), 1)
        self.assertEqual(table_refs[0]["target_caption_id"], "cap_tbl_001")
        self.assertEqual(image_refs[0]["target_caption_id"], "cap_img_001")

    # ------------------------------------------------------------------
    # Multiple references in one paragraph
    # ------------------------------------------------------------------

    def test_multiple_references_in_same_paragraph_all_get_ref_fields(self) -> None:
        """When a paragraph contains multiple references, ALL of them must
        become REF fields with blue color — not just the last one."""
        body = """
    <w:p><w:r><w:t>表 14.1.1.1 研究人群</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>表 14.1.1.2 试验完成情况总结</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>入组受试者总数见表14.1.1.1。完成情况见表 14.1.1.2。</w:t></w:r></w:p>
"""
        docx_bytes = self._docx(body)
        output, audit = process_docx(docx_bytes)
        xml = self._document_xml(output)

        self.assertEqual(xml.count("REF _CSR_Table_001"), 1)
        self.assertEqual(xml.count("REF _CSR_Table_002"), 1)
        self.assertEqual(xml.count('w:val="0000FF"'), 4)
        self.assertEqual(audit["summary"]["cross_references_created"], 2)

    # ------------------------------------------------------------------
    # Audit report structure
    # ------------------------------------------------------------------

    def test_audit_contains_strategy_and_object_types(self) -> None:
        sample = create_sample_docx()
        _, audit = process_docx(sample, strategy="continuous")

        self.assertEqual(audit["strategy"], "continuous")
        self.assertIn("tables", audit["summary"])
        self.assertIn("images", audit["summary"])
        self.assertIn("table_references", audit["summary"])
        self.assertIn("image_references", audit["summary"])
        self.assertIn("user_confirmed", audit["summary"])
        self.assertEqual(len(audit["logs"]), 10)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

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

    def _docx_with_image(self, body_xml: str) -> bytes:
        """Build a test docx with drawing namespace support."""
        document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
