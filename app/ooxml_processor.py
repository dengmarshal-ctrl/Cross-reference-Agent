from __future__ import annotations

import copy
import io
import json
import re
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

NS = {"w": WORD_NS, "wp": WP_NS}
ET.register_namespace("w", WORD_NS)
ET.register_namespace("r", OFFICE_REL_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)

W = f"{{{WORD_NS}}}"

CAPTION_RE = re.compile(
    r"^\s*(?P<label>表格|表)\s*(?P<number>\d+(?:[.\-]\d+)*)\s*(?P<title>.+?)\s*$"
)
IMAGE_CAPTION_RE = re.compile(
    r"^\s*(?P<label>图片|图)\s*(?P<number>\d+(?:[.\-]\d+)*)\s*(?P<title>.+?)\s*$"
)
BARE_NUMBER_CAPTION_RE = re.compile(
    r"^\s*(?P<number>\d+(?:\.\d+)+)\s+(?P<title>.+?)\s*$"
)
REFERENCE_RE = re.compile(
    r"(?P<prefix>统计列表|列表|表格|图片|表|图)(?P<space>\s*)(?P<number>\d+(?:[.\-]\d+)*)"
)
HEADING_RE = re.compile(r"^\s*(?P<chapter>\d+(?:\.\d+)?)\s+\S")

VALID_STRATEGIES = {"continuous", "chapter", "preserve"}

CROSSREF_COLOR = "0000FF"


@dataclass
class CaptionCandidate:
    id: str
    object_id: str
    object_type: str  # "table" or "image"
    paragraph_id: str | None
    duplicate_paragraph_id: str | None
    source_kind: str
    body_index: int
    object_index: int
    original_text: str | None
    original_number: str | None
    title: str
    new_number: str
    bookmark_name: str
    label: str  # "表" or "图"

    @property
    def new_label(self) -> str:
        return f"{self.label} {self.new_number}"

    @property
    def display_text(self) -> str:
        return f"{self.new_label} {self.title}".strip()


def analyze_docx(
    docx_bytes: bytes, *, strategy: str = "continuous"
) -> dict[str, Any]:
    """Return a transparent processing plan without mutating the document."""
    strategy = _validated_strategy(strategy)
    document_xml = _read_document_xml(docx_bytes)
    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    block_records = _collect_blocks(body)
    captions = _build_caption_candidates(block_records, strategy=strategy)
    references = _detect_references(block_records, captions)

    return _build_plan(block_records, captions, references)


def process_docx(
    docx_bytes: bytes,
    *,
    strategy: str = "continuous",
    confirmed_refs: set[str] | None = None,
) -> tuple[bytes, dict[str, Any]]:
    """Normalize captions and replace matching text with REF fields.

    *confirmed_refs* is an optional set of reference IDs that the user has
    manually confirmed.  Medium-confidence items whose ID appears in this set
    are promoted to high confidence and processed automatically.
    """
    strategy = _validated_strategy(strategy)
    confirmed_refs = confirmed_refs or set()

    document_xml = _read_document_xml(docx_bytes)
    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    block_records = _collect_blocks(body)
    captions = _build_caption_candidates(block_records, strategy=strategy)
    references = _detect_references(block_records, captions)
    plan = _build_plan(block_records, captions, references)

    table_captions = [c for c in captions if c.object_type == "table"]
    image_captions = [c for c in captions if c.object_type == "image"]
    logs: list[dict[str, Any]] = [
        {
            "step": "Step 1",
            "title": "解析文档结构",
            "detail": (
                f"识别到 {len(table_captions)} 个表格对象、"
                f"{len(image_captions)} 个图片对象、"
                f"{len(references)} 处正文引用。"
            ),
            "status": "done",
        }
    ]

    max_bookmark_id = _max_bookmark_id(root)
    caption_paragraph_ids = {
        paragraph_id
        for caption in captions
        for paragraph_id in (caption.paragraph_id, caption.duplicate_paragraph_id)
        if paragraph_id
    }

    # Insert new caption paragraphs for objects without any detected title.
    created_paragraphs: dict[str, ET.Element] = {}
    insert_offset = 0
    for caption in captions:
        if caption.source_kind != "missing":
            continue
        paragraph = _new_paragraph()
        body.insert(caption.body_index + insert_offset, paragraph)
        insert_offset += 1
        created_paragraphs[caption.id] = paragraph

    block_records = _collect_blocks(body)
    paragraph_by_id = {
        record["id"]: record["element"]
        for record in block_records
        if record["type"] == "paragraph"
    }
    table_by_id = {
        record["id"]: record["element"]
        for record in block_records
        if record["type"] == "table"
    }

    logs.append(
        {
            "step": "Step 2",
            "title": "识别静态题注",
            "detail": (
                f"识别到 {sum(1 for c in captions if c.source_kind == 'paragraph')} 个段落题注、"
                f"{sum(1 for c in captions if c.source_kind == 'table_first_row')} 个表格内题名、"
                f"{sum(1 for c in captions if c.source_kind == 'image_adjacent')} 个图题、"
                f"{sum(1 for c in captions if c.source_kind == 'missing')} 个缺失题注。"
            ),
            "status": "done",
        }
    )

    audit_items: list[dict[str, Any]] = []

    # Clear duplicate paragraphs (e.g. external para when table-internal title wins).
    for caption in captions:
        if not caption.duplicate_paragraph_id:
            continue
        duplicate_paragraph = paragraph_by_id.get(caption.duplicate_paragraph_id)
        if duplicate_paragraph is not None:
            _clear_paragraph_text(duplicate_paragraph)

    for index, caption in enumerate(captions, start=1):
        max_bookmark_id += 1
        if caption.source_kind == "paragraph" and caption.paragraph_id:
            paragraph = paragraph_by_id.get(caption.paragraph_id)
        elif caption.source_kind == "table_first_row":
            paragraph = _first_table_cell_paragraph(table_by_id.get(caption.object_id))
        elif caption.source_kind == "image_adjacent" and caption.paragraph_id:
            paragraph = paragraph_by_id.get(caption.paragraph_id)
        else:
            paragraph = created_paragraphs.get(caption.id)
        if paragraph is None:
            continue
        before = _paragraph_text(paragraph) or caption.original_text or ""
        _set_caption_paragraph(paragraph, caption, max_bookmark_id)
        after = _paragraph_text(paragraph)
        audit_items.append(
            {
                "type": "caption",
                "id": caption.id,
                "target": caption.object_id,
                "object_type": caption.object_type,
                "before": before,
                "after": after,
                "action": _caption_action_name(caption),
                "bookmark": caption.bookmark_name,
            }
        )

    logs.append(
        {
            "step": "Step 3",
            "title": "创建标准题注",
            "detail": f"将 {len(captions)} 个静态/缺失题注转换为带 SEQ 字段的题注。",
            "status": "done",
        }
    )

    mapping_examples = []
    for c in captions[:3]:
        if c.original_number:
            mapping_examples.append(f"{c.original_number} → {c.new_number}")
    logs.append(
        {
            "step": "Step 4",
            "title": "建立旧编号到新编号映射",
            "detail": (
                f"生成 {len(captions)} 条映射"
                + (f"，例如 {'、'.join(mapping_examples)}" if mapping_examples else "")
                + "。"
            ),
            "status": "done",
        }
    )

    table_refs = [r for r in references if _reference_object_type(r.get("prefix_label", "")) == "table"]
    image_refs = [r for r in references if _reference_object_type(r.get("prefix_label", "")) == "image"]
    logs.append(
        {
            "step": "Step 5",
            "title": "识别正文引用",
            "detail": f"发现 {len(table_refs)} 处表格引用、{len(image_refs)} 处图片引用。",
            "status": "done",
        }
    )

    logs.append(
        {
            "step": "Step 6",
            "title": "匹配引用目标",
            "detail": (
                f"高置信度 {sum(1 for r in references if r['confidence'] == 'high')} 处、"
                f"中置信度 {sum(1 for r in references if r['confidence'] == 'medium')} 处、"
                f"低置信度 {sum(1 for r in references if r['confidence'] == 'low')} 处。"
            ),
            "status": "done",
        }
    )

    # Collect eligible references, then group by paragraph and apply all at once
    # so that multiple references in the same paragraph don't destroy each other.
    current_blocks = _collect_blocks(body)
    current_paragraphs = {
        record["id"]: record["element"]
        for record in current_blocks
        if record["type"] == "paragraph"
    }
    applied_references = 0
    skipped_references = 0
    confirmed_count = 0

    eligible: list[tuple[dict[str, Any], CaptionCandidate, bool]] = []
    for reference in references:
        if reference["paragraph_id"] in caption_paragraph_ids:
            skipped_references += 1
            continue
        target = next(
            (c for c in captions if c.id == reference.get("target_caption_id")),
            None,
        )
        if current_paragraphs.get(reference["paragraph_id"]) is None or target is None:
            skipped_references += 1
            continue
        is_confirmed = reference["id"] in confirmed_refs
        effective_confidence = reference["confidence"]
        if effective_confidence == "medium" and is_confirmed:
            effective_confidence = "high"
            confirmed_count += 1
        if effective_confidence != "high":
            skipped_references += 1
            continue
        eligible.append((reference, target, is_confirmed))

    para_groups: dict[str, list[tuple[dict[str, Any], CaptionCandidate, bool]]] = {}
    for ref, target, is_confirmed in eligible:
        para_groups.setdefault(ref["paragraph_id"], []).append((ref, target, is_confirmed))

    for para_id, group in para_groups.items():
        paragraph = current_paragraphs.get(para_id)
        if paragraph is None:
            skipped_references += len(group)
            continue
        before = _paragraph_text(paragraph)
        pairs = [(ref, target) for ref, target, _ in group]
        replaced_pairs = _replace_references_in_paragraph(paragraph, pairs)
        after = _paragraph_text(paragraph)
        replaced_ids = {ref["id"] for ref, _ in replaced_pairs}
        for ref, target, is_confirmed in group:
            if ref["id"] in replaced_ids:
                applied_references += 1
                audit_items.append(
                    {
                        "type": "cross_reference",
                        "id": ref["id"],
                        "target": target.id,
                        "object_type": target.object_type,
                        "before": before,
                        "after": after,
                        "action": "ReplaceTextWithCrossReference",
                        "bookmark": target.bookmark_name,
                        "confidence": ref["confidence"],
                        "user_confirmed": is_confirmed,
                        "match_method": ref["match_method"],
                        "reason": ref["reason"],
                    }
                )
            else:
                skipped_references += 1

    logs.append(
        {
            "step": "Step 7",
            "title": "改写正文引用",
            "detail": f"改写 {applied_references} 处正文引用编号。",
            "status": "done",
        }
    )

    logs.append(
        {
            "step": "Step 8",
            "title": "创建真实交叉引用",
            "detail": (
                f"创建 {applied_references} 处 REF 字段"
                + (f"（含 {confirmed_count} 处用户确认项）" if confirmed_count else "")
                + f"，跳过 {skipped_references} 处。"
            ),
            "status": "done",
        }
    )

    logs.append(
        {
            "step": "Step 9",
            "title": "处理人工确认项",
            "detail": (
                f"共 {sum(1 for r in references if r['confidence'] == 'medium')} 项进入确认队列，"
                f"用户已确认 {confirmed_count} 项。"
            ),
            "status": "done",
        }
    )

    updated_xml = _serialize_xml(root)
    processed_bytes = _replace_document_xml(docx_bytes, updated_xml)

    logs.append(
        {
            "step": "Step 10",
            "title": "生成审计报告",
            "detail": "已生成完整审计报告，包含每条修改的修改前后文本、置信度和匹配方法。",
            "status": "done",
        }
    )

    audit = {
        "task_id": f"csr-demo-{uuid.uuid4().hex[:8]}",
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "strategy": strategy,
        "summary": {
            "tables": len(table_captions),
            "images": len(image_captions),
            "caption_actions": len(captions),
            "references_detected": len(references),
            "table_references": len(table_refs),
            "image_references": len(image_refs),
            "cross_references_created": applied_references,
            "user_confirmed": confirmed_count,
            "skipped_references": skipped_references,
        },
        "plan": plan,
        "logs": logs,
        "audit_items": audit_items,
        "limitations": [
            "Demo 使用 OOXML 直接写入 SEQ/REF 字段，真实产品建议由 WPS Word Skill 提供对象级能力。",
            "Demo 为便于演示会重建被处理段落的 runs，复杂局部样式需由底层 Word 能力保障。",
        ],
    }
    return processed_bytes, audit


def create_sample_docx() -> bytes:
    wp_ns = WP_NS
    a_ns = A_NS
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="{wp_ns}"
            xmlns:a="{a_ns}"
            xmlns:r="{OFFICE_REL_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>10.1 受试者分布</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>基于所有入组人群总结的患者分布见表格14.1.1.2。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>筛选失败原因详见表 14.1.1.3。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>表 14.1.1.2 试验完成情况总结 意向治疗分析集(ITT)</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>分析集</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>例数</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>ITT</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>30</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t>表 14.1.1.3 筛选失败原因汇总</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>原因</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>例数</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>不满足入选标准</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t>研究期间，重要方案偏离发生率相近（见表 3）。方案偏离的详细情况参见统计列表 16.2.2。</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>表 14.1.1.3 重要方案偏离情况 － 意向治疗分析集(ITT)</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>重要方案偏离</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>107</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t>10.2 研究流程</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>如图 1 所示，研究整体流程包括筛选、随机、治疗和随访四个阶段。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="5274310" cy="3076575"/>
            <wp:docPr id="1" name="Figure1"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>图 14.2.1 研究流程图</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>不良事件的分布如图 2 所示。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="5274310" cy="2540000"/>
            <wp:docPr id="2" name="Figure2"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>图 14.3.1 不良事件分布</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
    files = {
        "[Content_Types].xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
""",
        "_rels/.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>
""",
        "word/_rels/document.xml.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}"/>
""",
        "word/document.xml": document_xml,
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as docx:
        for name, content in files.items():
            docx.writestr(name, content)
    return buffer.getvalue()


def audit_to_json(audit: dict[str, Any]) -> bytes:
    return json.dumps(audit, ensure_ascii=False, indent=2).encode("utf-8")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _validated_strategy(strategy: str) -> str:
    if strategy not in VALID_STRATEGIES:
        return "continuous"
    return strategy


def _read_document_xml(docx_bytes: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as docx:
        try:
            return docx.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("The uploaded file is not a valid .docx document") from exc


def _replace_document_xml(docx_bytes: bytes, document_xml: bytes) -> bytes:
    source = io.BytesIO(docx_bytes)
    target = io.BytesIO()
    with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(
        target, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            if item.filename == "word/document.xml":
                dst.writestr(item, document_xml)
            else:
                dst.writestr(item, src.read(item.filename))
    return target.getvalue()


def _collect_blocks(body: ET.Element) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    paragraph_count = 0
    table_count = 0
    image_count = 0
    for index, child in enumerate(list(body)):
        if child.tag == f"{W}p":
            paragraph_count += 1
            has_image = _paragraph_has_image(child)
            if has_image:
                image_count += 1
            records.append(
                {
                    "id": f"p_{paragraph_count:04d}",
                    "type": "paragraph",
                    "body_index": index,
                    "text": _paragraph_text(child),
                    "element": child,
                    "has_image": has_image,
                    "image_id": f"img_{image_count:04d}" if has_image else None,
                }
            )
        elif child.tag == f"{W}tbl":
            table_count += 1
            records.append(
                {
                    "id": f"tbl_{table_count:04d}",
                    "type": "table",
                    "body_index": index,
                    "text": _table_text(child),
                    "element": child,
                    "has_image": False,
                    "image_id": None,
                }
            )
    return records


def _build_caption_candidates(
    blocks: list[dict[str, Any]], *, strategy: str = "continuous"
) -> list[CaptionCandidate]:
    captions: list[CaptionCandidate] = []

    # --- Table captions ---
    table_number = 0
    chapter_table_counters: dict[str, int] = {}
    for index, record in enumerate(blocks):
        if record["type"] != "table":
            continue
        table_number += 1
        previous = _previous_non_empty_paragraph(blocks, index)
        paragraph_match = CAPTION_RE.match(previous["text"]) if previous else None
        bare_match = (
            BARE_NUMBER_CAPTION_RE.match(previous["text"])
            if previous and not paragraph_match
            else None
        )
        table_caption = _first_row_caption(record["element"])

        if table_caption and paragraph_match:
            original_text = table_caption["text"]
            original_number = table_caption["number"]
            title = table_caption["title"]
            paragraph_id = None
            duplicate_paragraph_id = previous["id"]
            source_kind = "table_first_row"
        elif paragraph_match:
            original_text = previous["text"]
            original_number = paragraph_match.group("number")
            title = paragraph_match.group("title")
            paragraph_id = previous["id"]
            duplicate_paragraph_id = None
            source_kind = "paragraph"
        elif table_caption and bare_match:
            original_text = table_caption["text"]
            original_number = table_caption["number"]
            title = table_caption["title"]
            paragraph_id = None
            duplicate_paragraph_id = previous["id"]
            source_kind = "table_first_row"
        elif bare_match:
            original_text = previous["text"]
            original_number = bare_match.group("number")
            title = bare_match.group("title")
            paragraph_id = previous["id"]
            duplicate_paragraph_id = None
            source_kind = "paragraph"
        elif table_caption:
            original_text = table_caption["text"]
            original_number = table_caption["number"]
            title = table_caption["title"]
            paragraph_id = None
            duplicate_paragraph_id = None
            source_kind = "table_first_row"
        else:
            original_text = None
            original_number = None
            title = "未命名表格"
            paragraph_id = None
            duplicate_paragraph_id = None
            source_kind = "missing"

        new_number = _compute_new_number(
            strategy=strategy,
            sequential=table_number,
            original_number=original_number,
            body_index=record["body_index"],
            blocks=blocks,
            chapter_counters=chapter_table_counters,
        )

        captions.append(
            CaptionCandidate(
                id=f"cap_tbl_{table_number:03d}",
                object_id=record["id"],
                object_type="table",
                paragraph_id=paragraph_id,
                duplicate_paragraph_id=duplicate_paragraph_id,
                source_kind=source_kind,
                body_index=record["body_index"],
                object_index=table_number,
                original_text=original_text,
                original_number=original_number,
                title=title,
                new_number=new_number,
                bookmark_name=f"_CSR_Table_{table_number:03d}",
                label="表",
            )
        )

    # --- Image captions ---
    image_number = 0
    chapter_image_counters: dict[str, int] = {}
    for index, record in enumerate(blocks):
        if record["type"] != "paragraph" or not record.get("has_image"):
            continue
        image_number += 1

        caption_info = _find_image_caption(blocks, index)
        if caption_info:
            original_text = caption_info["text"]
            original_number = caption_info["number"]
            title = caption_info["title"]
            paragraph_id = caption_info["paragraph_id"]
            source_kind = "image_adjacent"
        else:
            original_text = None
            original_number = None
            title = "未命名图片"
            paragraph_id = None
            source_kind = "missing"

        new_number = _compute_new_number(
            strategy=strategy,
            sequential=image_number,
            original_number=original_number,
            body_index=record["body_index"],
            blocks=blocks,
            chapter_counters=chapter_image_counters,
        )

        captions.append(
            CaptionCandidate(
                id=f"cap_img_{image_number:03d}",
                object_id=record.get("image_id", f"img_{image_number:04d}"),
                object_type="image",
                paragraph_id=paragraph_id,
                duplicate_paragraph_id=None,
                source_kind=source_kind,
                body_index=record["body_index"],
                object_index=image_number,
                original_text=original_text,
                original_number=original_number,
                title=title,
                new_number=new_number,
                bookmark_name=f"_CSR_Figure_{image_number:03d}",
                label="图",
            )
        )

    return captions


def _compute_new_number(
    *,
    strategy: str,
    sequential: int,
    original_number: str | None,
    body_index: int,
    blocks: list[dict[str, Any]],
    chapter_counters: dict[str, int],
) -> str:
    if strategy == "preserve":
        return original_number or str(sequential)
    if strategy == "chapter":
        chapter = _detect_chapter_number(blocks, body_index)
        chapter_counters.setdefault(chapter, 0)
        chapter_counters[chapter] += 1
        return f"{chapter}-{chapter_counters[chapter]}"
    return str(sequential)


def _detect_chapter_number(blocks: list[dict[str, Any]], body_index: int) -> str:
    for i in range(len(blocks) - 1, -1, -1):
        block = blocks[i]
        if block["body_index"] >= body_index:
            continue
        if block["type"] != "paragraph":
            continue
        match = HEADING_RE.match(block["text"])
        if match:
            return match.group("chapter").split(".")[0]
    return "0"


def _find_image_caption(
    blocks: list[dict[str, Any]], image_index: int
) -> dict[str, Any] | None:
    """Look for a figure caption paragraph adjacent to the image paragraph."""
    # Prefer caption below the image (common in CSR reports).
    below = _next_non_empty_paragraph(blocks, image_index)
    if below:
        match = IMAGE_CAPTION_RE.match(below["text"])
        if match:
            return {
                "text": below["text"],
                "number": match.group("number"),
                "title": match.group("title"),
                "paragraph_id": below["id"],
            }
    above = _previous_non_empty_paragraph(blocks, image_index)
    if above:
        match = IMAGE_CAPTION_RE.match(above["text"])
        if match:
            return {
                "text": above["text"],
                "number": match.group("number"),
                "title": match.group("title"),
                "paragraph_id": above["id"],
            }
    return None


def _detect_references(
    blocks: list[dict[str, Any]], captions: list[CaptionCandidate]
) -> list[dict[str, Any]]:
    caption_paragraph_ids = {
        paragraph_id
        for caption in captions
        for paragraph_id in (caption.paragraph_id, caption.duplicate_paragraph_id)
        if paragraph_id
    }
    by_old_number = _index_captions_by_number(captions, use_new_number=False)
    by_new_number = _index_captions_by_number(captions, use_new_number=True)
    caption_positions = {caption.id: caption.body_index for caption in captions}
    references: list[dict[str, Any]] = []
    ref_index = 0
    for record in blocks:
        if record["type"] != "paragraph" or record["id"] in caption_paragraph_ids:
            continue
        for match in REFERENCE_RE.finditer(record["text"]):
            ref_index += 1
            number = match.group("number")
            prefix = f"{match.group('prefix')}{match.group('space')}"
            prefix_label = match.group("prefix")
            target, confidence, reason, match_method = _match_reference_target(
                number=number,
                prefix_label=prefix_label,
                paragraph_body_index=record["body_index"],
                context_text=record["text"],
                old_number_index=by_old_number,
                new_number_index=by_new_number,
                captions=captions,
                caption_positions=caption_positions,
            )
            references.append(
                {
                    "id": f"ref_{ref_index:03d}",
                    "paragraph_id": record["id"],
                    "raw_text": match.group(0),
                    "prefix": prefix,
                    "prefix_label": prefix_label,
                    "original_number": number,
                    "context": record["text"],
                    "object_type": _reference_object_type(prefix_label),
                    "target_caption_id": target.id if target else None,
                    "target_title": target.title if target else None,
                    "new_number": target.new_number if target else None,
                    "proposed_text": f"{prefix}{target.new_number}" if target else None,
                    "confidence": confidence,
                    "match_method": match_method,
                    "reason": reason,
                }
            )
    return references


def _build_plan(
    blocks: list[dict[str, Any]],
    captions: list[CaptionCandidate],
    references: list[dict[str, Any]],
) -> dict[str, Any]:
    table_captions = [c for c in captions if c.object_type == "table"]
    image_captions = [c for c in captions if c.object_type == "image"]
    table_refs = [r for r in references if r.get("object_type") == "table"]
    image_refs = [r for r in references if r.get("object_type") == "image"]
    return {
        "document_summary": {
            "paragraphs": sum(1 for block in blocks if block["type"] == "paragraph"),
            "tables": len(table_captions),
            "images": len(image_captions),
            "table_references": len(table_refs),
            "image_references": len(image_refs),
            "high_confidence_references": sum(
                1 for reference in references if reference["confidence"] == "high"
            ),
            "medium_confidence_references": sum(
                1 for reference in references if reference["confidence"] == "medium"
            ),
            "low_confidence_references": sum(
                1 for reference in references if reference["confidence"] == "low"
            ),
        },
        "caption_actions": [
            {
                "id": caption.id,
                "object_id": caption.object_id,
                "object_type": caption.object_type,
                "source_paragraph_id": caption.paragraph_id,
                "duplicate_paragraph_id": caption.duplicate_paragraph_id,
                "source_kind": caption.source_kind,
                "original_text": caption.original_text,
                "original_number": caption.original_number,
                "new_label": caption.new_label,
                "title": caption.title,
                "display_text": caption.display_text,
                "bookmark": caption.bookmark_name,
                "action": _caption_action_name(caption),
                "label": caption.label,
            }
            for caption in captions
        ],
        "reference_actions": references,
    }


def _caption_action_name(caption: CaptionCandidate) -> str:
    if caption.source_kind == "paragraph":
        return "ConvertTextToCaption"
    if caption.source_kind == "table_first_row":
        return "ConvertTableRowTitleToCaption"
    if caption.source_kind == "image_adjacent":
        return "ConvertTextToCaption"
    return "InsertCaption"


def _reference_object_type(prefix_label: str) -> str:
    if prefix_label in {"图", "图片"}:
        return "image"
    return "table"


def _match_reference_target(
    *,
    number: str,
    prefix_label: str,
    paragraph_body_index: int,
    context_text: str,
    old_number_index: dict[str, list[CaptionCandidate]],
    new_number_index: dict[str, list[CaptionCandidate]],
    captions: list[CaptionCandidate],
    caption_positions: dict[str, int],
) -> tuple[CaptionCandidate | None, str, str, str]:
    target_type = _reference_object_type(prefix_label)
    type_captions = [c for c in captions if c.object_type == target_type]
    is_precise = "." in number or "-" in number

    # --- Phase 1: precise multi-part numbers (e.g. "14.1.1.1") use exact
    # match first because they uniquely identify a caption and must not be
    # overridden by paragraph-level semantic context. ---
    if is_precise:
        old_number_targets = [
            c for c in old_number_index.get(number, []) if c.object_type == target_type
        ]
        if old_number_targets:
            target = _best_context_candidate(paragraph_body_index, old_number_targets, caption_positions)
            if target:
                if len(old_number_targets) == 1:
                    return target, "high", "原始编号唯一匹配到对应题注", "source_number_exact"
                return (
                    target,
                    "high",
                    "存在重复局部编号，按同章节/段落后的最近同编号对象匹配",
                    "duplicate_source_number_nearby",
                )

    # --- Phase 2: semantic context matching (whole paragraph vs caption
    # title).  For simple numbers like "表 3" this runs before exact-match
    # so that a semantically closer caption can win over a distant one
    # that happens to share the same short number. ---
    if prefix_label in {"表", "表格", "图", "图片"}:
        semantic_target, semantic_reason = _semantic_nearby_candidate(
            context_text=context_text,
            paragraph_body_index=paragraph_body_index,
            captions=type_captions,
            caption_positions=caption_positions,
        )
        if semantic_target:
            return semantic_target, "high", semantic_reason, "semantic_nearby_context"

    # --- Phase 3: short-number nearby match ---
    if _is_simple_number(number):
        nearest_caption = _nearest_caption_after(paragraph_body_index, type_captions, caption_positions)
        if nearest_caption and _caption_matches_short_number(nearest_caption, number):
            return (
                nearest_caption,
                "high",
                "短编号引用按当前段落后的最近相关对象匹配",
                "short_number_nearby_caption",
            )

    # --- Phase 4: exact old-number match (simple numbers that were not
    # captured by semantic or short-number nearby) ---
    if not is_precise:
        old_number_targets = [
            c for c in old_number_index.get(number, []) if c.object_type == target_type
        ]
        if old_number_targets:
            target = _best_context_candidate(paragraph_body_index, old_number_targets, caption_positions)
            if target:
                if len(old_number_targets) == 1:
                    return target, "high", "原始编号唯一匹配到对应题注", "source_number_exact"
                return (
                    target,
                    "high",
                    "存在重复局部编号，按同章节/段落后的最近同编号对象匹配",
                    "duplicate_source_number_nearby",
                )

    # --- Phase 5: new (renumbered) number match ---
    old_has_number = any(
        c.object_type == target_type for c in old_number_index.get(number, [])
    )
    if not old_has_number:
        new_number_targets = [
            c for c in new_number_index.get(number, []) if c.object_type == target_type
        ]
        target = _best_context_candidate(paragraph_body_index, new_number_targets, caption_positions)
        if target:
            return target, "high", "正文编号匹配题注重排后的新编号", "generated_number_exact"

    # --- Phase 6: fallback to nearest caption ---
    nearest_caption = _nearest_caption_after(paragraph_body_index, type_captions, caption_positions)
    if nearest_caption:
        return (
            nearest_caption,
            "medium",
            "未命中编号，按同章节/段落后的最近对象推荐匹配",
            "nearby_fallback",
        )

    return None, "low", "未找到具有相同编号或邻近位置的对应题注", "unmatched"


def _is_simple_number(number: str) -> bool:
    return bool(re.fullmatch(r"\d+", number))


def _caption_matches_short_number(caption: CaptionCandidate, number: str) -> bool:
    if caption.new_number == number or caption.original_number == number:
        return True
    if not caption.original_number:
        return False
    parts = re.split(r"[.\-]", caption.original_number)
    return bool(parts and parts[-1] == number)


def _semantic_nearby_candidate(
    *,
    context_text: str,
    paragraph_body_index: int,
    captions: list[CaptionCandidate],
    caption_positions: dict[str, int],
) -> tuple[CaptionCandidate | None, str]:
    candidates: list[tuple[float, int, list[str], CaptionCandidate]] = []
    normalized_context = _normalize_for_semantic_match(context_text)
    for caption in captions:
        position = caption_positions.get(caption.id, -1)
        if position <= paragraph_body_index:
            continue
        distance = position - paragraph_body_index
        if distance > 12:
            continue
        score, matched_terms = _caption_context_score(normalized_context, caption.title)
        if score <= 0:
            continue
        candidates.append((score, distance, matched_terms, caption))

    if not candidates:
        return None, ""

    candidates.sort(key=lambda item: (-item[0], item[1]))
    best_score, best_distance, matched_terms, best_caption = candidates[0]
    if best_score < 0.28:
        return None, ""
    return (
        best_caption,
        "正文整句与附近题注语义匹配，关键词："
        + "、".join(matched_terms[:5])
        + f"；段落距离：{best_distance}",
    )


def _caption_context_score(normalized_context: str, caption_title: str) -> tuple[float, list[str]]:
    terms = _semantic_terms(caption_title)
    if not terms:
        return 0.0, []
    matched_terms: list[str] = []
    matched_weight = 0
    for term in terms:
        if term in normalized_context and not any(term in existing for existing in matched_terms):
            matched_terms.append(term)
            matched_weight += len(term)
    denominator = max(10, min(28, len(_normalize_for_semantic_match(caption_title))))
    return matched_weight / denominator, matched_terms


def _semantic_terms(text: str) -> list[str]:
    normalized = _normalize_for_semantic_match(text)
    stop_terms = {
        "意向治疗分析集",
        "治疗分析集",
        "分析集",
        "情况",
        "总结",
        "统计",
        "列表",
        "表格",
        "研究",
        "患者",
        "受试者",
        "ITT",
    }
    terms: set[str] = set()
    for chunk in re.findall(r"[\u4e00-\u9fffA-Za-z]+", normalized):
        if chunk in stop_terms:
            continue
        if re.fullmatch(r"[A-Za-z]+", chunk):
            if len(chunk) >= 3 and chunk.upper() not in stop_terms:
                terms.add(chunk.upper())
            continue
        max_len = min(8, len(chunk))
        for length in range(max_len, 1, -1):
            for start in range(0, len(chunk) - length + 1):
                term = chunk[start : start + length]
                if term not in stop_terms and not any(term in stop for stop in stop_terms):
                    terms.add(term)
    return sorted(terms, key=lambda value: (-len(value), value))


def _normalize_for_semantic_match(text: str) -> str:
    return re.sub(r"[\s　,，.。;；:：()（）\\－—_、/\-]+", "", text).upper()


def _index_captions_by_number(
    captions: list[CaptionCandidate], *, use_new_number: bool
) -> dict[str, list[CaptionCandidate]]:
    index: dict[str, list[CaptionCandidate]] = {}
    for caption in captions:
        number = caption.new_number if use_new_number else caption.original_number
        if not number:
            continue
        index.setdefault(number, []).append(caption)
    return index


def _best_context_candidate(
    paragraph_body_index: int,
    candidates: list[CaptionCandidate],
    caption_positions: dict[str, int],
) -> CaptionCandidate | None:
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    after_candidates = [
        caption
        for caption in candidates
        if caption_positions.get(caption.id, -1) > paragraph_body_index
    ]
    if after_candidates:
        after_candidates.sort(key=lambda caption: caption_positions.get(caption.id, 10**9))
        return after_candidates[0]

    candidates.sort(key=lambda caption: abs(caption_positions.get(caption.id, 10**9) - paragraph_body_index))
    return candidates[0]


def _nearest_caption_after(
    paragraph_body_index: int,
    captions: list[CaptionCandidate],
    caption_positions: dict[str, int],
) -> CaptionCandidate | None:
    candidates = [
        caption
        for caption in captions
        if caption_positions.get(caption.id, -1) > paragraph_body_index
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda caption: caption_positions.get(caption.id, 10**9))
    nearest = candidates[0]
    distance = caption_positions.get(nearest.id, 10**9) - paragraph_body_index
    return nearest if distance <= 8 else None


def _previous_non_empty_paragraph(
    blocks: list[dict[str, Any]], index: int
) -> dict[str, Any] | None:
    for candidate_index in range(index - 1, -1, -1):
        candidate = blocks[candidate_index]
        if candidate["type"] == "paragraph" and candidate["text"].strip():
            return candidate
        if candidate["type"] == "table":
            return None
    return None


def _next_non_empty_paragraph(
    blocks: list[dict[str, Any]], index: int
) -> dict[str, Any] | None:
    for candidate_index in range(index + 1, len(blocks)):
        candidate = blocks[candidate_index]
        if candidate["type"] == "paragraph" and candidate["text"].strip():
            return candidate
        if candidate["type"] == "table":
            return None
    return None


def _paragraph_has_image(paragraph: ET.Element) -> bool:
    for run in paragraph.findall(f"{W}r"):
        if run.find(f"{W}drawing") is not None:
            return True
        if run.find(f"{W}pict") is not None:
            return True
    return False


def _paragraph_text(paragraph: ET.Element) -> str:
    texts = []
    for text_node in paragraph.findall(".//w:t", NS):
        texts.append(text_node.text or "")
    return "".join(texts)


def _table_text(table: ET.Element) -> str:
    texts = []
    for text_node in table.findall(".//w:t", NS):
        if text_node.text:
            texts.append(text_node.text)
    return " ".join(texts)


def _first_row_caption(table: ET.Element | None) -> dict[str, str] | None:
    if table is None:
        return None
    first_row = table.find("w:tr", NS)
    if first_row is None:
        return None
    text = _row_text(first_row).strip()
    match = CAPTION_RE.match(text)
    if match:
        return {
            "text": text,
            "number": match.group("number"),
            "title": match.group("title"),
        }
    bare = BARE_NUMBER_CAPTION_RE.match(text)
    if bare:
        return {
            "text": text,
            "number": bare.group("number"),
            "title": bare.group("title"),
        }
    return None


def _row_text(row: ET.Element) -> str:
    texts = []
    for text_node in row.findall(".//w:t", NS):
        if text_node.text:
            texts.append(text_node.text)
    return "".join(texts)


def _first_table_cell_paragraph(table: ET.Element | None) -> ET.Element | None:
    if table is None:
        return None
    first_row = table.find("w:tr", NS)
    if first_row is None:
        return None
    first_cell = first_row.find("w:tc", NS)
    if first_cell is None:
        return None
    paragraph = first_cell.find("w:p", NS)
    if paragraph is not None:
        return paragraph
    paragraph = _new_paragraph()
    first_cell.insert(0, paragraph)
    return paragraph


def _new_paragraph() -> ET.Element:
    return ET.Element(f"{W}p")


def _set_caption_paragraph(
    paragraph: ET.Element, caption: CaptionCandidate, bookmark_id: int
) -> None:
    ppr = paragraph.find("w:pPr", NS)
    saved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    paragraph.clear()
    if saved_ppr is not None:
        paragraph.append(saved_ppr)

    paragraph.append(_bookmark_start(bookmark_id, caption.bookmark_name))
    paragraph.append(_run(f"{caption.label} "))
    paragraph.append(_field(f" SEQ {caption.label} \\* ARABIC ", caption.new_number))
    paragraph.append(_run(f" {caption.title}"))
    paragraph.append(_bookmark_end(bookmark_id))


def _clear_paragraph_text(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    saved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    paragraph.clear()
    if saved_ppr is not None:
        paragraph.append(saved_ppr)


def _replace_references_in_paragraph(
    paragraph: ET.Element,
    refs_with_captions: list[tuple[dict[str, Any], CaptionCandidate]],
) -> list[tuple[dict[str, Any], CaptionCandidate]]:
    """Replace ALL references in a paragraph in one pass.

    Returns the list of (reference, caption) pairs that were successfully
    replaced.  Processing every reference at once avoids the bug where
    rebuilding the paragraph for one reference destroys the REF field
    created for a previous one.
    """
    original_text = _paragraph_text(paragraph)

    # Locate each reference in the text using progressive search so that
    # duplicate raw_text values match successive occurrences.
    positioned: list[tuple[int, int, dict[str, Any], CaptionCandidate]] = []
    search_start = 0
    for ref, caption in refs_with_captions:
        raw_text = ref["raw_text"]
        start = original_text.find(raw_text, search_start)
        if start >= 0:
            end = start + len(raw_text)
            positioned.append((start, end, ref, caption))
            search_start = end

    if not positioned:
        return []

    # Rebuild the paragraph once with all replacements.
    ppr = paragraph.find("w:pPr", NS)
    saved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    paragraph.clear()
    if saved_ppr is not None:
        paragraph.append(saved_ppr)

    last_end = 0
    replaced: list[tuple[dict[str, Any], CaptionCandidate]] = []
    for start, end, ref, caption in positioned:
        if start > last_end:
            paragraph.append(_run(original_text[last_end:start]))
        if ref["prefix"]:
            paragraph.append(_run(ref["prefix"], color=CROSSREF_COLOR))
        paragraph.append(
            _field(
                f" REF {caption.bookmark_name} \\h ",
                caption.new_number,
                color=CROSSREF_COLOR,
            )
        )
        last_end = end
        replaced.append((ref, caption))

    if last_end < len(original_text):
        paragraph.append(_run(original_text[last_end:]))

    return replaced


def _run(text: str, *, color: str | None = None) -> ET.Element:
    run = ET.Element(f"{W}r")
    if color:
        rpr = ET.SubElement(run, f"{W}rPr")
        color_node = ET.SubElement(rpr, f"{W}color")
        color_node.set(f"{W}val", color)
    text_node = ET.SubElement(run, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def _field(instruction: str, display_text: str, *, color: str | None = None) -> ET.Element:
    field = ET.Element(f"{W}fldSimple")
    field.set(f"{W}instr", instruction)
    field.append(_run(display_text, color=color))
    return field


def _bookmark_start(bookmark_id: int, name: str) -> ET.Element:
    node = ET.Element(f"{W}bookmarkStart")
    node.set(f"{W}id", str(bookmark_id))
    node.set(f"{W}name", name)
    return node


def _bookmark_end(bookmark_id: int) -> ET.Element:
    node = ET.Element(f"{W}bookmarkEnd")
    node.set(f"{W}id", str(bookmark_id))
    return node


def _max_bookmark_id(root: ET.Element) -> int:
    max_id = 0
    for node in root.findall(".//w:bookmarkStart", NS):
        raw_id = node.get(f"{W}id")
        if raw_id and raw_id.isdigit():
            max_id = max(max_id, int(raw_id))
    return max_id


def _serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
