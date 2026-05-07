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

NS = {"w": WORD_NS}
ET.register_namespace("w", WORD_NS)
ET.register_namespace("r", OFFICE_REL_NS)

W = f"{{{WORD_NS}}}"
CAPTION_RE = re.compile(
    r"^\s*(?P<label>表格|表)\s*(?P<number>\d+(?:[.\-]\d+)*)\s*(?P<title>.+?)\s*$"
)
REFERENCE_RE = re.compile(
    r"(?P<prefix>统计列表|列表|表格|表)(?P<space>\s*)(?P<number>\d+(?:[.\-]\d+)*)"
)


@dataclass
class CaptionCandidate:
    id: str
    table_id: str
    paragraph_id: str | None
    duplicate_paragraph_id: str | None
    source_kind: str
    body_index: int
    table_index: int
    original_text: str | None
    original_number: str | None
    title: str
    new_number: str
    bookmark_name: str

    @property
    def new_label(self) -> str:
        return f"表 {self.new_number}"

    @property
    def display_text(self) -> str:
        return f"{self.new_label} {self.title}".strip()


def analyze_docx(docx_bytes: bytes) -> dict[str, Any]:
    """Return a transparent processing plan without mutating the document."""
    document_xml = _read_document_xml(docx_bytes)
    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    block_records = _collect_blocks(body)
    captions = _build_caption_candidates(block_records)
    references = _detect_references(block_records, captions)

    return _build_plan(block_records, captions, references)


def process_docx(docx_bytes: bytes) -> tuple[bytes, dict[str, Any]]:
    """Normalize table captions and replace matching text with REF fields."""
    document_xml = _read_document_xml(docx_bytes)
    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("word/document.xml does not contain w:body")

    block_records = _collect_blocks(body)
    captions = _build_caption_candidates(block_records)
    references = _detect_references(block_records, captions)
    plan = _build_plan(block_records, captions, references)
    logs: list[dict[str, Any]] = [
        {
            "step": "Step 1",
            "title": "解析文档结构",
            "detail": f"识别到 {len(captions)} 个表格对象，{len(references)} 处正文表格引用。",
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

    # Insert new caption paragraphs only for tables without any detected title.
    # Existing titles are overwritten in-place, including title rows inside tables.
    created_paragraphs: dict[str, ET.Element] = {}
    insert_offset = 0
    for caption in captions:
        if caption.source_kind != "missing":
            continue
        paragraph = _new_paragraph()
        body.insert(caption.body_index + insert_offset, paragraph)
        insert_offset += 1
        created_paragraphs[caption.id] = paragraph

    # Re-collect blocks because inserting paragraphs changes body order.
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
            "title": "创建标准题注",
            "detail": f"将 {len(captions)} 个静态/缺失表题转换为带 SEQ 字段的题注。",
            "status": "done",
        }
    )

    audit_items: list[dict[str, Any]] = []
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
            paragraph = _first_table_cell_paragraph(table_by_id.get(caption.table_id))
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
                "target": caption.table_id,
                "before": before,
                "after": after,
                "action": "ConvertTextToCaption",
                "bookmark": caption.bookmark_name,
            }
        )

    logs.append(
        {
            "step": "Step 3",
            "title": "建立旧编号到新编号映射",
            "detail": f"生成 {len(captions)} 条映射，例如 {captions[0].original_number if captions else '-'} -> {captions[0].new_number if captions else '-'}。",
            "status": "done",
        }
    )

    # References were detected before caption normalization. Apply them to the
    # original paragraph IDs, skipping caption paragraphs.
    current_blocks = _collect_blocks(body)
    current_paragraphs = {
        record["id"]: record["element"]
        for record in current_blocks
        if record["type"] == "paragraph"
    }
    applied_references = 0
    skipped_references = 0
    for reference in references:
        if reference["paragraph_id"] in caption_paragraph_ids:
            skipped_references += 1
            continue
        paragraph = current_paragraphs.get(reference["paragraph_id"])
        target = next(
            (caption for caption in captions if caption.id == reference.get("target_caption_id")),
            None,
        )
        if paragraph is None or target is None or reference["confidence"] != "high":
            skipped_references += 1
            continue

        before = _paragraph_text(paragraph)
        replaced = _replace_reference_in_paragraph(paragraph, reference, target)
        after = _paragraph_text(paragraph)
        if replaced:
            applied_references += 1
            audit_items.append(
                {
                    "type": "cross_reference",
                    "id": reference["id"],
                    "target": target.id,
                    "before": before,
                    "after": after,
                    "action": "ReplaceTextWithCrossReference",
                    "bookmark": target.bookmark_name,
                    "confidence": reference["confidence"],
                }
            )
        else:
            skipped_references += 1

    logs.append(
        {
            "step": "Step 4",
            "title": "创建真实交叉引用",
            "detail": f"创建 {applied_references} 处 REF 字段，跳过 {skipped_references} 处低置信度或无法定位项。",
            "status": "done",
        }
    )

    updated_xml = _serialize_xml(root)
    processed_bytes = _replace_document_xml(docx_bytes, updated_xml)
    audit = {
        "task_id": f"csr-demo-{uuid.uuid4().hex[:8]}",
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "tables": len(captions),
            "caption_actions": len(captions),
            "references_detected": len(references),
            "cross_references_created": applied_references,
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
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
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
    for index, child in enumerate(list(body)):
        if child.tag == f"{W}p":
            paragraph_count += 1
            records.append(
                {
                    "id": f"p_{paragraph_count:04d}",
                    "type": "paragraph",
                    "body_index": index,
                    "text": _paragraph_text(child),
                    "element": child,
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
                }
            )
    return records


def _build_caption_candidates(blocks: list[dict[str, Any]]) -> list[CaptionCandidate]:
    captions: list[CaptionCandidate] = []
    table_number = 0
    for index, record in enumerate(blocks):
        if record["type"] != "table":
            continue
        table_number += 1
        previous = _previous_non_empty_paragraph(blocks, index)
        paragraph_match = CAPTION_RE.match(previous["text"]) if previous else None
        table_caption = _first_row_caption(record["element"])

        # Prefer an external caption paragraph. If absent, use a table-internal
        # title row, which is common in CSR source tables exported from systems.
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
        captions.append(
            CaptionCandidate(
                id=f"cap_tbl_{table_number:03d}",
                table_id=record["id"],
                paragraph_id=paragraph_id,
                duplicate_paragraph_id=duplicate_paragraph_id,
                source_kind=source_kind,
                body_index=record["body_index"],
                table_index=table_number,
                original_text=original_text,
                original_number=original_number,
                title=title,
                new_number=str(table_number),
                bookmark_name=f"_CSR_Table_{table_number:03d}",
            )
        )
    return captions


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
    table_positions = {caption.id: caption.body_index for caption in captions}
    references: list[dict[str, Any]] = []
    ref_index = 0
    for record in blocks:
        if record["type"] != "paragraph" or record["id"] in caption_paragraph_ids:
            continue
        for match in REFERENCE_RE.finditer(record["text"]):
            ref_index += 1
            number = match.group("number")
            prefix = f"{match.group('prefix')}{match.group('space')}"
            target, confidence, reason, match_method = _match_reference_target(
                number=number,
                prefix_label=match.group("prefix"),
                paragraph_body_index=record["body_index"],
                context_text=record["text"],
                old_number_index=by_old_number,
                new_number_index=by_new_number,
                captions=captions,
                table_positions=table_positions,
            )
            references.append(
                {
                    "id": f"ref_{ref_index:03d}",
                    "paragraph_id": record["id"],
                    "raw_text": match.group(0),
                    "prefix": prefix,
                    "original_number": number,
                    "context": record["text"],
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
    return {
        "document_summary": {
            "paragraphs": sum(1 for block in blocks if block["type"] == "paragraph"),
            "tables": len(captions),
            "table_references": len(references),
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
                "table_id": caption.table_id,
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
    return "InsertCaption"


def _match_reference_target(
    *,
    number: str,
    prefix_label: str,
    paragraph_body_index: int,
    context_text: str,
    old_number_index: dict[str, list[CaptionCandidate]],
    new_number_index: dict[str, list[CaptionCandidate]],
    captions: list[CaptionCandidate],
    table_positions: dict[str, int],
) -> tuple[CaptionCandidate | None, str, str, str]:
    if prefix_label in {"表", "表格"}:
        semantic_target, semantic_reason = _semantic_nearby_candidate(
            context_text=context_text,
            paragraph_body_index=paragraph_body_index,
            captions=captions,
            table_positions=table_positions,
        )
        if semantic_target:
            return semantic_target, "high", semantic_reason, "semantic_nearby_context"

    if _is_simple_number(number):
        nearest_caption = _nearest_caption_after(paragraph_body_index, captions, table_positions)
        if nearest_caption and _caption_matches_short_number(nearest_caption, number):
            return (
                nearest_caption,
                "high",
                "短编号引用按当前段落后的最近相关表格匹配",
                "short_number_nearby_caption",
            )

    old_number_targets = old_number_index.get(number, [])
    if old_number_targets:
        target = _best_context_candidate(paragraph_body_index, old_number_targets, table_positions)
        if target:
            if len(old_number_targets) == 1:
                return target, "high", "原始统计表编号唯一匹配到表格题注", "source_number_exact"
            return (
                target,
                "high",
                "存在重复局部编号，按同章节/段落后的最近同编号表格匹配",
                "duplicate_source_number_nearby",
            )

    # Treat generated caption numbers as global only when no source caption uses
    # the same local number. Otherwise `见表1` in every section would incorrectly
    # bind to the first global caption.
    if number not in old_number_index:
        new_number_targets = new_number_index.get(number, [])
        target = _best_context_candidate(paragraph_body_index, new_number_targets, table_positions)
        if target:
            return target, "high", "正文编号匹配题注重排后的新编号", "generated_number_exact"

    nearest_caption = _nearest_caption_after(paragraph_body_index, captions, table_positions)
    if nearest_caption:
        return (
            nearest_caption,
            "medium",
            "未命中编号，按同章节/段落后的最近表格推荐匹配",
            "nearby_fallback",
        )

    return None, "low", "未找到具有相同编号或邻近位置的表格题注", "unmatched"


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
    table_positions: dict[str, int],
) -> tuple[CaptionCandidate | None, str]:
    candidates: list[tuple[float, int, list[str], CaptionCandidate]] = []
    normalized_context = _normalize_for_semantic_match(context_text)
    for caption in captions:
        position = table_positions.get(caption.id, -1)
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
        "正文整句与附近表题语义匹配，关键词："
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
    table_positions: dict[str, int],
) -> CaptionCandidate | None:
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    after_candidates = [
        caption
        for caption in candidates
        if table_positions.get(caption.id, -1) > paragraph_body_index
    ]
    if after_candidates:
        after_candidates.sort(key=lambda caption: table_positions.get(caption.id, 10**9))
        return after_candidates[0]

    candidates.sort(key=lambda caption: abs(table_positions.get(caption.id, 10**9) - paragraph_body_index))
    return candidates[0]


def _nearest_caption_after(
    paragraph_body_index: int,
    captions: list[CaptionCandidate],
    table_positions: dict[str, int],
) -> CaptionCandidate | None:
    candidates = [
        caption
        for caption in captions
        if table_positions.get(caption.id, -1) > paragraph_body_index
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda caption: table_positions.get(caption.id, 10**9))
    nearest = candidates[0]
    distance = table_positions.get(nearest.id, 10**9) - paragraph_body_index
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
    if not match:
        return None
    return {
        "text": text,
        "number": match.group("number"),
        "title": match.group("title"),
    }


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
    paragraph.append(_run("表 "))
    paragraph.append(_field(" SEQ 表 \\* ARABIC ", caption.new_number))
    paragraph.append(_run(f" {caption.title}"))
    paragraph.append(_bookmark_end(bookmark_id))


def _clear_paragraph_text(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    saved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    paragraph.clear()
    if saved_ppr is not None:
        paragraph.append(saved_ppr)


def _replace_reference_in_paragraph(
    paragraph: ET.Element, reference: dict[str, Any], caption: CaptionCandidate
) -> bool:
    original_text = _paragraph_text(paragraph)
    raw_text = reference["raw_text"]
    start = original_text.find(raw_text)
    if start < 0:
        return False
    end = start + len(raw_text)
    before_text = original_text[:start]
    after_text = original_text[end:]

    ppr = paragraph.find("w:pPr", NS)
    saved_ppr = copy.deepcopy(ppr) if ppr is not None else None
    paragraph.clear()
    if saved_ppr is not None:
        paragraph.append(saved_ppr)
    if before_text:
        paragraph.append(_run(before_text))
    if reference["prefix"]:
        paragraph.append(_run(reference["prefix"]))
    paragraph.append(
        _field(
            f" REF {caption.bookmark_name} \\h ",
            caption.new_number,
        )
    )
    if after_text:
        paragraph.append(_run(after_text))
    return True


def _run(text: str) -> ET.Element:
    run = ET.Element(f"{W}r")
    text_node = ET.SubElement(run, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def _field(instruction: str, display_text: str) -> ET.Element:
    field = ET.Element(f"{W}fldSimple")
    field.set(f"{W}instr", instruction)
    field.append(_run(display_text))
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

