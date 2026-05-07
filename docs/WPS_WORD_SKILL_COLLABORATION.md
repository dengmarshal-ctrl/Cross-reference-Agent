# CSR 文档治理 Agent：企业事业部与 WPS Word 团队分工及 Skill 能力需求

## 1. 背景

CSR 文档题注与交叉引用治理 Agent 位于 AI 写作平台下游，目标是将 AI 生成的 CSR 中文报告初稿治理为具备标准题注、真实交叉引用、编号一致性和审计追踪的可交付文档。

从落地角度看，该 Agent 不能只靠业务层直接修改 OOXML。业务层可以完成 CSR 规则、语义理解、匹配策略、执行编排和审计，但真正的文档对象操作应尽量由 WPS Word 团队提供稳定底层能力。

核心协作原则：

> WPS Word 团队提供稳定、可编排、可验证的文档对象级原子能力；企业事业部团队在上层完成业务理解、规则编排、Agent 决策、用户确认和审计报告。

## 2. 总体技术分层

```text
AI 写作平台 / 用户界面
        ↓
业务 Agent 编排层：企业事业部负责
        ↓
WPS Word Skill / CLI 能力层：WPS Word 团队负责
        ↓
WPS 文档内核 / OOXML / 渲染 / 域更新能力
```

## 3. 两队职责边界

| 能力 | 企业事业部 Agent 团队 | WPS Word 团队 |
| --- | --- | --- |
| CSR 业务规则 | 负责 | 不负责 |
| 医药文档语义理解 | 负责 | 不负责 |
| 表/图/引用识别策略 | 负责 | 可提供基础对象解析 |
| 引用匹配策略 | 负责 | 不负责 |
| 置信度判断 | 负责 | 不负责 |
| 用户确认队列 | 负责 | 不负责 |
| 审计报告业务内容 | 负责 | 提供底层操作日志 |
| 文档结构解析 | 使用 | 负责提供 |
| 标准题注创建 | 调用 | 负责实现 |
| 书签/锚点创建 | 调用 | 负责实现 |
| 真实交叉引用字段 | 调用 | 负责实现 |
| 域刷新 | 调用 | 负责实现 |
| 文档完整性校验 | 调用 | 负责实现 |
| 文档回滚能力 | 编排 | 提供底层能力 |
| 版式/样式稳定性 | 提需求并验收 | 负责保障 |
| WPS/Office 兼容性 | 提需求并验收 | 负责保障 |

## 4. 企业事业部 Agent 团队负责的能力

### 4.1 CSR Policy Engine

负责 CSR 业务规则：

- 中文 CSR 题注规则
- 表格/图片命名规范
- 表题位置规则
- 表格内题名识别规则
- 编号策略
- 高中低置信度阈值
- 哪些引用可以自动改
- 哪些引用必须人工确认

### 4.2 Document Understanding Layer

负责理解 WPS Word 返回的文档结构：

- 章节层级
- 段落上下文
- 表格/图片位置
- 表题/图题候选
- 正文引用候选
- 表格内部首行题名
- 已有题注和交叉引用

### 4.3 Reference Understanding Engine

负责识别正文引用：

```text
见表格14.1.1.2
如表 2 所示
详见下表
参考图 3
参见统计列表 16.2.2
见表1
```

输出结构化结果：

```json
{
  "reference_id": "ref_001",
  "paragraph_id": "p_102",
  "raw_text": "见表格14.1.1.2",
  "type": "table",
  "original_number": "14.1.1.2",
  "context": "基于所有入组人群总结的患者分布见表格14.1.1.2。",
  "confidence": 0.94
}
```

### 4.4 Matching Engine

负责把正文引用和题注对象匹配。

匹配依据：

- 原编号
- 新编号
- 题注标题
- 正文整句语义
- 表格位置
- 所在章节
- 段落距离
- 重复局部编号
- 目标唯一性

匹配方法需要可解释：

| 匹配方法 | 说明 |
| --- | --- |
| source_number_exact | 原始编号精确匹配 |
| generated_number_exact | 重排后新编号精确匹配 |
| duplicate_source_number_nearby | 重复局部编号 + 附近同编号表 |
| short_number_nearby_caption | 短编号 + 段落邻近 |
| semantic_nearby_context | 正文整句语义 + 附近表题 |
| nearby_fallback | 段落邻近兜底推荐 |
| unmatched | 未匹配 |

### 4.5 Semantic Matching Layer

负责处理不能只靠编号解决的场景。

示例：

正文：

```text
研究期间，在意向治疗分析集（ITT）中，共107例患者发生了重要方案偏离，其中JS001组54例（24.7%），安慰剂组53例（24.0%），两组发生率相近（见表4）。
```

附近表题：

```text
表 3 重要方案偏离情况 － 意向治疗分析集(ITT)
```

业务 Agent 应判断：

- 正文整句中包含“重要方案偏离”。
- 附近表题也包含“重要方案偏离情况”。
- 两者语义高度一致。
- 即使正文写的是“表4”，也应优先匹配当前附近的“重要方案偏离情况”表。

### 4.6 Execution Planner

负责生成执行计划，而不是直接操作文档。

示例：

```json
{
  "caption_actions": [
    {
      "action": "convert_text_to_caption",
      "object_id": "tbl_003",
      "source_paragraph_id": "p_205",
      "caption_title": "试验完成情况总结 意向治疗分析集(ITT)"
    }
  ],
  "crossref_actions": [
    {
      "action": "replace_text_with_cross_reference",
      "paragraph_id": "p_102",
      "source_text": "表格14.1.1.2",
      "target_caption_id": "cap_tbl_003",
      "new_display": "表格3",
      "match_method": "source_number_exact",
      "reason": "原始统计表编号唯一匹配到表格题注"
    }
  ]
}
```

### 4.7 Confirmation Queue

负责承载中低置信度项。

用户可选择：

- 接受推荐
- 更换目标
- 跳过
- 标记无需处理

### 4.8 Audit & Trace System

负责记录：

- Agent 为什么这么判断
- 匹配方法是什么
- 匹配原因是什么
- 调用了哪些 WPS Skills
- 修改前是什么
- 修改后是什么
- 用户确认了什么
- 哪些项失败
- 哪些项跳过

## 5. WPS Word 团队需要提供的能力

业务 Agent 不应长期直接修改 OOXML。WPS Word 团队应提供文档对象级 CLI/Skill 能力，供 Agent 安全调用。

### 5.1 文档结构解析 Skill

#### Skill 名称建议

```text
ParseDocumentStructure
```

#### 作用

把 `.docx` / WPS 文档解析成结构化 JSON，供 Agent 理解。

#### WPS 需要返回

- 段落列表
- 表格列表
- 图片列表
- 标题层级
- 页码/位置
- 样式信息
- 已有题注
- 已有书签
- 已有交叉引用
- 域代码
- 超链接
- 对象前后上下文
- 表格内部首行/合并单元格文本
- 可操作对象 ID

#### 示例返回

```json
{
  "document_id": "doc_001",
  "paragraphs": [
    {
      "id": "p_102",
      "text": "基于所有入组人群总结的患者分布见表格14.1.1.2。",
      "page": 12,
      "style": "正文",
      "section": "10.1"
    }
  ],
  "tables": [
    {
      "id": "tbl_003",
      "page": 13,
      "section": "10.1",
      "nearby_text_before": "表 14.1.1.2 试验完成情况总结 意向治疗分析集(ITT)",
      "first_row_text": "",
      "caption_id": null
    }
  ]
}
```

#### 关键要求

WPS 不能只返回纯文本，要返回可操作对象 ID。否则 Agent 知道“这里有表”，但无法精准操作它。

### 5.2 题注创建 / 题注标准化 Skill

#### Skill 名称建议

```text
InsertCaption
NormalizeCaption
ConvertTextToCaption
ConvertTableRowTitleToCaption
```

#### 作用

把普通文本题注转成真正的 WPS/Office 标准题注。

#### 需要支持

- 表格题注
- 图片题注
- 中文标签：表、图
- 保留原题注主体文字
- 自动编号
- 指定编号策略
- 指定题注位置：对象上方/下方/原位覆盖
- 复用已有静态标题
- 识别并处理表格内部首行题名
- 避免重复插入题注

#### 示例调用

```json
{
  "object_id": "tbl_003",
  "object_type": "table",
  "caption_label": "表",
  "caption_title": "试验完成情况总结 意向治疗分析集(ITT)",
  "numbering_policy": "continuous",
  "position": "above",
  "reuse_existing_text_paragraph_id": "p_205"
}
```

#### 示例返回

```json
{
  "caption_id": "cap_tbl_003",
  "bookmark_id": "bm_tbl_003",
  "display_text": "表 3 试验完成情况总结 意向治疗分析集(ITT)",
  "number": "3"
}
```

### 5.3 编号策略 Skill

#### Skill 名称建议

```text
ApplyCaptionNumberingPolicy
GetCaptionNumberingPolicy
PreviewCaptionNumbering
```

#### 作用

让 Agent 可以按用户选择的规则管理题注编号。

#### 第一版至少支持

1. 全文连续编号

```text
表 1、表 2、表 3
```

2. 按章节编号

```text
表 10-1、表 10-2
```

3. 沿用文档默认编号规则

```text
使用当前 WPS/Office 题注设置
```

#### 关键能力

WPS 最好提供预览编号接口：

```text
不真正修改文档，只告诉 Agent 如果执行后每个表/图会变成什么编号。
```

这样业务 Agent 才能提前展示：

```text
表 14.1.1.2 → 表 3
表 14.1.1.3 → 表 4
```

### 5.4 书签 / 锚点 Skill

#### Skill 名称建议

```text
CreateBookmark
GetBookmark
EnsureObjectAnchor
```

#### 作用

为表格、图片、题注创建稳定可引用的锚点。

#### 要求

- 每个题注对象有唯一 ID。
- 每个表格/图片可以被精确定位。
- 重复执行时要幂等，不能创建重复书签。
- 书签名规则稳定。
- 支持查询已有书签。

#### 示例调用

```json
{
  "target_id": "cap_tbl_003",
  "bookmark_name": "_CSR_Table_003"
}
```

#### 示例返回

```json
{
  "bookmark_id": "bm_003",
  "bookmark_name": "_CSR_Table_003"
}
```

### 5.5 交叉引用创建 Skill

#### Skill 名称建议

```text
CreateCrossReference
ReplaceTextWithCrossReference
UpdateCrossReference
ValidateCrossReference
```

#### 作用

把正文中的普通文本引用替换为真正交叉引用字段。

#### 示例

原文：

```text
基于所有入组人群总结的患者分布见表格14.1.1.2。
```

Agent 判断后调用 WPS：

```json
{
  "paragraph_id": "p_102",
  "text_range": {
    "start": 18,
    "end": 28
  },
  "replacement_prefix": "表格",
  "target_caption_id": "cap_tbl_003",
  "reference_display": "number_only"
}
```

期望结果：

```text
基于所有入组人群总结的患者分布见表格3。
```

其中“3”或“表格3”是真实交叉引用字段。

#### 需要支持的展示模式

```text
只显示编号：3
显示标签+编号：表 3
显示完整题注：表 3 试验完成情况总结...
```

MVP 建议默认：

- 正文保留“表格/表/图”等原措辞。
- 交叉引用字段主要绑定编号部分。

### 5.6 域刷新与文档校验 Skill

#### Skill 名称建议

```text
UpdateFields
ValidateDocument
ValidateReferences
OpenCheck
RepairDocument
```

#### 作用

完成处理后，确保文档真的可用。

#### 必须支持

- 刷新所有题注编号。
- 刷新所有交叉引用。
- 校验引用是否断链。
- 校验书签是否存在。
- 校验文档是否可打开。
- 校验是否有损坏字段。
- 校验是否有乱码或 XML 错误。

#### 返回示例

```json
{
  "valid": true,
  "field_update_count": 86,
  "broken_reference_count": 0,
  "warning_count": 2,
  "warnings": [
    {
      "type": "unresolved_reference",
      "paragraph_id": "p_332",
      "message": "未找到表 7 对应目标"
    }
  ]
}
```

### 5.7 差异与回滚 Skill

#### Skill 名称建议

```text
CreateCheckpoint
GetDocumentDiff
RollbackToCheckpoint
ExportOperationLog
```

#### 作用

保证 Agent 执行不是黑盒、不是不可逆。

#### 需要能力

- 修改前创建 checkpoint。
- 记录每次 Skill 调用。
- 输出修改前后差异。
- 失败时回滚。
- 允许用户下载原文、处理后文档、审计报告。

## 6. 推荐 CLI 形态

业务 Agent 最好通过 CLI 或 HTTP Skill 调用。

CLI 形态示例：

```bash
wps-word-skill parse --input report.docx --output structure.json
wps-word-skill caption normalize --input report.docx --plan caption_plan.json --output step1.docx
wps-word-skill crossref apply --input step1.docx --plan crossref_plan.json --output final.docx
wps-word-skill validate --input final.docx --output validate_report.json
```

## 7. 最小可落地 Skill 清单

### 7.1 P0：必须提供

1. `ParseDocumentStructure`
2. `ConvertTextToCaption`
3. `ConvertTableRowTitleToCaption`
4. `InsertCaption`
5. `CreateBookmark`
6. `ReplaceTextWithCrossReference`
7. `UpdateFields`
8. `ValidateDocument`
9. `ValidateReferences`
10. `CreateCheckpoint`
11. `RollbackToCheckpoint`

这些能力足够支撑 MVP 闭环。

### 7.2 P1：增强能力

1. `PreviewCaptionNumbering`
2. `GetDocumentDiff`
3. `DetectExistingCaption`
4. `DetectExistingCrossReference`
5. `RepairBrokenCrossReference`
6. `ExportOperationLog`
7. `NormalizeCaptionStyle`
8. `MoveCaptionOutsideTable`

### 7.3 P2：后续能力

1. 图表清单自动生成
2. 附录表编号规则
3. 公式题注与交叉引用
4. 章节交叉引用
5. 文献引用联动
6. 批量文档治理

## 8. 推荐落地流程

```text
Step 1：Agent 调用 WPS 解析文档
        ↓
Step 2：Agent 识别疑似题注和正文引用
        ↓
Step 3：Agent 生成题注治理计划
        ↓
Step 4：WPS 预览题注编号结果
        ↓
Step 5：Agent 建立旧编号 → 新编号映射
        ↓
Step 6：Agent 结合编号、章节/段落、正文整句语义生成交叉引用计划
        ↓
Step 7：用户确认处理计划
        ↓
Step 8：Agent 调用 WPS 执行题注创建
        ↓
Step 9：Agent 调用 WPS 执行交叉引用创建
        ↓
Step 10：WPS 刷新域并校验文档
        ↓
Step 11：Agent 输出处理结果和审计报告
```

## 9. 关键数据结构建议

### 9.1 文档对象注册表

```json
{
  "objects": [
    {
      "id": "tbl_003",
      "type": "table",
      "section": "1.2",
      "page": 28,
      "body_index": 185,
      "source_caption": "表1 重要方案偏离情况 － 意向治疗分析集(ITT)",
      "normalized_caption": "表 3 重要方案偏离情况 － 意向治疗分析集(ITT)",
      "caption_id": "cap_tbl_003",
      "bookmark_id": "bm_tbl_003"
    }
  ]
}
```

### 9.2 引用匹配结果

```json
{
  "reference_id": "ref_012",
  "paragraph_id": "p_128",
  "raw_text": "见表4",
  "context": "研究期间，在意向治疗分析集（ITT）中，共107例患者发生了重要方案偏离，两组发生率相近（见表4）。",
  "target_caption_id": "cap_tbl_003",
  "target_title": "重要方案偏离情况 － 意向治疗分析集(ITT)",
  "proposed_text": "见表3",
  "confidence": "high",
  "match_method": "semantic_nearby_context",
  "reason": "正文整句与附近表题语义匹配，关键词：重要方案偏离；段落距离：2"
}
```

### 9.3 执行动作

```json
{
  "action": "replace_text_with_cross_reference",
  "paragraph_id": "p_128",
  "source_text": "表4",
  "replacement_prefix": "表",
  "target_caption_id": "cap_tbl_003",
  "display_mode": "number_only",
  "expected_display_text": "表3"
}
```

## 10. WPS Word 团队需重点保障的非功能要求

### 10.1 幂等性

同一份执行计划重复执行，不应重复创建题注、书签、交叉引用。

### 10.2 样式保真

题注创建和引用替换不应破坏：

- 字体
- 字号
- 加粗/斜体
- 段落缩进
- 表格边框
- 页眉页脚
- 目录
- 图表清单

### 10.3 域兼容性

创建的字段需要：

- WPS 可识别
- Office Word 可识别
- 刷新后编号一致
- 点击可跳转

### 10.4 可校验

每次执行后需要返回结构化校验结果：

- 文档是否可打开
- 字段是否可刷新
- 交叉引用是否断链
- 书签是否缺失
- 题注编号是否连续

### 10.5 可回滚

处理失败时必须支持回滚到 checkpoint。

## 11. 当前 OOXML Demo 与生产 Skill 的关系

当前 Demo 证明业务 Agent 可以在上层完成：

- CSR 规则理解
- 正文引用识别
- 旧编号到新编号映射
- 章节/段落上下文匹配
- 轻量语义匹配
- 执行计划生成
- 审计报告
- 可视化执行日志

但 Demo 直接修改 OOXML，只适合作为技术沟通原型。生产环境中，以下能力不建议由业务层长期自研：

- 标准题注真实创建
- 表格内题名转标准题注
- 字段刷新
- 交叉引用字段创建
- WPS/Office 兼容性
- 复杂样式保真
- 文档损坏修复
- 回滚

这些应由 WPS Word 团队以 CLI/Skill 形式提供。

## 12. 向 WPS Word 团队提出的需求摘要

可以这样对 WPS Word 团队提需求：

> 我们需要建设一个 AI 写作下游的 CSR 文档合规 Agent。业务团队负责识别医药 CSR 报告中的表格、图片、题注和正文引用关系，并负责编号、章节/段落上下文、语义匹配和审计逻辑。但我们需要 WPS Word 团队提供稳定的文档原子操作能力，包括文档结构解析、标准题注创建、表格内题名转题注、书签锚点创建、真实交叉引用字段创建、域刷新、文档校验、差异追踪和回滚。
>
> 我们不希望在业务层直接修改 OOXML，而是希望通过 WPS 提供的 CLI/Skills，把业务 Agent 的处理计划安全地落到文档对象上，确保文档可打开、可跳转、可刷新、可审计。

## 13. 最关键的协作原则

对 WPS Word 团队需要明确：

> 请提供“文档对象级”的可编排 Skills，而不是只提供文本替换能力。

这个 Agent 的本质不是查找替换，而是文档对象治理。

必须能操作：

- 表格对象
- 图片对象
- 题注对象
- 书签对象
- 域对象
- 交叉引用对象

如果 WPS 只提供“替换文本”，最终做出来的只是伪交叉引用，不具备真正合规价值。

