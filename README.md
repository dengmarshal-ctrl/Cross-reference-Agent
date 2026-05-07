# CSR 文档题注与交叉引用治理 Agent Demo

这是一个面向 **医药 CSR 中文报告** 的 Web Demo，用来演示 AI 写作下游 Agent 如何在不依赖 WPS Word 团队底层能力的情况下，先通过 OOXML 直接实现一条最小闭环：

1. 识别正文中的表格引用，例如 `见表格14.1.1.2`。
2. 识别表格上方的静态表题，例如 `表 14.1.1.2 试验完成情况总结...`。
3. 将静态表题转换成带 `SEQ` 字段的题注。
4. 建立旧编号到新编号的映射，例如 `14.1.1.2 -> 1`。
5. 将正文引用改写为新编号，并写入 `REF` 交叉引用字段。
6. 输出处理后 `.docx` 和审计报告。

> 这个项目用于和开发、WPS Word 团队讨论技术边界与底层 Skill 需求，不是生产级 Word 内核替代方案。

## 快速启动

```bash
python3 -m app.server
```

打开：

```text
http://127.0.0.1:8000
```

页面中可以先点击“下载示例文档”，再上传该文档进行扫描和执行治理。

## 运行测试

```bash
python3 -m unittest discover
```

## Demo 能力范围

当前 Demo 聚焦 MVP 中最关键的技术链路：

- `.docx` 文件读取。
- `word/document.xml` 结构解析。
- 表格对象识别。
- 静态中文表题识别。
- 正文表格引用识别。
- 高置信度编号匹配。
- 题注字段写入：`SEQ 表 \* ARABIC`。
- 书签写入：`_CSR_Table_001`。
- 交叉引用字段写入：`REF _CSR_Table_001 \h`。
- 审计 JSON 输出。

## 当前限制

为了保持 Demo 轻量可运行，当前没有引入 FastAPI、python-docx 或商业 Office/WPS SDK。实现采用 Python 标准库直接处理 OOXML，因此存在一些生产限制：

- 只处理表格题注和表格正文引用，暂不处理图片、公式、章节、附录。
- 处理被命中的段落时会重建 runs，复杂局部样式可能丢失。
- 字段显示值会写入文档，但最终字段刷新、兼容性和版式稳定性仍应由 WPS/Office 内核保障。
- 只实现高置信度自动处理，低置信度项仅在计划中报告。

## 与 WPS Word Skill 的关系

这个 Demo 证明业务 Agent 可以在上层完成：

- CSR 规则理解。
- 正文引用识别。
- 旧编号到新编号映射。
- 执行计划生成。
- 审计报告。
- 可视化执行日志。

但生产环境中，以下能力建议由 WPS Word 团队以 CLI/Skill 形式提供：

- `ParseDocumentStructure`
- `ConvertTextToCaption`
- `InsertCaption`
- `PreviewCaptionNumbering`
- `CreateBookmark`
- `ReplaceTextWithCrossReference`
- `UpdateFields`
- `ValidateReferences`
- `ValidateDocument`
- `CreateCheckpoint`
- `RollbackToCheckpoint`

业务 Agent 应调用这些对象级 Skill，而不是长期直接修改 OOXML。
