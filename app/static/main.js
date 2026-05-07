const state = {
  file: null,
  plan: null,
  processedDoc: null,
  audit: null,
  confirmedRefs: new Set(),
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  fileName: document.querySelector("#fileName"),
  sampleBtn: document.querySelector("#sampleBtn"),
  analyzeBtn: document.querySelector("#analyzeBtn"),
  processBtn: document.querySelector("#processBtn"),
  clearBtn: document.querySelector("#clearBtn"),
  docTitle: document.querySelector("#docTitle"),
  summaryBadge: document.querySelector("#summaryBadge"),
  emptyState: document.querySelector("#emptyState"),
  planView: document.querySelector("#planView"),
  metricTables: document.querySelector("#metricTables"),
  metricImages: document.querySelector("#metricImages"),
  metricRefs: document.querySelector("#metricRefs"),
  metricHigh: document.querySelector("#metricHigh"),
  metricMedium: document.querySelector("#metricMedium"),
  metricLow: document.querySelector("#metricLow"),
  captionList: document.querySelector("#captionList"),
  referenceList: document.querySelector("#referenceList"),
  confirmSection: document.querySelector("#confirmSection"),
  confirmList: document.querySelector("#confirmList"),
  confirmCount: document.querySelector("#confirmCount"),
  confirmAllBtn: document.querySelector("#confirmAllBtn"),
  logList: document.querySelector("#logList"),
  downloadPanel: document.querySelector("#downloadPanel"),
  resultSummary: document.querySelector("#resultSummary"),
  downloadDocBtn: document.querySelector("#downloadDocBtn"),
  downloadAuditBtn: document.querySelector("#downloadAuditBtn"),
};

function getStrategy() {
  const checked = document.querySelector('input[name="strategy"]:checked');
  return checked ? checked.value : "continuous";
}

els.fileInput.addEventListener("change", () => {
  state.file = els.fileInput.files[0] || null;
  state.plan = null;
  state.processedDoc = null;
  state.audit = null;
  state.confirmedRefs.clear();
  els.fileName.textContent = state.file ? state.file.name : "尚未选择文件";
  els.docTitle.textContent = state.file ? state.file.name : "等待上传文档";
  els.processBtn.disabled = true;
  els.downloadPanel.classList.add("hidden");
  resetPlan();
  log("已选择文档", state.file ? `文件：${state.file.name}` : "未选择文件");
});

els.sampleBtn.addEventListener("click", async () => {
  const response = await fetch("/api/sample");
  const blob = await response.blob();
  downloadBlob(blob, "csr-crossref-sample.docx");
  log("已生成示例文档", "请下载后重新上传该 .docx，体验完整扫描和治理流程。");
});

els.analyzeBtn.addEventListener("click", async () => {
  if (!ensureFile()) return;
  const strategy = getStrategy();
  const strategyLabel = { continuous: "全文连续编号", chapter: "按章节编号", preserve: "保留原始编号" }[strategy];
  log("Step 1：上传并解析文档", `策略：${strategyLabel}；正在读取 word/document.xml 并构建对象注册表。`);
  const response = await fetch(`/api/analyze?strategy=${strategy}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: state.file,
  });
  const data = await readJsonResponse(response);
  state.plan = data;
  state.confirmedRefs.clear();
  renderPlan(data);
  els.processBtn.disabled = false;
  const s = data.document_summary;
  log(
    "Step 2：生成处理计划",
    `发现 ${s.tables} 个表格、${s.images} 个图片、${s.table_references} 处表格引用、${s.image_references} 处图片引用。`
  );
});

els.processBtn.addEventListener("click", async () => {
  if (!ensureFile()) return;
  const strategy = getStrategy();
  const confirmedParam = [...state.confirmedRefs].join(",");
  log("Step 3：执行题注与交叉引用创建", "正在写入 SEQ 题注字段、书签和 REF 交叉引用字段。");
  const url = `/api/process?filename=${encodeURIComponent(state.file.name)}&strategy=${strategy}&confirmed=${encodeURIComponent(confirmedParam)}`;
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: state.file,
  });
  const data = await readJsonResponse(response);
  state.processedDoc = {
    filename: data.filename,
    blob: base64ToBlob(
      data.docx_base64,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ),
  };
  state.audit = {
    filename: data.audit_filename,
    blob: base64ToBlob(data.audit_base64, "application/json"),
    raw: data.audit,
  };
  renderExecution(data.audit);
});

els.downloadDocBtn.addEventListener("click", () => {
  if (state.processedDoc) downloadBlob(state.processedDoc.blob, state.processedDoc.filename);
});

els.downloadAuditBtn.addEventListener("click", () => {
  if (state.audit) downloadBlob(state.audit.blob, state.audit.filename);
});

els.clearBtn.addEventListener("click", () => {
  els.logList.innerHTML = "";
});

els.confirmAllBtn.addEventListener("click", () => {
  if (!state.plan) return;
  for (const ref of state.plan.reference_actions) {
    if (ref.confidence === "medium" && ref.target_caption_id) {
      state.confirmedRefs.add(ref.id);
    }
  }
  renderConfirmQueue(state.plan.reference_actions);
});

function ensureFile() {
  if (state.file) return true;
  log("缺少文档", "请先选择一个 .docx 文件，或下载示例文档后上传。");
  return false;
}

async function readJsonResponse(response) {
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "请求失败");
  }
  return data;
}

function renderPlan(plan) {
  const summary = plan.document_summary;
  els.emptyState.classList.add("hidden");
  els.planView.classList.remove("hidden");
  els.summaryBadge.textContent = "扫描完成";
  els.metricTables.textContent = summary.tables;
  els.metricImages.textContent = summary.images;
  els.metricRefs.textContent = summary.table_references + summary.image_references;
  els.metricHigh.textContent = summary.high_confidence_references;
  els.metricMedium.textContent = summary.medium_confidence_references || 0;
  els.metricLow.textContent = summary.low_confidence_references;

  els.captionList.innerHTML = "";
  for (const action of plan.caption_actions) {
    const typeIcon = action.object_type === "image" ? "[图]" : "[表]";
    els.captionList.appendChild(
      item(
        `${typeIcon} ${action.original_text || "无静态题注"} → ${action.display_text}`,
        `动作：${action.action}；对象：${action.object_id}；类型：${action.object_type}；书签：${action.bookmark}`
      )
    );
  }

  els.referenceList.innerHTML = "";
  for (const ref of plan.reference_actions) {
    const typeIcon = ref.object_type === "image" ? "[图]" : "[表]";
    const confidenceClass =
      ref.confidence === "high"
        ? "confidence-high"
        : ref.confidence === "medium"
          ? "confidence-medium"
          : "confidence-low";
    els.referenceList.appendChild(
      item(
        `${typeIcon} ${ref.raw_text} → ${ref.proposed_text || "未匹配"}`,
        `段落：${ref.paragraph_id}；目标：${ref.target_title || "无"}；匹配方法：${formatMatchMethod(ref.match_method)}；原因：${ref.reason}`,
        ref.confidence,
        confidenceClass
      )
    );
  }

  renderConfirmQueue(plan.reference_actions);
}

function renderConfirmQueue(references) {
  const mediumRefs = references.filter(r => r.confidence === "medium" && r.target_caption_id);
  if (mediumRefs.length === 0) {
    els.confirmSection.classList.add("hidden");
    return;
  }
  els.confirmSection.classList.remove("hidden");
  els.confirmList.innerHTML = "";

  for (const ref of mediumRefs) {
    const isConfirmed = state.confirmedRefs.has(ref.id);
    const wrapper = document.createElement("div");
    wrapper.className = "confirm-item" + (isConfirmed ? " confirmed" : "");

    const header = document.createElement("div");
    header.className = "confirm-item-head";
    header.innerHTML = `
      <div>
        <strong>原文引用：</strong><span>${ref.raw_text}</span>
        <br/><strong>上下文：</strong><span class="context-text">${ref.context}</span>
      </div>
    `;

    const body = document.createElement("div");
    body.className = "confirm-item-body";
    body.innerHTML = `
      <div><strong>推荐目标：</strong>${ref.proposed_text || "无"} ${ref.target_title || ""}</div>
      <div><strong>匹配方法：</strong>${formatMatchMethod(ref.match_method)}</div>
      <div><strong>推荐原因：</strong>${ref.reason}</div>
    `;

    const actions = document.createElement("div");
    actions.className = "confirm-item-actions";

    if (isConfirmed) {
      const badge = document.createElement("span");
      badge.className = "confirm-badge confirmed-badge";
      badge.textContent = "已确认";
      actions.appendChild(badge);

      const undoBtn = document.createElement("button");
      undoBtn.className = "ghost small";
      undoBtn.textContent = "撤销";
      undoBtn.addEventListener("click", () => {
        state.confirmedRefs.delete(ref.id);
        renderConfirmQueue(references);
      });
      actions.appendChild(undoBtn);
    } else {
      const acceptBtn = document.createElement("button");
      acceptBtn.className = "small";
      acceptBtn.textContent = "接受";
      acceptBtn.addEventListener("click", () => {
        state.confirmedRefs.add(ref.id);
        renderConfirmQueue(references);
      });
      actions.appendChild(acceptBtn);

      const skipBtn = document.createElement("button");
      skipBtn.className = "secondary small";
      skipBtn.textContent = "跳过";
      skipBtn.addEventListener("click", () => {
        state.confirmedRefs.delete(ref.id);
        wrapper.style.opacity = "0.4";
      });
      actions.appendChild(skipBtn);
    }

    wrapper.appendChild(header);
    wrapper.appendChild(body);
    wrapper.appendChild(actions);
    els.confirmList.appendChild(wrapper);
  }

  els.confirmCount.textContent = `已确认 ${state.confirmedRefs.size} 项`;
}

function formatMatchMethod(method) {
  const labels = {
    semantic_nearby_context: "正文整句语义 + 附近题注",
    short_number_nearby_caption: "短编号 + 段落邻近",
    source_number_exact: "原始编号精确匹配",
    duplicate_source_number_nearby: "重复局部编号 + 附近同编号对象",
    generated_number_exact: "重排后新编号精确匹配",
    nearby_fallback: "段落邻近兜底推荐",
    unmatched: "未匹配",
  };
  return labels[method] || method || "未记录";
}

function renderExecution(audit) {
  for (const entry of audit.logs) {
    log(`${entry.step}：${entry.title}`, entry.detail);
  }
  const summary = audit.summary;
  els.downloadPanel.classList.remove("hidden");
  els.resultSummary.textContent =
    `已创建 ${summary.caption_actions} 个题注（${summary.tables} 表 + ${summary.images} 图），` +
    `${summary.cross_references_created} 处真实交叉引用` +
    (summary.user_confirmed > 0 ? `（含 ${summary.user_confirmed} 处用户确认项）` : "") +
    `，跳过 ${summary.skipped_references} 处。`;
  els.summaryBadge.textContent = "治理完成";
}

function resetPlan() {
  els.emptyState.classList.remove("hidden");
  els.planView.classList.add("hidden");
  els.summaryBadge.textContent = "未扫描";
  els.captionList.innerHTML = "";
  els.referenceList.innerHTML = "";
  els.confirmSection.classList.add("hidden");
  els.confirmList.innerHTML = "";
  state.confirmedRefs.clear();
}

function item(title, detail, badgeText, badgeClass) {
  const wrapper = document.createElement("div");
  wrapper.className = "item";

  const head = document.createElement("div");
  head.className = "item-head";
  const titleNode = document.createElement("span");
  titleNode.textContent = title;
  head.appendChild(titleNode);
  if (badgeText) {
    const badge = document.createElement("span");
    badge.className = badgeClass || "";
    badge.textContent = badgeText;
    head.appendChild(badge);
  }

  const small = document.createElement("small");
  small.textContent = detail;
  wrapper.appendChild(head);
  wrapper.appendChild(small);
  return wrapper;
}

function log(title, detail) {
  const wrapper = document.createElement("div");
  wrapper.className = "log";
  const strong = document.createElement("strong");
  strong.textContent = title;
  const span = document.createElement("span");
  span.textContent = detail;
  wrapper.appendChild(strong);
  wrapper.appendChild(span);
  els.logList.prepend(wrapper);
}

function base64ToBlob(base64, type) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
