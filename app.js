(function () {
  "use strict";

  const E = window.CourierEngine;
  const state = { settings: loadSettings(), records: [], issues: [], files: [], result: null };
  const $ = (selector) => document.querySelector(selector);
  const els = {
    dropZone: $("#dropZone"), fileInput: $("#fileInput"), dashboard: $("#dashboard"), fileStatus: $("#fileStatus"),
    totalCost: $("#totalCost"), topUpCost: $("#topUpCost"), labelDeduction: $("#labelDeduction"), totalCount: $("#totalCount"), averageCost: $("#averageCost"), rangeTitle: $("#rangeTitle"),
    costCaption: $("#costCaption"), monthCount: $("#monthCount"), monthList: $("#monthList"), sourceSummary: $("#sourceSummary"),
    sourceDetails: $("#sourceDetails"), exportButton: $("#exportButton"), settingsButton: $("#settingsButton"),
    settingsDialog: $("#settingsDialog"), settingsForm: $("#settingsForm"), closeSettings: $("#closeSettings"),
    resetFormula: $("#resetFormula"), resetSettings: $("#resetSettings"), formulaSection: $("#formulaSection"),
    formulaPreview: $("#formulaPreview"), toast: $("#toast")
  };

  function loadSettings() {
    try { return E.normalizeSettings(JSON.parse(localStorage.getItem("courier-cost-settings") || "null")); }
    catch (_) { return E.normalizeSettings(E.DEFAULT_SETTINGS); }
  }

  function money(value) { return `¥${Number(value || 0).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
  function decimal(value, digits = 2) { return Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: digits }); }
  function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char])); }
  let toastTimer;
  function toast(message) {
    els.toast.textContent = message;
    els.toast.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove("show"), 2600);
  }

  async function readFile(file) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true, raw: true });
    const records = [];
    const issues = [];
    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
      const extracted = E.extractRowsFromSheet(rows, sheetName, file.name);
      records.push(...extracted.records);
      issues.push(...extracted.issues);
    });
    return { records, issues };
  }

  async function handleFiles(fileList) {
    const files = [...fileList].filter((file) => /\.(xls|xlsx|csv)$/i.test(file.name));
    if (!files.length) return toast("请选择 .xls、.xlsx 或 .csv 表格");
    els.fileStatus.textContent = "正在读取…";
    els.fileStatus.classList.remove("ready");
    try {
      const batches = await Promise.all(files.map(readFile));
      const all = batches.flatMap((batch) => batch.records);
      const deduped = E.deduplicate(all);
      state.records = deduped.records;
      state.issues = batches.flatMap((batch) => batch.issues);
      if (deduped.duplicates.length) state.issues.push({ type: "duplicate", message: `已跳过 ${deduped.duplicates.length} 条重复运单` });
      state.files = files.map((file) => file.name);
      if (!state.records.length) throw new Error("没有找到可计算的寄件记录，请检查列名和数据内容");
      recalculate();
      els.fileStatus.textContent = `已读取 ${state.records.length.toLocaleString("zh-CN")} 件`;
      els.fileStatus.classList.add("ready");
      els.dashboard.hidden = false;
      els.dashboard.scrollIntoView({ behavior: "smooth", block: "start" });
      toast(`已完成 ${state.records.length.toLocaleString("zh-CN")} 件寄件计算`);
      if (state.settings.autoExport) setTimeout(exportWorkbook, 250);
    } catch (error) {
      els.fileStatus.textContent = "导入失败";
      toast(error.message || "表格读取失败");
    } finally {
      els.fileInput.value = "";
    }
  }

  function recalculate() {
    state.result = E.summarize(state.records, state.settings);
    render();
  }

  function render() {
    const { overall, months, formula } = state.result;
    els.totalCost.textContent = money(overall.totalCost);
    els.topUpCost.textContent = money(overall.topUpCost);
    els.labelDeduction.textContent = `已扣除面单 ${money(overall.labelTotal)}（${money(state.settings.labelPrice)}/个）`;
    els.totalCount.textContent = overall.count.toLocaleString("zh-CN");
    els.averageCost.textContent = money(overall.averageCost);
    els.monthCount.textContent = `${months.length} 个月`;
    els.rangeTitle.textContent = months.length === 1 ? months[0].label : `${months[0].label} – ${months[months.length - 1].label}`;
    els.costCaption.textContent = state.settings.rounding === "ceil" ? "续重按 1kg 向上进位" : state.settings.rounding === "exact" ? "续重按实际重量" : "按自定义公式";

    els.monthList.innerHTML = months.map((month) => {
      const maxBucket = Math.max(1, ...month.buckets.map((bucket) => bucket.count));
      const bucketRows = month.buckets.map((bucket) => `
        <div class="weight-row">
          <span class="weight-label">${escapeHtml(bucket.label)}</span>
          <div class="weight-track"><div class="weight-fill" style="width:${(bucket.count / maxBucket) * 100}%"></div></div>
          <span class="weight-count">${bucket.count} 件</span>
        </div>`).join("");
      return `<details class="month-item">
        <summary>
          <span class="month-name">${escapeHtml(month.label)}</span>
          <span class="month-metric"><span>总价格</span><strong>${money(month.totalCost)}</strong></span>
          <span class="month-metric"><span>总件数</span><strong>${month.count.toLocaleString("zh-CN")}</strong></span>
          <span class="month-metric"><span>均价</span><strong>${money(month.averageCost)}</strong></span>
          <span class="chevron" aria-hidden="true"></span>
        </summary>
        <div class="month-expanded">
          <div class="secondary-stats">
            <div class="secondary-stat"><span>首重件数</span><strong>${month.firstWeightCount.toLocaleString("zh-CN")}</strong></div>
            <div class="secondary-stat"><span>续重件数</span><strong>${month.additionalWeightCount.toLocaleString("zh-CN")}</strong></div>
            <div class="secondary-stat"><span>已购面单成本</span><strong>${money(month.labelTotal)}</strong></div>
            <div class="secondary-stat"><span>需补差价</span><strong>${money(month.topUpCost)}</strong></div>
            <div class="secondary-stat"><span>总计费重量</span><strong>${decimal(month.totalWeight)} kg</strong></div>
          </div>
          <div class="weight-bars">${bucketRows}</div>
        </div>
      </details>`;
    }).join("");

    const issueCount = state.issues.length;
    els.sourceSummary.textContent = `${state.files.length} 个文件 · ${issueCount ? `${issueCount} 项提示` : "数据检查通过"}`;
    els.sourceDetails.innerHTML = `
      <div class="file-list">${state.files.map((name) => `<span class="file-chip">${escapeHtml(name)}</span>`).join("")}</div>
      <div class="issue-box ${issueCount ? "" : "clean"}">${issueCount ? state.issues.slice(0, 12).map((issue) => escapeHtml(issue.message)).join("<br>") + (issueCount > 12 ? `<br>另有 ${issueCount - 12} 项已记录到导出文件` : "") : "未发现重复运单、无效重量或无法识别的月份。"}</div>`;
  }

  function exportWorkbook() {
    if (!state.result || !window.XLSX) return;
    const { overall, months, records, settings, formula } = state.result;
    const summaryRows = [
      ["月份", "总件数", "总价格（元）", "面单成本（元）", "需补差价（元）", "均价（元/件）", "补差均价（元/件）", "首重件数", "续重件数", "总计费重量（kg）", "平均重量（kg）"],
      ...months.map((month) => [month.label, month.count, round2(month.totalCost), round2(month.labelTotal), round2(month.topUpCost), round2(month.averageCost), round2(month.topUpAverage), month.firstWeightCount, month.additionalWeightCount, round2(month.totalWeight), round2(month.averageWeight)]),
      ["全部月份", overall.count, round2(overall.totalCost), round2(overall.labelTotal), round2(overall.topUpCost), round2(overall.averageCost), round2(overall.topUpAverage), overall.firstWeightCount, overall.additionalWeightCount, round2(overall.totalWeight), round2(overall.averageWeight)]
    ];
    const bucketRows = [["月份", ...overall.buckets.map((bucket) => bucket.label)]];
    months.forEach((month) => bucketRows.push([month.label, ...month.buckets.map((bucket) => bucket.count)]));
    bucketRows.push(["全部月份", ...overall.buckets.map((bucket) => bucket.count)]);
    const detailRows = [["月份", "运单编号", "内部计费重量（kg）", "计算费用（元）", "来源文件", "工作表", "来源行"]];
    records.forEach((row) => detailRows.push([E.monthLabel(row.month), row.waybill, row.weight, round2(row.cost), row.fileName, row.sheetName, row.sourceRow]));
    const settingRows = [
      ["项目", "当前设置"], ["首重重量（kg）", settings.firstWeight], ["首重价格（元）", settings.firstPrice],
      ["续重价格（元/kg）", settings.additionalPrice], ["续重步长（kg）", settings.step], ["续重方式", settings.rounding],
      ["面单单价（元/个）", settings.labelPrice],
      ["实际使用公式", formula], ["重量分界值（kg）", settings.bucketCuts.join(", ")]
    ];
    const issueRows = [["类型", "提示"], ...state.issues.map((issue) => [issue.type, issue.message])];
    const workbook = XLSX.utils.book_new();
    appendSheet(workbook, "月度汇总", summaryRows, [14, 12, 16, 16, 16, 16, 18, 12, 12, 18, 16]);
    appendSheet(workbook, "重量分段", bucketRows, [14, ...overall.buckets.map(() => 14)]);
    appendSheet(workbook, "计费明细", detailRows, [14, 20, 20, 16, 32, 16, 10]);
    appendSheet(workbook, "计费设置", settingRows, [22, 76]);
    appendSheet(workbook, "异常记录", issueRows.length > 1 ? issueRows : [["类型", "提示"], ["检查", "未发现异常"]], [16, 72]);
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `寄件成本统计_${stamp}.xlsx`);
    toast("统计表已导出");
  }

  function appendSheet(workbook, name, rows, widths) {
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = widths.map((wch) => ({ wch }));
    sheet["!autofilter"] = { ref: sheet["!ref"] };
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  }

  function round2(value) { return Math.round((Number(value) + Number.EPSILON) * 100) / 100; }

  function populateSettings(settings) {
    const form = els.settingsForm.elements;
    form.firstWeight.value = settings.firstWeight;
    form.firstPrice.value = settings.firstPrice;
    form.additionalPrice.value = settings.additionalPrice;
    form.labelPrice.value = settings.labelPrice;
    form.step.value = settings.step;
    form.rounding.value = settings.rounding;
    form.formula.value = settings.formula;
    form.bucketCuts.value = settings.bucketCuts.join(", ");
    form.autoExport.checked = settings.autoExport;
    updateFormulaUI();
  }

  function formSettings() {
    const form = els.settingsForm.elements;
    return E.normalizeSettings({
      firstWeight: form.firstWeight.value,
      firstPrice: form.firstPrice.value,
      additionalPrice: form.additionalPrice.value,
      labelPrice: form.labelPrice.value,
      step: form.step.value,
      rounding: form.rounding.value,
      formula: form.formula.value,
      bucketCuts: form.bucketCuts.value.split(/[,，]/).map((value) => Number(value.trim())),
      autoExport: form.autoExport.checked
    });
  }

  function updateFormulaUI() {
    const settings = formSettings();
    els.formulaSection.hidden = settings.rounding !== "custom";
    try {
      const sample = E.calculateCost(2.3, settings);
      els.formulaPreview.classList.remove("error");
      els.formulaPreview.textContent = `示例：2.3kg = ${money(sample)}`;
    } catch (error) {
      els.formulaPreview.classList.add("error");
      els.formulaPreview.textContent = error.message;
    }
  }

  els.dropZone.addEventListener("click", () => els.fileInput.click());
  els.dropZone.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); els.fileInput.click(); } });
  els.fileInput.addEventListener("change", () => handleFiles(els.fileInput.files));
  ["dragenter", "dragover"].forEach((name) => els.dropZone.addEventListener(name, (event) => { event.preventDefault(); els.dropZone.classList.add("dragging"); }));
  ["dragleave", "drop"].forEach((name) => els.dropZone.addEventListener(name, (event) => { event.preventDefault(); els.dropZone.classList.remove("dragging"); }));
  els.dropZone.addEventListener("drop", (event) => handleFiles(event.dataTransfer.files));
  els.exportButton.addEventListener("click", exportWorkbook);
  els.settingsButton.addEventListener("click", () => { populateSettings(state.settings); els.settingsDialog.showModal(); });
  els.closeSettings.addEventListener("click", () => els.settingsDialog.close());
  els.settingsDialog.addEventListener("click", (event) => { if (event.target === els.settingsDialog) els.settingsDialog.close(); });
  els.settingsForm.addEventListener("input", updateFormulaUI);
  els.resetFormula.addEventListener("click", () => { els.settingsForm.elements.formula.value = E.DEFAULT_FORMULA; updateFormulaUI(); });
  els.resetSettings.addEventListener("click", () => populateSettings(E.DEFAULT_SETTINGS));
  els.settingsForm.addEventListener("submit", (event) => {
    event.preventDefault();
    try {
      const next = formSettings();
      E.calculateCost(2.3, next);
      state.settings = next;
      localStorage.setItem("courier-cost-settings", JSON.stringify(next));
      if (state.records.length) recalculate();
      els.settingsDialog.close();
      toast("设置已保存，统计已重新计算");
    } catch (error) { toast(error.message); }
  });

  populateSettings(state.settings);
})();
