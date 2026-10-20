(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.CourierEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const DEFAULT_FORMULA = "firstPrice + ceil(max(0, weight - firstWeight) / step) * additionalPrice";
  const EXACT_FORMULA = "firstPrice + max(0, weight - firstWeight) * additionalPrice";
  const DEFAULT_SETTINGS = Object.freeze({
    firstWeight: 1,
    firstPrice: 2.5,
    additionalPrice: 1,
    labelPrice: 2,
    step: 1,
    rounding: "ceil",
    formula: DEFAULT_FORMULA,
    bucketCuts: [1, 2, 3, 5],
    autoExport: true
  });

  const HEADER_ALIASES = {
    waybill: ["运单编号", "运单号", "快递单号", "物流单号"],
    date: ["寄件时间", "发货时间", "揽件时间", "下单时间", "寄件日期", "日期"],
    weight: ["内部计费重量", "计费重量", "结算重量", "收费重量", "重量"]
  };

  function normalizeHeader(value) {
    return String(value == null ? "" : value).replace(/[\s\n\r（）()]/g, "").toLowerCase();
  }

  function findHeaderIndex(rows) {
    const aliases = Object.fromEntries(Object.entries(HEADER_ALIASES).map(([key, values]) => [key, values.map(normalizeHeader)]));
    for (let rowIndex = 0; rowIndex < Math.min(rows.length, 25); rowIndex += 1) {
      const normalized = (rows[rowIndex] || []).map(normalizeHeader);
      const columns = {};
      for (const [key, names] of Object.entries(aliases)) {
        columns[key] = normalized.findIndex((cell) => names.includes(cell));
      }
      if (columns.weight >= 0 && (columns.date >= 0 || columns.waybill >= 0)) return { rowIndex, columns };
    }
    return null;
  }

  function excelSerialToDate(serial) {
    if (!Number.isFinite(serial) || serial < 1) return null;
    return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") return excelSerialToDate(value);
    if (typeof value !== "string") return null;
    const text = value.trim();
    let match = text.match(/(20\d{2})\s*[年\-/\.]\s*(\d{1,2})/);
    if (match) return new Date(Number(match[1]), Number(match[2]) - 1, 1);
    const parsed = new Date(text.replace(/-/g, "/"));
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function inferMonth(sheetName, fileName) {
    const combined = `${fileName || ""} ${sheetName || ""}`;
    const yearMatch = combined.match(/(20\d{2})\s*年?/);
    const monthMatch = String(sheetName || "").match(/(?:^|\D)(1[0-2]|0?[1-9])\s*月?(?:\D|$)/);
    if (!yearMatch || !monthMatch) return null;
    return `${yearMatch[1]}-${String(Number(monthMatch[1])).padStart(2, "0")}`;
  }

  function monthKey(value, fallback) {
    const date = parseDate(value);
    if (!date) return fallback || null;
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  function monthLabel(key) {
    const [year, month] = String(key).split("-");
    return `${year}年${Number(month)}月`;
  }

  function numberValue(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    if (typeof value !== "string") return null;
    const cleaned = value.replace(/[,，\s]/g, "").replace(/kg$/i, "");
    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function normalizeSettings(input) {
    const base = { ...DEFAULT_SETTINGS, ...(input || {}) };
    base.firstWeight = Math.max(0.01, Number(base.firstWeight) || DEFAULT_SETTINGS.firstWeight);
    base.firstPrice = Math.max(0, Number(base.firstPrice) || 0);
    base.additionalPrice = Math.max(0, Number(base.additionalPrice) || 0);
    base.labelPrice = Math.max(0, Number(base.labelPrice) || 0);
    base.step = Math.max(0.01, Number(base.step) || 1);
    base.rounding = ["ceil", "exact", "custom"].includes(base.rounding) ? base.rounding : "ceil";
    base.formula = String(base.formula || DEFAULT_FORMULA).trim();
    base.bucketCuts = Array.from(new Set((Array.isArray(base.bucketCuts) ? base.bucketCuts : [1, 2, 3, 5])
      .map(Number).filter((value) => Number.isFinite(value) && value > 0))).sort((a, b) => a - b);
    if (!base.bucketCuts.length) base.bucketCuts = [1, 2, 3, 5];
    base.autoExport = Boolean(base.autoExport);
    return base;
  }

  function activeFormula(settings) {
    if (settings.rounding === "exact") return EXACT_FORMULA;
    if (settings.rounding === "ceil") return DEFAULT_FORMULA;
    return settings.formula;
  }

  function compileFormula(formula) {
    const expression = String(formula || "").trim();
    if (!expression) throw new Error("公式不能为空");
    if (!/^[0-9A-Za-z_+\-*/%().,\s]+$/.test(expression)) throw new Error("公式含有不支持的符号");
    const allowed = new Set(["weight", "firstWeight", "firstPrice", "additionalPrice", "step", "ceil", "floor", "round", "max", "min", "abs"]);
    const identifiers = expression.match(/[A-Za-z_][A-Za-z0-9_]*/g) || [];
    const invalid = identifiers.find((name) => !allowed.has(name));
    if (invalid) throw new Error(`不支持的变量或函数：${invalid}`);
    const mapped = expression.replace(/\b(ceil|floor|round|max|min|abs)\b/g, "Math.$1");
    const fn = Function("weight", "firstWeight", "firstPrice", "additionalPrice", "step", `"use strict"; return (${mapped});`);
    return function evaluate(weight, settings) {
      const result = fn(weight, settings.firstWeight, settings.firstPrice, settings.additionalPrice, settings.step);
      if (!Number.isFinite(result) || result < 0) throw new Error("公式结果必须是大于或等于 0 的数字");
      return result;
    };
  }

  function calculateCost(weight, inputSettings) {
    const settings = normalizeSettings(inputSettings);
    return compileFormula(activeFormula(settings))(Number(weight), settings);
  }

  function extractRowsFromSheet(rows, sheetName, fileName) {
    const header = findHeaderIndex(rows);
    if (!header) return { records: [], issues: [{ type: "sheet", message: `${sheetName}：未找到“内部计费重量”等必要列` }] };
    const { columns } = header;
    const fallbackMonth = inferMonth(sheetName, fileName);
    const records = [];
    const issues = [];
    for (let index = header.rowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      if (row.every((cell) => cell == null || String(cell).trim() === "")) continue;
      const weight = numberValue(row[columns.weight]);
      const month = columns.date >= 0 ? monthKey(row[columns.date], fallbackMonth) : fallbackMonth;
      const waybill = columns.waybill >= 0 ? String(row[columns.waybill] == null ? "" : row[columns.waybill]).trim() : "";
      if (!(weight > 0)) {
        issues.push({ type: "row", message: `${sheetName} 第${index + 1}行：计费重量无效` });
        continue;
      }
      if (!month) {
        issues.push({ type: "row", message: `${sheetName} 第${index + 1}行：无法识别寄件月份` });
        continue;
      }
      records.push({ fileName, sheetName, sourceRow: index + 1, waybill, month, weight });
    }
    return { records, issues };
  }

  function deduplicate(records) {
    const seen = new Set();
    const unique = [];
    const duplicates = [];
    for (const record of records) {
      const key = record.waybill ? `waybill:${record.waybill}` : `${record.fileName}|${record.sheetName}|${record.sourceRow}`;
      if (seen.has(key)) duplicates.push(record);
      else { seen.add(key); unique.push(record); }
    }
    return { records: unique, duplicates };
  }

  function bucketDefinitions(cuts) {
    return cuts.map((cut, index) => ({
      min: index === 0 ? null : cuts[index - 1],
      max: cut,
      label: index === 0 ? `≤${cut}kg` : `${cuts[index - 1]}–${cut}kg`
    })).concat([{ min: cuts[cuts.length - 1], max: null, label: `>${cuts[cuts.length - 1]}kg` }]);
  }

  function summarize(inputRecords, inputSettings) {
    const settings = normalizeSettings(inputSettings);
    const evaluator = compileFormula(activeFormula(settings));
    const records = inputRecords.map((record) => ({ ...record, cost: evaluator(record.weight, settings) }));
    const grouped = new Map();
    for (const record of records) {
      if (!grouped.has(record.month)) grouped.set(record.month, []);
      grouped.get(record.month).push(record);
    }
    const bucketDefs = bucketDefinitions(settings.bucketCuts);
    const summarizeGroup = (month, group) => {
      const totalCost = group.reduce((sum, row) => sum + row.cost, 0);
      const totalWeight = group.reduce((sum, row) => sum + row.weight, 0);
      const labelTotal = group.length * settings.labelPrice;
      const buckets = bucketDefs.map((bucket) => ({ ...bucket, count: group.filter((row) => (bucket.min == null || row.weight > bucket.min) && (bucket.max == null || row.weight <= bucket.max)).length }));
      return {
        month,
        label: month === "TOTAL" ? "全部月份" : monthLabel(month),
        count: group.length,
        totalCost,
        averageCost: group.length ? totalCost / group.length : 0,
        labelTotal,
        topUpCost: totalCost - labelTotal,
        topUpAverage: group.length ? (totalCost - labelTotal) / group.length : 0,
        totalWeight,
        averageWeight: group.length ? totalWeight / group.length : 0,
        firstWeightCount: group.filter((row) => row.weight <= settings.firstWeight).length,
        additionalWeightCount: group.filter((row) => row.weight > settings.firstWeight).length,
        buckets
      };
    };
    const months = [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([month, group]) => summarizeGroup(month, group));
    return { settings, records, months, overall: summarizeGroup("TOTAL", records), formula: activeFormula(settings) };
  }

  return {
    DEFAULT_FORMULA,
    EXACT_FORMULA,
    DEFAULT_SETTINGS,
    normalizeSettings,
    activeFormula,
    compileFormula,
    calculateCost,
    extractRowsFromSheet,
    deduplicate,
    summarize,
    monthLabel,
    bucketDefinitions
  };
});
