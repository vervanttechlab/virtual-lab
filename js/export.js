// Export a student's answers for one activity as an Excel file (.xlsx) to turn in on Google Classroom.
// Sheet 1 "Answers": every question, the student's answer, the auto-check result and points.
// Then the student's workbook: formulas (with Excel-compatible names), values, number formats, the yellow/green/header
// cell colours, highlight rules, ToolPak output sheets and charts (as pictures).
import { drawChartOn } from "./grid.js";

const WB = window.WBCore;
const GREEN = "FF107C41", HEAD = "FF1F4E79", YELLOW = "FFFFF2CC", GIVEN = "FFE2EFDA", LINE = "FFD4D4D4";

let loading = null;
function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (!loading) {
    loading = new Promise((ok, bad) => {
      const s = document.createElement("script");
      s.src = "vendor/exceljs.min.js";
      s.onload = () => ok(window.ExcelJS);
      s.onerror = () => { loading = null; bad(new Error("Could not load the Excel exporter. Check your connection and try again.")); };
      document.head.append(s);
    });
  }
  return loading;
}

// Newer Excel functions must be stored with the _xlfn. prefix inside .xlsx files, or Excel shows #NAME?.
const XLFN = ["MODE.SNGL", "MODE.MULT", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "QUARTILE.INC", "QUARTILE.EXC", "PERCENTILE.INC",
  "PERCENTILE.EXC", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "T.TEST", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV",
  "T.INV.2T", "FORECAST.LINEAR", "XLOOKUP", "IFS", "MAXIFS", "MINIFS", "CONCAT", "TEXTJOIN", "COVARIANCE.S", "COVARIANCE.P",
  "CONFIDENCE.NORM", "CHISQ.TEST", "F.TEST", "Z.TEST", "SKEW.P", "RANK.EQ", "RANK.AVG"];
const XLFN_RE = new RegExp(`(?<![\\w.])(${XLFN.map((f) => f.replace(/\./g, "\\.")).sort((a, b) => b.length - a.length).join("|")})\\s*\\(`, "gi");
export function toExcelFormula(f) {
  return String(f).replace(/^=/, "").replace(XLFN_RE, (m, name) => `_xlfn.${name.toUpperCase()}(`);
}

const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const thin = { style: "thin", color: { argb: LINE } };
const box = { top: thin, left: thin, bottom: thin, right: thin };

function cellResult(v) {
  if (v && typeof v === "object") return { error: v.value || "#VALUE!" };
  return v ?? "";
}

/** What the student answered for a part, as text. */
function answerText(p, ans, wb, sub) {
  const a = ans[p.id];
  switch (p.type) {
    case "choice": { const o = (p.options || []).find((x) => x[0] === a); return o ? o[1] : (a ?? ""); }
    case "multi": return (a || []).map((v) => ((p.options || []).find((x) => x[0] === v) || [v, v])[1]).join("; ");
    case "cells": return `${p.cells.length} cells (see the workbook sheets)`;
    case "textcells": return p.cells.map((ref) => `${ref}: ${sub.cells[ref]?.v ?? ""}`).join("\n");
    case "spec": return "";
    case "sql": return a && typeof a === "object" ? String(a.sql || "") : "";
    default: return a == null ? "" : String(a);
  }
}

function resultText(p, r) {
  if (!r) return "";
  if (r.manual) return "Graded by the trainer";
  if (p.type === "cells") {
    const d = r.detail || {}, ok = Object.values(d).filter((x) => x === 1).length;
    return `${ok} of ${p.cells.length} correct`;
  }
  if (p.type === "spec") return r.frac === 1 ? "Done" : "Not found";
  return r.frac === 1 ? "Correct" : "Not correct";
}

async function chartPng(wb, spec) {
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-10000px;top:0;width:720px;height:400px;background:#fff";
  const cv = document.createElement("canvas");
  cv.width = 720; cv.height = 400;
  cv.style.cssText = "width:720px;height:400px";
  host.append(cv);
  document.body.append(host);
  try {
    const ch = drawChartOn(wb, cv, spec, { print: true });
    // white background behind the transparent canvas
    const out = document.createElement("canvas");
    out.width = cv.width; out.height = cv.height;
    const g = out.getContext("2d");
    g.fillStyle = "#ffffff"; g.fillRect(0, 0, out.width, out.height);
    g.drawImage(cv, 0, 0);
    const url = out.toDataURL("image/png");
    ch.destroy();
    return url;
  } finally {
    host.remove();
  }
}

/**
 * @returns {Promise<{blob: Blob, filename: string}>}
 */
export async function buildSubmissionFile({ day, act, wb, ans, sub, score, student }) {
  const ExcelJS = await loadExcelJS();
  const book = new ExcelJS.Workbook();
  book.creator = "DAL3 Virtual Lab";
  book.created = new Date();

  // ---------------------------------------------------------- Answers sheet
  const ws = book.addWorksheet("Answers", { views: [{ state: "frozen", ySplit: 7 }] });
  ws.columns = [{ width: 6 }, { width: 58 }, { width: 46 }, { width: 22 }, { width: 10 }];
  const title = ws.getCell("A1");
  title.value = `Day ${day.day} · Activity ${act.id}: ${act.title}`;
  title.font = { bold: true, size: 14, color: { argb: GREEN } };
  const info = [["Student", `${student.name || ""} (${student.id || ""})`], ["Course", "TESDA Data Analytics Level III · DAL3 Virtual Lab"],
    ["Exported", new Date().toLocaleString()],
    ["Auto-checked score", `${score.auto} of ${score.autoMax} pts` + (score.manualMax ? `  (+ ${score.manualMax} pts graded by the trainer)` : "")]];
  info.forEach(([k, v], i) => {
    ws.getCell(`A${i + 2}`).value = k;
    ws.getCell(`A${i + 2}`).font = { color: { argb: "FF605E5C" } };
    ws.mergeCells(`A${i + 2}:B${i + 2}`);
    ws.getCell(`C${i + 2}`).value = v;
    ws.getCell(`C${i + 2}`).font = { bold: i === 0 || i === 3 };
  });
  const head = ws.getRow(7);
  head.values = ["No.", "Question", "Your answer", "Result", "Points"];
  head.eachCell((c) => { c.fill = solid(GREEN); c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.border = box; c.alignment = { vertical: "middle" }; });
  head.height = 20;
  act.parts.forEach((p, i) => {
    const r = score.parts[p.id];
    const row = ws.getRow(8 + i);
    row.values = [i + 1, p.prompt || p.label || "", answerText(p, ans, wb, sub), resultText(p, r),
      r && !r.manual ? `${r.pts} / ${p.points}` : `— / ${p.points}`];
    row.eachCell({ includeEmpty: true }, (c, col) => {
      if (col > 5) return;
      c.border = box;
      c.alignment = { wrapText: true, vertical: "top" };
    });
    const res = row.getCell(4);
    if (r && !r.manual) res.font = { bold: true, color: { argb: r.frac === 1 ? GREEN : r.frac > 0 ? "FF7A5900" : "FFA4262C" } };
  });
  const endRow = 8 + act.parts.length + 1;
  ws.getCell(`A${endRow}`).value = act.rubric ? `Rubric: ${act.rubric}` : "Written answers are graded by the trainer.";
  ws.getCell(`A${endRow}`).font = { italic: true, color: { argb: "FF605E5C" } };

  // ---------------------------------------------------------- the student's workbook
  if (wb) {
    for (const name of wb.sheetNames) {
      const meta = wb.meta[name] || {};
      const sh = book.addWorksheet(name.slice(0, 31));
      const { rows, cols } = wb.dims(name);
      const widths = [];
      for (let c = 0; c < cols; c++) widths.push({ width: Math.max(6, Math.min(60, (meta.colW && meta.colW[WB.numToCol(c + 1)]) || 10)) });
      sh.columns = widths;
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const raw = wb.raw(name, r, c), style = wb.style(name, r, c);
          if ((raw === "" || raw == null) && !style) continue;
          const cell = sh.getCell(r + 1, c + 1);
          if (typeof raw === "string" && raw.startsWith("=")) cell.value = { formula: toExcelFormula(raw), result: cellResult(wb.value(name, r, c)) };
          else if (raw !== "" && raw != null) cell.value = raw;
          const nf = wb.numFmt(name, r, c);
          if (nf) cell.numFmt = nf;
          if (style === "i") cell.fill = solid(YELLOW);
          else if (style === "g") { cell.fill = solid(GIVEN); cell.font = { bold: true }; }
          else if (style === "h") { cell.fill = solid(HEAD); cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; }
        }
      }
      for (const rule of wb.specs.rules.filter((x) => x.sheet === name)) {
        const fill = { red: ["FFFFC7CE", "FF9C0006"], yellow: ["FFFFEB9C", "FF9C5700"], green: ["FFC6EFCE", "FF006100"] }[rule.color || "red"];
        sh.addConditionalFormatting({ ref: rule.range, rules: [{ type: "cellIs", operator: "equal", formulae: [`"${String(rule.value).replace(/"/g, '""')}"`],
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: fill[0] } }, font: { color: { argb: fill[1] } } } }] });
      }
      let top = 1;
      for (const spec of wb.specs.charts.filter((x) => x.sheet === name)) {
        try {
          const id = book.addImage({ base64: await chartPng(wb, spec), extension: "png" });
          sh.addImage(id, { tl: { col: cols + 1, row: top }, ext: { width: 720, height: 400 } });
          top += 22;
        } catch (e) { console.warn("chart export failed", e); }
      }
    }
  }

  const buf = await book.xlsx.writeBuffer();
  const safe = (s) => String(s || "").normalize("NFKD").replace(/[^\w-]+/g, "").slice(0, 24);
  const last = String(student.name || "").trim().split(/\s+/).pop();
  const filename = `Day${String(day.day).padStart(2, "0")}_Activity${act.id}_${safe(student.id)}${last ? "_" + safe(last) : ""}.xlsx`;
  return { blob: new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename };
}

export function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
