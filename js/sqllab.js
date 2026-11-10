// SQL lab: real SQLite in the browser (sql.js, WebAssembly) on aling_nena.db, the same database the PC students open in
// VS Code + SQLTools. Every query runs on a FRESH copy of the database, so DELETE/UPDATE/DROP can never break anything.
import { el, icon, toast } from "./ui.js";

const C = window.DAL3Check;
let enginePromise = null;

function loadScriptOnce(src) {
  return new Promise((ok, bad) => {
    if (document.querySelector(`script[data-src="${src}"]`)) return ok();
    const s = document.createElement("script");
    s.src = src;
    s.dataset.src = src;
    s.onload = ok;
    s.onerror = () => bad(new Error("Could not load the SQL engine. Check your connection and try again."));
    document.head.append(s);
  });
}

/** {SQL, bytes} once per page load (about 0.7 MB, cached by the service worker for offline use). */
export function loadSql() {
  if (!enginePromise) {
    enginePromise = (async () => {
      await loadScriptOnce("vendor/sql-wasm.js");
      const [SQL, buf] = await Promise.all([
        window.initSqlJs({ locateFile: (f) => `vendor/${f}` }),
        fetch("data/aling_nena.db").then((r) => { if (!r.ok) throw new Error("Could not load the database file."); return r.arrayBuffer(); }),
      ]);
      return { SQL, bytes: new Uint8Array(buf) };
    })().catch((e) => { enginePromise = null; throw e; });
  }
  return enginePromise;
}

const TABLES = {
  branches: ["branch_id", "branch_name", "city", "manager", "opened_on", "floor_area_sqm", "status"],
  products: ["product_id", "product_name", "category", "unit_price", "supplier"],
  customers: ["customer_id", "first_name", "last_name", "age_group", "city", "mobile", "email", "joined_on"],
  sales: ["sale_id", "sale_date", "sale_time", "branch_id", "product_id", "customer_id", "quantity", "unit_price", "total_amount", "payment_method"],
};
const PK = { branches: "branch_id", products: "product_id", customers: "customer_id", sales: "sale_id" };
const FK = { "sales.branch_id": "branches", "sales.product_id": "products", "sales.customer_id": "customers" };

/** Plain-English version of SQLite's error messages. */
export function friendlyError(msg) {
  const m = String(msg || "").replace(/^Error:\s*/, "");
  let x;
  if ((x = m.match(/no such table: (\S+)/i))) return `There is no table called "${x[1]}". The tables are: branches, products, customers, sales.`;
  if ((x = m.match(/no such column: (\S+)/i))) return `There is no column called "${x[1]}". Check the spelling in the Tables list (tap "Tables").`;
  if ((x = m.match(/ambiguous column name: (\S+)/i))) return `"${x[1]}" is in two of your tables. Put the table alias in front, e.g. s.${x[1].split(".").pop()}.`;
  if (/incomplete input/i.test(m)) return "The query is not finished. Check for a missing closing quote ' or bracket ).";
  if (/unrecognized token/i.test(m)) return "A quote was opened but not closed. Text needs a quote on both sides: 'GCash'.";
  if ((x = m.match(/near "([^"]*)": syntax error/i))) return `Syntax error near "${x[1]}". Common causes: a comma before FROM, a missing comma between columns, or a missing keyword.`;
  return m;
}

/** Run SQL on a fresh database. Returns the LAST result set: {columns, rows}. Throws Error(friendly message). */
export async function runSql(text) {
  const { SQL, bytes } = await loadSql();
  const db = new SQL.Database(bytes);
  try {
    const res = db.exec(String(text || ""));
    const last = res[res.length - 1];
    return last ? { columns: last.columns, rows: last.values } : { columns: [], rows: [] };
  } catch (e) {
    throw new Error(friendlyError(e.message));
  } finally {
    db.close();
  }
}

/** Fingerprint of a student's query result (null when there is no query or it fails). */
export async function fingerprintOf(text, ordered) {
  if (!String(text || "").trim()) return null;
  try {
    const r = await runSql(text);
    return { fp: C.sqlFingerprint(r.columns.length, r.rows, !!ordered), rows: r.rows.length, cols: r.columns.length };
  } catch (e) {
    return { fp: "", error: e.message };
  }
}

export function toCsv(res) {
  const q = (v) => {
    if (v == null) return "";
    const s = String(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return [res.columns.map(q).join(","), ...res.rows.map((r) => r.map(q).join(","))].join("\r\n") + "\r\n";
}

export function resultTable(res, max = 50) {
  if (!res.columns.length) return el("p", { class: "small muted" }, "The query ran, but it returns no table (for example a comment only).");
  const shown = res.rows.slice(0, max);
  const t = el("table", { class: "sqlres" },
    el("thead", {}, el("tr", {}, ...res.columns.map((c) => el("th", {}, c)))),
    el("tbody", {}, ...shown.map((r) => el("tr", {}, ...r.map((v) => el("td", { class: v == null ? "null" : typeof v === "number" ? "num" : "" }, v == null ? "NULL" : String(v)))))));
  return el("div", { class: "sqlwrap" }, el("div", { class: "sqlscroll" }, t),
    el("div", { class: "sqlcount" }, res.rows.length > max ? `Showing ${max} of ${res.rows.length} rows` : `${res.rows.length} row${res.rows.length === 1 ? "" : "s"}`));
}

/** Collapsible list of tables and columns (like the SQLTools tree). Tapping a name inserts it into the last editor. */
export function schemaPanel(insert) {
  const d = el("details", { class: "schema" }, el("summary", {}, icon("db"), el("span", {}, "Tables"), el("span", { class: "small muted" }, "aling_nena.db · tap a name to insert it")));
  for (const [t, cols] of Object.entries(TABLES)) {
    const tb = el("button", { type: "button", class: "sname t" }, t);
    tb.addEventListener("click", () => insert(t));
    const cs = cols.map((c) => {
      const b = el("button", { type: "button", class: "sname" }, c, PK[t] === c ? el("span", { class: "kb" }, "PK") : FK[`${t}.${c}`] ? el("span", { class: "kb fk" }, `FK → ${FK[`${t}.${c}`]}`) : "");
      b.addEventListener("click", () => insert(c));
      return b;
    });
    d.append(el("div", { class: "stable" }, tb, el("div", { class: "scols" }, ...cs)));
  }
  return d;
}

const TOKENS = ["SELECT ", "* ", "FROM ", "WHERE ", "AND ", "OR ", "= ", "<> ", "> ", "< ", ">= ", "'", ", ", "; ", "( ", ") ", "% ", "LIKE ", "IN ",
  "BETWEEN ", "IS NULL ", "DISTINCT ", "ORDER BY ", "DESC ", "LIMIT ", "COUNT(*) ", "SUM(", "AVG(", "ROUND(", "MIN(", "MAX(", "AS ",
  "INNER JOIN ", "LEFT JOIN ", "ON "];

let lastEditor = null;
export function insertAtCursor(ta, text) {
  ta = ta || lastEditor;
  if (!ta) return;
  const s = ta.selectionStart ?? ta.value.length, e = ta.selectionEnd ?? ta.value.length;
  let t = text.replace(/ $/, "");
  if (/[\w)*']$/.test(ta.value.slice(0, s)) && /^[A-Za-z*(]/.test(t)) t = " " + t;   // keep words apart
  if (text.endsWith(" ")) t += " ";
  ta.value = ta.value.slice(0, s) + t + ta.value.slice(e);
  ta.selectionStart = ta.selectionEnd = s + t.length;
  ta.dispatchEvent(new Event("input", { bubbles: true }));
  ta.focus();
}

/**
 * A SQL editor with a phone keypad, Run button and result area.
 *   value, onInput(text), onRun(result|null, error|null)
 *   extra: buttons appended after Run (e.g. Download CSV)
 */
export function sqlEditor({ value = "", onInput, onRun, placeholder = "Type your SQL query here, e.g. SELECT * FROM branches;", rows = 4, extra = [] }) {
  const ta = el("textarea", { class: "sqlbox mono", rows, placeholder, spellcheck: false, autocapitalize: "off", autocomplete: "off", autocorrect: "off" });
  ta.value = value;
  ta.addEventListener("focus", () => { lastEditor = ta; });
  ta.addEventListener("input", () => onInput && onInput(ta.value));
  const keys = el("div", { class: "sqlkeys", role: "toolbar", "aria-label": "SQL words" }, ...TOKENS.map((t) => {
    const b = el("button", { type: "button" }, t.trim() || t);
    b.addEventListener("pointerdown", (e) => e.preventDefault());                 // keep the cursor in the editor
    b.addEventListener("click", () => insertAtCursor(ta, t));
    return b;
  }));
  const out = el("div", { class: "sqlout" });
  const run = el("button", { class: "btn small primary", type: "button" }, icon("play"), "Run");
  let last = null;
  async function doRun() {
    if (!ta.value.trim()) { toast("Type a query first."); return; }
    run.disabled = true;
    out.replaceChildren(el("p", { class: "small muted" }, "Running…"));
    try {
      last = await runSql(ta.value);
      out.replaceChildren(resultTable(last));
      onRun && onRun(last, null);
    } catch (e) {
      last = null;
      out.replaceChildren(el("p", { class: "sqlerr" }, e.message));
      onRun && onRun(null, e.message);
    } finally {
      run.disabled = false;
    }
  }
  run.addEventListener("click", doRun);
  ta.addEventListener("keydown", (e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); doRun(); } });
  const box = el("div", { class: "sqled" }, ta, keys, el("div", { class: "btnrow" }, run, ...extra, el("span", { class: "small muted hide-sm" }, "Ctrl + Enter also runs")), out);
  box.textarea = ta;
  box.lastResult = () => last;
  box.run = doRun;
  return box;
}

export function downloadCsv(res, filename) {
  const blob = new Blob(["﻿" + toCsv(res)], { type: "text/csv;charset=utf-8" });
  const a = el("a", { href: URL.createObjectURL(blob), download: filename });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
