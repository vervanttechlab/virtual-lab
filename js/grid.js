// Spreadsheet UI: grid, formula bar, symbol keypad, toolbar, sheet tabs, charts, Analysis ToolPak, Highlight rules.
// Works with touch (phones) and with mouse + keyboard (desktop shortcuts as in Excel).
import { formatValue } from "./workbook.js";
import * as ST from "./stats.js";
import { el, modal, toast, esc, icon } from "./ui.js";

const WB = window.WBCore;
const FUNCS = ["SUM", "COUNT", "COUNTA", "AVERAGE", "MEDIAN", "MODE.SNGL", "MIN", "MAX", "VAR.S", "STDEV.S", "STDEV.P", "QUARTILE.INC",
  "SKEW", "COUNTIF", "COUNTIFS", "SUMIFS", "AVERAGEIFS", "IF", "OR", "AND", "ABS", "SQRT", "ROUND", "CORREL", "RSQ", "SLOPE", "INTERCEPT",
  "FORECAST.LINEAR", "NORM.DIST", "STANDARDIZE", "T.TEST", "XLOOKUP", "TRIM", "PROPER", "TEXT"];
const REF_RE = /((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/g;

/** Cycle the reference at the caret: A1 -> $A$1 -> A$1 -> $A1 -> A1 (Excel F4). */
export function toggleRefAt(text, caret) {
  const one = /\$?[A-Za-z]{1,3}\$?\d+/g;
  let m, hit = null;
  while ((m = one.exec(text))) {
    if (caret >= m.index && caret <= m.index + m[0].length) hit = m;
  }
  if (!hit) return null;
  const t = hit[0], p = t.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  const state = (p[1] ? 2 : 0) + (p[3] ? 1 : 0);          // 0 A1, 3 $A$1, 1 A$1, 2 $A1
  const next = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
  const nt = (next & 2 ? "$" : "") + p[2] + (next & 1 ? "$" : "") + p[4];
  return { text: text.slice(0, hit.index) + nt + text.slice(hit.index + t.length), caret: hit.index + nt.length };
}

export class SheetUI {
  /**
   * @param {HTMLElement} root
   * @param {import('./workbook.js').Workbook} wb
   * @param {{sheet?:string, readOnly?:boolean, onChange?:Function, marks?:Object}} opts
   */
  constructor(root, wb, opts = {}) {
    this.root = root;
    this.wb = wb;
    this.opts = opts;
    this.sheet = opts.sheet && wb.sid(opts.sheet) !== undefined ? opts.sheet : wb.sheetNames[0];
    this.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.anchor = { r: 0, c: 0 };
    this.selectMode = false;
    this.editing = false;
    this.marks = opts.marks || {};                    // {"Sheet!A1": "ok"|"bad"}
    this.charts = new Map();
    this.build();
    this.unsub = wb.onChange((info) => this.onWbChange(info));
    this.renderSheet();
    SheetUI.live.add(this);
  }
  destroy() { SheetUI.live.delete(this); this.unsub && this.unsub(); this.charts.forEach((c) => c.destroy()); this.root.innerHTML = ""; }
  /** Before an app update reloads the page: keep a cell the student is still typing in (as if they pressed Enter). */
  static commitAll() { for (const ui of SheetUI.live) { try { if (ui.editing) ui.commit(); } catch (e) { console.warn(e); } } }
  static live = new Set();

  // ------------------------------------------------------------------ layout
  build() {
    const ro = this.opts.readOnly;
    this.root.innerHTML = "";
    this.root.classList.add("sheetui");
    this.nameBox = el("span", { class: "namebox", title: "Name Box: the selected cell or range" }, "A1");
    this.fx = el("input", { class: "fx", type: "text", autocomplete: "off", autocapitalize: "off", spellcheck: "false",
      placeholder: ro ? "" : "Select a cell, then type a value or =formula", readOnly: !!ro, "aria-label": "Formula bar" });
    const ok = el("button", { class: "fbtn ok", type: "button", title: "Enter", "aria-label": "Enter" }, icon("check"));
    const cancel = el("button", { class: "fbtn no", type: "button", title: "Cancel (Esc)", "aria-label": "Cancel" }, icon("close"));
    ok.addEventListener("pointerdown", (e) => e.preventDefault());
    cancel.addEventListener("pointerdown", (e) => e.preventDefault());
    ok.addEventListener("click", () => this.commit());
    cancel.addEventListener("click", () => this.cancelEdit());
    const fbar = el("div", { class: "fbar" }, this.nameBox, ro ? "" : cancel, ro ? "" : ok, el("span", { class: "fxsep" }),
      el("span", { class: "fxlabel" }, "fx"), this.fx);

    const keys = ["=", "(", ")", ":", ",", "$ F4", '"', "&", "<", ">", "+", "-", "*", "/", "^", "!"];
    this.keypad = el("div", { class: "keypad", role: "toolbar", "aria-label": "Formula symbols" },
      ...keys.map((k) => {
        const b = el("button", { class: "key", type: "button" }, k);
        b.addEventListener("pointerdown", (e) => e.preventDefault());       // keep focus in the formula bar
        b.addEventListener("click", () => this.keyPress(k));
        return b;
      }));

    // Ribbon (Excel-style tabs + command groups)
    const rb = (ic, label, title, fn) => {
      const b = el("button", { class: "rbtn", type: "button", title }, icon(ic), el("span", {}, label));
      b.addEventListener("pointerdown", (e) => { if (this.editing) e.preventDefault(); });
      b.addEventListener("click", fn);
      return b;
    };
    const grp = (...b) => el("div", { class: "rgroup" }, ...b);
    this.selBtn = rb("select", "Select range", "Select a range: select the first cell, turn this on, then select the last cell (or Shift + click)",
      () => this.toggleSelectMode());
    const copy = () => { this.wb.copy(this.sheet, this.sel); toast("Copied. Select where to paste, then Paste."); };
    const RIBBON = {
      Home: [
        grp(rb("undo", "Undo", "Undo (Ctrl + Z)", () => { if (!this.wb.undo()) toast("Nothing to undo."); })),
        grp(this.selBtn),
        grp(rb("copy", "Copy", "Copy (Ctrl + C)", copy), rb("paste", "Paste", "Paste (Ctrl + V)", () => this.doPaste()),
          rb("clear", "Clear", "Clear the selected cells (Delete)", () => this.wb.clear(this.sheet, this.sel))),
        grp(rb("filldown", "Fill Down", "Copy the first row of the selection down (Ctrl + D)", () => this.doFill("down")),
          rb("fillright", "Fill Right", "Copy the first column of the selection right (Ctrl + R)", () => this.doFill("right")),
          rb("fillend", "Fill to End", "Fill down to the last data row (double-click the fill handle)", () => this.doFill("end"))),
        grp(rb("format", "Number Format", "Comma style, percent, decimals", () => this.formatDialog()),
          rb("highlight", "Highlight Cells", "Conditional Formatting › Highlight Cells Rules › Equal To", () => this.highlightDialog())),
      ],
      Formulas: [
        grp(rb("sum", "AutoSum", "AutoSum (Alt + =)", () => this.autoSum()), rb("fx", "Insert Function", "Insert a function", () => this.funcPicker())),
        grp(rb("lock", "Lock Reference ($)", "Add $ signs to the reference at the cursor (F4)", () => this.keyPress("$ F4"))),
      ],
      Data: [
        grp(rb("analysis", "Data Analysis", "Analysis ToolPak: Descriptive Statistics, t-Test, Regression", () => this.analysisDialog())),
      ],
      Insert: [
        grp(rb("chart", "Chart", "Insert a chart from the selection", () => this.chartDialog())),
      ],
    };
    this.ribbon = el("div", { class: "ribbon", role: "toolbar", "aria-label": "Commands" });
    this.rtabs = el("div", { class: "rtabs", role: "tablist" });
    const showRibbon = (name) => {
      this.rtabs.querySelectorAll(".rtab").forEach((t) => t.classList.toggle("on", t.dataset.t === name));
      this.ribbon.replaceChildren(...RIBBON[name]);
    };
    for (const name of Object.keys(RIBBON)) {
      const t = el("button", { class: "rtab", type: "button", role: "tab", "data-t": name }, name);
      t.addEventListener("pointerdown", (e) => { if (this.editing) e.preventDefault(); });
      t.addEventListener("click", () => showRibbon(name));
      this.rtabs.append(t);
    }
    showRibbon("Home");

    this.gridWrap = el("div", { class: "gridwrap", tabindex: "0", "aria-label": "Spreadsheet grid" });
    this.tabs = el("div", { class: "tabs", role: "tablist" });
    this.selStats = el("div", { class: "selstats", "aria-live": "polite" });
    this.chartArea = el("div", { class: "chartarea" });
    this.root.append(ro ? "" : this.rtabs, ro ? "" : this.ribbon, fbar, ro ? "" : this.keypad, this.gridWrap,
      el("div", { class: "sheetfoot" }, this.tabs, this.selStats), this.chartArea);

    // formula bar events
    this.fx.addEventListener("focus", () => { if (!ro) this.beginEdit(false); });
    this.fx.addEventListener("input", () => { this.lastPick = null; if (this.editing) this.previewCell(); });
    this.fx.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); this.commit(); this.move(1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); this.commit(); this.move(0, 1); }
      else if (e.key === "Escape") { e.preventDefault(); this.cancelEdit(); }
      else if (e.key === "F4") { e.preventDefault(); this.keyPress("$ F4"); }
    });
    // grid events
    this.gridWrap.addEventListener("pointerdown", (e) => this.onCellDown(e));
    this.gridWrap.addEventListener("dblclick", (e) => { const td = e.target.closest("td[data-r]"); if (td && !ro) { this.fx.focus(); } });
    this.gridWrap.addEventListener("keydown", (e) => this.onGridKey(e));
  }

  // ------------------------------------------------------------------ rendering
  colPx(sheet, c) {
    const m = this.wb.meta[sheet];
    const w = m && m.colW && m.colW[WB.numToCol(c + 1)];
    const px = Math.round(Math.max(5, Math.min(w || 10, 46)) * 7.2 + 10);
    return c === 0 ? Math.min(px, Math.max(90, Math.round(window.innerWidth * 0.34))) : px;   // column A is frozen: keep it narrow on phones
  }
  renderSheet() {
    const sheet = this.sheet, { rows, cols } = this.wb.dims(sheet);
    const R = Math.min(rows + 3, 1200), Cn = Math.min(cols + 1, 40);
    this.view = { R, C: Cn };
    const colgroup = ["<colgroup><col style='width:44px'>"];
    let tableW = 44;
    for (let c = 0; c < Cn; c++) { const w = this.colPx(sheet, c); tableW += w; colgroup.push(`<col style="width:${w}px">`); }
    colgroup.push("</colgroup>");
    const head = ["<thead><tr><th class='corner'></th>"];
    for (let c = 0; c < Cn; c++) head.push(`<th data-col="${c}">${WB.numToCol(c + 1)}</th>`);
    head.push("</tr></thead>");
    const body = ["<tbody>"];
    for (let r = 0; r < R; r++) {
      body.push(`<tr><th data-row="${r}">${r + 1}</th>`);
      for (let c = 0; c < Cn; c++) body.push(this.cellHTML(sheet, r, c));
      body.push("</tr>");
    }
    body.push("</tbody>");
    this.gridWrap.innerHTML = `<table class="grid" style="width:${tableW}px">${colgroup.join("")}${head.join("")}${body.join("")}</table>`;
    this.table = this.gridWrap.firstChild;
    // keep scrolled-to cells clear of the frozen header row and frozen column A
    this.gridWrap.style.scrollPaddingLeft = `${44 + this.colPx(sheet, 0)}px`;
    this.gridWrap.style.scrollPaddingTop = "30px";
    this.renderTabs();
    this.renderCharts();
    this.paintSelection();
  }
  cellHTML(sheet, r, c) {
    const v = this.wb.value(sheet, r, c);
    const st = this.wb.style(sheet, r, c);
    const cls = [];
    if (st) cls.push("s-" + st);
    if (typeof v === "number") cls.push("num");
    if (v && typeof v === "object") cls.push("err");
    const mk = this.marks[`${sheet}!${WB.addr(r, c)}`];
    if (mk) cls.push("mk-" + mk);
    const hl = this.wb.highlightOf(sheet, r, c);
    if (hl) cls.push("hl-" + hl);
    const txt = formatValue(v, this.wb.numFmt(sheet, r, c));
    return `<td data-r="${r}" data-c="${c}"${cls.length ? ` class="${cls.join(" ")}"` : ""}>${esc(txt)}</td>`;
  }
  td(r, c) { return this.table && this.table.rows[r + 1] && this.table.rows[r + 1].cells[c + 1]; }
  refreshCell(r, c) {
    const td = this.td(r, c);
    if (!td) return;
    const tmp = document.createElement("tbody");
    tmp.innerHTML = `<tr>${this.cellHTML(this.sheet, r, c)}</tr>`;
    const n = tmp.firstChild.firstChild;
    td.className = n.className;
    td.textContent = n.textContent;
    if (this.inSel(r, c)) td.classList.add("sel");
    if (r === this.anchor.r && c === this.anchor.c) td.classList.add("anchor");
  }
  refreshAll() {
    if (!this.table) return;
    const { rows, cols } = this.wb.dims(this.sheet);
    if (rows + 3 > this.view.R || cols + 1 > this.view.C) return this.renderSheet();
    for (let r = 0; r < this.view.R; r++) for (let c = 0; c < this.view.C; c++) this.refreshCell(r, c);
  }
  setMarks(marks) { this.marks = marks || {}; this.refreshAll(); }
  renderTabs() {
    this.tabs.innerHTML = "";
    for (const n of this.wb.sheetNames) {
      const b = el("button", { class: "tab" + (n === this.sheet ? " active" : ""), role: "tab", type: "button" }, n.replace(/_/g, " "));
      b.addEventListener("pointerdown", (e) => { if (this.editing) e.preventDefault(); });
      b.addEventListener("click", () => this.switchSheet(n));
      if (this.wb.extra[n] && !this.opts.readOnly) {
        const x = el("span", { class: "tabx", title: "Delete this output sheet" }, "×");
        x.addEventListener("click", (e) => {
          e.stopPropagation();
          if (confirm(`Delete the sheet "${n}"?`)) { this.wb.removeExtraSheet(n); this.switchSheet(this.wb.sheetNames[0]); }
        });
        b.append(x);
      }
      this.tabs.append(b);
    }
  }
  switchSheet(name) {
    if (this.editing && this.fx.value.startsWith("=")) { this.editSheetSwitch = name; }
    else if (this.editing) this.commit();
    this.sheet = name;
    this.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.anchor = { r: 0, c: 0 };
    this.renderSheet();
    this.updateNameBox();
    this.opts.onSheet && this.opts.onSheet(name);
  }
  onWbChange(info) {
    if (info.newSheet) { this.switchSheet(info.sheet); return; }
    if (info.removed) { this.renderTabs(); return; }
    if (info.spec) { this.renderCharts(); this.refreshAll(); this.opts.onChange && this.opts.onChange(); return; }
    if (info.full) this.renderSheet(); else this.refreshAll();
    if (!this.editing) this.updateNameBox();
    this.opts.onChange && this.opts.onChange();
  }

  // ------------------------------------------------------------------ selection
  inSel(r, c) { const s = this.sel; return r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2; }
  paintSelection() {
    if (!this.table) return;
    this.table.querySelectorAll("td.sel, td.anchor").forEach((t) => t.classList.remove("sel", "anchor"));
    const s = this.sel;
    for (let r = s.r1; r <= Math.min(s.r2, this.view.R - 1); r++)
      for (let c = s.c1; c <= Math.min(s.c2, this.view.C - 1); c++) { const t = this.td(r, c); t && t.classList.add("sel"); }
    const a = this.td(this.anchor.r, this.anchor.c);
    if (a) a.classList.add("anchor");
    this.table.querySelectorAll("th.hsel").forEach((t) => t.classList.remove("hsel"));
    for (let c = s.c1; c <= s.c2; c++) { const th = this.table.rows[0].cells[c + 1]; th && th.classList.add("hsel"); }
    for (let r = s.r1; r <= s.r2 && r < this.view.R; r++) { const th = this.table.rows[r + 1].cells[0]; th && th.classList.add("hsel"); }
    this.updateSelStats();
  }
  /** Excel status bar: Average / Count / Sum of the selected numbers. */
  updateSelStats() {
    if (!this.selStats) return;
    const s = this.sel;
    if (s.r1 === s.r2 && s.c1 === s.c2) { this.selStats.textContent = ""; return; }
    let n = 0, cnt = 0, sum = 0;
    for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
      const v = this.wb.value(this.sheet, r, c);
      if (v !== null && v !== "" && v !== undefined) cnt++;
      if (typeof v === "number") { n++; sum += v; }
    }
    const f = (x) => formatValue(Number(x.toPrecision(10)));
    this.selStats.replaceChildren(...(n ? [el("span", { class: "opt" }, `Average: ${f(sum / n)}`), el("span", {}, `Count: ${cnt}`), el("span", {}, `Sum: ${f(sum)}`)]
      : [el("span", {}, `Count: ${cnt}`)]));
  }
  select(r, c, extend) {
    if (extend) {
      this.sel = { r1: Math.min(this.anchor.r, r), r2: Math.max(this.anchor.r, r), c1: Math.min(this.anchor.c, c), c2: Math.max(this.anchor.c, c) };
    } else {
      this.anchor = { r, c };
      this.sel = { r1: r, c1: c, r2: r, c2: c };
    }
    this.paintSelection();
    this.updateNameBox();
    this.opts.onSelect && this.opts.onSelect(this.sheet, this.sel);
  }
  selRef(withSheet) { return (withSheet ? this.sheet + "!" : "") + WB.rangeStr(this.sel); }
  updateNameBox() {
    this.nameBox.textContent = WB.rangeStr(this.sel);
    if (!this.editing) {
      const raw = this.wb.raw(this.sheet, this.anchor.r, this.anchor.c);
      this.fx.value = raw === "" ? "" : String(raw);
    }
  }
  toggleSelectMode(force) {
    this.selectMode = force ?? !this.selectMode;
    this.selBtn.classList.toggle("on", this.selectMode);
    if (this.selectMode) toast("Select is ON: tap the last cell of the range.");
  }
  move(dr, dc) {
    const r = Math.max(0, this.anchor.r + dr), c = Math.max(0, this.anchor.c + dc);
    this.select(r, c, false);
    const t = this.td(r, c);
    t && t.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  // ------------------------------------------------------------------ pointer & keys
  onCellDown(e) {
    const td = e.target.closest("td[data-r]");
    if (!td) return;
    const r = +td.dataset.r, c = +td.dataset.c;
    if (this.editing && this.fx.value.trim().startsWith("=")) {
      e.preventDefault();                           // keep the keyboard open; insert a reference (Excel point mode)
      this.insertRef(r, c, e.shiftKey || this.selectMode);
      return;
    }
    if (this.editing) this.commit();
    const extend = e.shiftKey || this.selectMode;
    this.select(r, c, extend);
    if (extend && this.selectMode) this.toggleSelectMode(false);
    if (e.pointerType === "mouse") this.gridWrap.focus({ preventScroll: true });
  }
  onGridKey(e) {
    if (this.opts.readOnly || this.editing) return;
    const k = e.key, ctrl = e.ctrlKey || e.metaKey;
    const mv = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, 1] }[k];
    if (mv && !ctrl) {
      e.preventDefault();
      if (e.shiftKey && k.startsWith("Arrow")) {
        const er = (this.sel.r1 === this.anchor.r ? this.sel.r2 : this.sel.r1) + mv[0];
        const ec = (this.sel.c1 === this.anchor.c ? this.sel.c2 : this.sel.c1) + mv[1];
        this.select(Math.max(0, er), Math.max(0, ec), true);
      } else this.move(...mv);
      return;
    }
    if (ctrl && /^[dDrRcCvVzZ]$/.test(k)) {
      e.preventDefault();
      ({ d: () => this.doFill("down"), r: () => this.doFill("right"), c: () => { this.wb.copy(this.sheet, this.sel); toast("Copied."); },
        v: () => this.doPaste(), z: () => this.wb.undo() })[k.toLowerCase()]();
      return;
    }
    if (k === "Delete" || k === "Backspace") { e.preventDefault(); this.wb.clear(this.sheet, this.sel); return; }
    if (k === "F2") { e.preventDefault(); this.fx.focus(); return; }
    if (e.altKey && k === "=") { e.preventDefault(); this.autoSum(); return; }
    if (k.length === 1 && !ctrl && !e.altKey) {     // start typing: replace the cell content
      e.preventDefault();
      this.beginEdit(true);
      this.fx.value = k;
      this.fx.focus();
      this.fx.setSelectionRange(1, 1);
    }
  }

  // ------------------------------------------------------------------ editing
  beginEdit(replace) {
    if (this.opts.readOnly) return;
    if (!this.editing) {
      this.editing = true;
      this.editCell = { sheet: this.sheet, r: this.anchor.r, c: this.anchor.c };
      this.root.classList.add("is-editing");
      if (replace) this.fx.value = "";
    }
  }
  previewCell() {
    if (this.editCell.sheet !== this.sheet) return;
    const td = this.td(this.editCell.r, this.editCell.c);
    if (td) td.textContent = this.fx.value;
  }
  commit() {
    if (!this.editing) return;
    const { sheet, r, c } = this.editCell;
    this.editing = false;
    this.root.classList.remove("is-editing");
    this.lastPick = null;
    let text = this.fx.value;
    if (text.trim().startsWith("=")) {                    // auto-close parentheses like Excel
      const open = (text.match(/\(/g) || []).length - (text.match(/\)/g) || []).length;
      if (open > 0) text += ")".repeat(open);
    }
    const err = this.wb.setCell(sheet, r, c, text);
    if (err) toast("That formula has a problem: " + err.replace(/^.*?:\s*/, ""), "bad");
    if (this.sheet !== sheet) this.switchSheet(sheet);
    const v = this.wb.value(sheet, r, c);
    if (v && typeof v === "object" && v.type === "NAME") toast(`#NAME? Excel does not know a word in this formula. Check the spelling.`, "bad");
    this.fx.blur();
  }
  cancelEdit() {
    if (!this.editing) return;
    this.editing = false;
    this.root.classList.remove("is-editing");
    const { sheet } = this.editCell;
    if (this.sheet !== sheet) this.switchSheet(sheet);
    this.refreshCell(this.editCell.r, this.editCell.c);
    this.updateNameBox();
    this.fx.blur();
  }
  insertText(t) {
    const s = this.fx.selectionStart ?? this.fx.value.length, e = this.fx.selectionEnd ?? s;
    this.fx.value = this.fx.value.slice(0, s) + t + this.fx.value.slice(e);
    const p = s + t.length;
    this.fx.setSelectionRange(p, p);
    this.previewCell();
  }
  insertRef(r, c, extend) {
    const pre = this.sheet !== this.editCell.sheet ? (/[^\w.]/.test(this.sheet) ? `'${this.sheet}'!` : this.sheet + "!") : "";
    let ref = WB.addr(r, c);
    const caret = this.fx.selectionStart ?? this.fx.value.length;
    if (this.lastPick && this.lastPick.end === caret) {   // tapping again replaces the last inserted reference, or extends it
      const prev = this.fx.value.slice(this.lastPick.start, this.lastPick.end);
      if (extend) ref = prev.replace(/:.*$/, "") + ":" + ref;
      else ref = pre + ref;
      this.fx.value = this.fx.value.slice(0, this.lastPick.start) + ref + this.fx.value.slice(this.lastPick.end);
      this.lastPick = { start: this.lastPick.start, end: this.lastPick.start + ref.length };
      this.fx.setSelectionRange(this.lastPick.end, this.lastPick.end);
    } else {
      ref = pre + ref;
      this.insertText(ref);
      this.lastPick = { start: caret, end: caret + ref.length };
    }
    if (extend && this.selectMode) this.toggleSelectMode(false);
    this.fx.focus();
  }
  keyPress(k) {
    if (this.opts.readOnly) return;
        if (!this.editing) { this.beginEdit(false); this.fx.focus(); }
    if (k === "$ F4") {
      const res = toggleRefAt(this.fx.value, this.fx.selectionStart ?? this.fx.value.length);
      if (res) { this.fx.value = res.text; this.fx.setSelectionRange(res.caret, res.caret); this.previewCell(); }
      else toast("Put the cursor on a cell reference (like G8), then press F4.");
      this.lastPick = null;
      return;
    }
    this.insertText(k);
    this.lastPick = null;
  }
  funcPicker() {
    const list = el("div", { class: "funclist" });
    for (const f of FUNCS) {
      const b = el("button", { type: "button" }, f);
      b.addEventListener("click", () => {
        m.close();
        if (!this.editing) { this.beginEdit(!this.fx.value.startsWith("=")); }
        this.fx.focus();
        if (!this.fx.value.startsWith("=")) { this.fx.value = "="; this.fx.setSelectionRange(1, 1); }
        this.insertText(f + "(");
      });
      list.append(b);
    }
    const m = modal("Insert a function", list);
  }
  autoSum() {                                          // Alt + =  : SUM of the numbers just above
    const { r, c } = this.anchor;
    let top = r - 1;
    while (top >= 0 && typeof this.wb.value(this.sheet, top, c) === "number") top--;
    const r1 = top + 1;
    this.beginEdit(true);
    this.fx.focus();
    this.fx.value = r1 <= r - 1 ? `=SUM(${WB.addr(r1, c)}:${WB.addr(r - 1, c)})` : "=SUM(";
    this.fx.setSelectionRange(this.fx.value.length, this.fx.value.length);
    this.previewCell();
  }
  doFill(kind) {
    const err = kind === "down" ? this.wb.fillDown(this.sheet, this.sel) : kind === "right" ? this.wb.fillRight(this.sheet, this.sel)
      : this.wb.fillToEnd(this.sheet, this.sel);
    if (err) toast(err);
  }
  doPaste() {
    this.wb.pasteTarget = this.sel.r1 === this.sel.r2 && this.sel.c1 === this.sel.c2 ? null : this.sel;
    const err = this.wb.paste(this.sheet, this.sel.r1, this.sel.c1);
    this.wb.pasteTarget = null;
    if (err) toast(err);
  }

  // ------------------------------------------------------------------ dialogs
  formatDialog() {
    const opts = [["", "General  1234.5678"], ["#,##0", "Comma, no decimals  1,234"], ["#,##0.00", "Comma Style  1,234.56"], ["0.00", "Number, 2 decimals  1234.56"],
      ["0.000", "Number, 3 decimals  0.123"], ["0%", "Percent (Ctrl + Shift + %)  56%"], ["0.0%", "Percent, 1 decimal  56.0%"]];
    const box = el("div", { class: "optlist" });
    for (const [nf, label] of opts) {
      const [name, sample] = label.split(/\s{2,}/);
      const b = el("button", { type: "button" }, el("span", {}, name), el("span", {}, sample || ""));
      b.addEventListener("click", () => { this.wb.setFormat(this.sheet, this.sel, nf || null); m.close(); });
      box.append(b);
    }
    const m = modal(`Format ${this.selRef()}`, box);
  }

  chartDialog(existing) {
    const s = existing || { type: "column", range: this.selRef(), gap: 150, trendline: false, title: "", sheet: this.sheet };
    const f = el("form", { class: "form" });
    f.innerHTML = `
      <label>Chart type <select name="type">
        <option value="column">Column (use gap 5% for a histogram)</option><option value="line">Line</option><option value="scatter">Scatter (X Y)</option></select></label>
      <label>Data range <input name="range" value="${esc(s.range || "")}" placeholder="e.g. C3:D14"></label>
      <p class="hint">Column/Line: first column = labels, next columns = values. Scatter: first column = X, second column = Y.
        A text first row is used as series names.</p>
      <label>Gap width % <input name="gap" type="number" min="0" max="500" value="${s.gap ?? 150}"></label>
      <label class="row"><input name="trendline" type="checkbox" ${s.trendline ? "checked" : ""}> Add a linear trendline (show equation and R²)</label>
      <label>Chart title <input name="title" value="${esc(s.title || "")}"></label>
      <div class="actions">${existing ? '<button type="button" class="btn danger" data-del>Delete chart</button>' : ""}
        <button class="btn" type="button" data-cancel>Cancel</button>
        <button class="btn primary" type="submit">${existing ? "Update" : "Insert chart"}</button></div>`;
    f.type.value = s.type;
    const m = modal(existing ? "Edit chart" : "Insert chart", f);
    f.querySelector("[data-del]")?.addEventListener("click", () => { this.wb.removeSpec("charts", existing.id); m.close(); });
    f.addEventListener("submit", (e) => {
      e.preventDefault();
      const range = f.range.value.trim();
      const g = WB.parseRef(range, this.sheet);
      if (!g || this.wb.sid(g.sheet) === undefined) return toast("Type a valid range, e.g. C3:D14", "bad");
      if (g.c2 === g.c1 && f.type.value === "scatter") return toast("A scatter chart needs two columns: X and Y.", "bad");
      const spec = this.chartSeries({ ...s, type: f.type.value, range, gap: Number(f.gap.value), trendline: f.trendline.checked,
        title: f.title.value.trim(), sheet: s.sheet || this.sheet });
      if (!spec) return toast("No numbers found in that range.", "bad");
      this.wb.addSpec("charts", spec);
      m.close();
      toast("Chart added below the sheet.");
    });
  }
  /** Work out labels/series ranges for a chart spec (stored as ranges so the trainer and the checker can read them). */
  chartSeries(spec) {
    const g = WB.parseRef(spec.range, spec.sheet);
    const sh = g.sheet || spec.sheet;
    const isNum = (r, c) => typeof this.wb.value(sh, r, c) === "number";
    let r1 = g.r1;
    const header = g.r2 > g.r1 && !isNum(g.r1, g.c2);
    if (header) r1++;
    let labelCol = null, valCols = [];
    if (spec.type === "scatter") { labelCol = g.c1; valCols = [g.c1 + 1 <= g.c2 ? g.c1 + 1 : g.c1]; }
    else if (g.c2 > g.c1 && !isNum(r1, g.c1)) { labelCol = g.c1; for (let c = g.c1 + 1; c <= g.c2; c++) valCols.push(c); }
    else if (g.c2 > g.c1 && spec.type === "column" && g.c2 - g.c1 === 1) { labelCol = g.c1; valCols = [g.c2]; }
    else for (let c = g.c1; c <= g.c2; c++) valCols.push(c);
    const rng = (c) => `${WB.addr(r1, c)}:${WB.addr(g.r2, c)}`;
    const ranges = valCols.map(rng);
    if (!ranges.length) return null;
    return { ...spec, sheet: sh, ranges, labels: labelCol !== null ? rng(labelCol) : null,
      names: valCols.map((c) => (header ? String(this.wb.value(sh, g.r1, c) ?? "") : `Series ${c - g.c1 + 1}`)) };
  }
  renderCharts() {
    this.charts.forEach((c) => c.destroy());
    this.charts.clear();
    this.chartArea.innerHTML = "";
    const list = this.wb.specs.charts.filter((c) => c.sheet === this.sheet);
    for (const spec of list) {
      const card = el("div", { class: "chartcard" });
      const cv = el("canvas");
      card.append(cv);
      if (!this.opts.readOnly) {
        const ed = el("button", { class: "btn small", type: "button" }, "Edit chart");
        ed.addEventListener("click", () => this.chartDialog(spec));
        card.append(ed);
      }
      this.chartArea.append(card);
      try { this.charts.set(spec.id, this.drawChart(cv, spec)); } catch (e) { card.append(el("p", { class: "hint" }, "Chart error: " + e.message)); }
    }
  }
  drawChart(canvas, spec) { return drawChartOn(this.wb, canvas, spec); }

  analysisDialog() {
    const f = el("form", { class: "form" });
    f.innerHTML = `
      <label>Analysis tool <select name="tool">
        <option value="descr">Descriptive Statistics</option>
        <option value="ttest">t-Test: Two-Sample Assuming Unequal Variances</option>
        <option value="regression">Regression</option></select></label>
      <label><span data-l1>Input Range</span> <input name="input" value="${esc(this.selRef())}" placeholder="$C$1:$F$91"></label>
      <label data-two hidden><span data-l2>Variable 2 Range</span> <input name="input2" placeholder="$C$1:$C$91"></label>
      <p class="hint" data-hint>Grouped by columns. Ranges are on the sheet <b>${esc(this.sheet)}</b> unless you type Sheet!A1:B2.</p>
      <label class="row"><input name="labels" type="checkbox" checked> Labels in first row</label>
      <label class="row"><input name="summary" type="checkbox" checked> Summary statistics</label>
      <label>New Worksheet Ply <input name="out" value=""></label>
      <div class="actions"><button class="btn" type="button" data-cancel>Cancel</button><button class="btn primary" type="submit">OK</button></div>`;
    const sync = () => {
      const t = f.tool.value;
      f.querySelector("[data-two]").hidden = t === "descr";
      f.querySelector("[data-l1]").textContent = t === "ttest" ? "Variable 1 Range" : t === "regression" ? "Input Y Range" : "Input Range";
      f.querySelector("[data-l2]").textContent = t === "regression" ? "Input X Range" : "Variable 2 Range";
      f.summary.parentElement.hidden = t !== "descr";
      f.out.placeholder = { descr: "ToolPak", ttest: "t-Test", regression: "Regression" }[t];
    };
    f.tool.addEventListener("change", sync);
    sync();
    const m = modal("Data Analysis", f);
    f.addEventListener("submit", (e) => {
      e.preventDefault();
      const tool = f.tool.value, labels = f.labels.checked;
      const outName = f.out.value.trim() || f.out.placeholder;
      const src = this.sheet;                 // the data's sheet: addOutputSheet() below switches to the new output sheet
      try {
        const a = this.wb.columns(f.input.value, labels, this.sheet);
        let rows, colW = { A: 30, B: 18, C: 18, D: 18, E: 18, F: 18, G: 18, H: 18 };
        if (tool === "descr") {
          rows = ST.descriptive(a.cols.filter((c) => c.values.length > 0));
          colW = {};
          a.cols.forEach((_, i) => { colW[WB.numToCol(2 * i + 1)] = 22; colW[WB.numToCol(2 * i + 2)] = 16; });
        } else {
          const b = this.wb.columns(f.input2.value, labels, this.sheet);
          if (tool === "ttest") rows = ST.tTestUnequal(a.cols[0].values, b.cols[0].values, a.cols[0].label, b.cols[0].label);
          else rows = ST.regression(a.cols[0].values, b.cols[0].values, a.cols[0].label, labels ? b.cols[0].label : "X Variable 1");
        }
        const qualify = (ref) => (ref.includes("!") ? ref : `${src}!${ref}`);
        const name = this.wb.addOutputSheet(outName, rows, colW);
        this.wb.addSpec("analyses", { tool, sheet: src, input: qualify(f.input.value.trim()),
          input2: tool === "descr" ? null : qualify(f.input2.value.trim()), labels, out: name });
        m.close();
        toast(`Done: results are on the new sheet "${name}". Remember: ToolPak output does not update when the data changes.`);
      } catch (err) {
        toast(err.message, "bad");
      }
    });
  }

  highlightDialog() {
    const f = el("form", { class: "form" });
    f.innerHTML = `
      <p class="hint">Highlight Cells Rules › Equal To: cells in the range whose value equals the text are filled with colour.</p>
      <label>Apply to range <input name="range" value="${esc(this.selRef())}"></label>
      <label>Format cells that are EQUAL TO <input name="value" value="OUTLIER"></label>
      <label>with <select name="color"><option value="red">Light red fill</option><option value="yellow">Yellow fill</option><option value="green">Green fill</option></select></label>
      <div class="rules"></div>
      <div class="actions"><button class="btn" type="button" data-cancel>Cancel</button><button class="btn primary" type="submit">OK</button></div>`;
    const rules = f.querySelector(".rules");
    const listRules = () => {
      rules.innerHTML = "";
      this.wb.specs.rules.filter((r) => r.sheet === this.sheet).forEach((r) => {
        const b = el("button", { type: "button", title: "Delete this rule" }, `${r.range} = "${r.value}"  ×`);
        b.addEventListener("click", () => { this.wb.removeSpec("rules", r.id); listRules(); });
        rules.append(b);
      });
    };
    listRules();
    const m = modal("Highlight cells", f);
    f.addEventListener("submit", (e) => {
      e.preventDefault();
      const g = WB.parseRef(f.range.value.trim(), this.sheet);
      if (!g) return toast("Type a valid range, e.g. C6:C95", "bad");
      this.wb.addSpec("rules", { sheet: this.sheet, range: WB.rangeStr(g), value: f.value.value, color: f.color.value });
      m.close();
    });
  }
}

// Draw a chart spec on a canvas (used on screen and, with {print:true}, for the exported workbook).
export function drawChartOn(wb, canvas, spec, opt = {}) {
  const vals = (rng) => {
    const g = WB.parseRef(rng, spec.sheet), out = [];
    for (let r = g.r1; r <= g.r2; r++) out.push(wb.value(spec.sheet, r, g.c1));
    return out;
  };
  const css = getComputedStyle(document.documentElement);
  const ink = opt.print ? "#201f1e" : css.getPropertyValue("--ink").trim() || "#201f1e";
  const grid = opt.print ? "#e1dfdd" : css.getPropertyValue("--line").trim() || "#e1dfdd";
  const colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000"];          // Excel's default series colours
  let data, options = { responsive: !opt.print, maintainAspectRatio: false, animation: false, devicePixelRatio: opt.print ? 2 : undefined,
    plugins: { legend: { labels: { color: ink } }, title: { display: !!spec.title, text: spec.title, color: ink, font: { size: 15 } } },
    scales: { x: { ticks: { color: ink }, grid: { color: grid } }, y: { ticks: { color: ink }, grid: { color: grid } } } };
  if (spec.type === "scatter") {
    const xs = vals(spec.labels), ys = vals(spec.ranges[0]);
    const pts = xs.map((x, i) => ({ x, y: ys[i] })).filter((p) => typeof p.x === "number" && typeof p.y === "number");
    const ds = [{ label: spec.names[0] || "Y", data: pts, backgroundColor: colors[0], pointRadius: 4, type: "scatter" }];
    if (spec.trendline && pts.length > 1) {
      const L = ST.fitLine(pts.map((p) => p.x), pts.map((p) => p.y));
      const xmin = Math.min(...pts.map((p) => p.x)), xmax = Math.max(...pts.map((p) => p.x));
      ds.push({ label: `y = ${L.slope.toFixed(4)}x ${L.intercept < 0 ? "−" : "+"} ${Math.abs(L.intercept).toFixed(2)}   R² = ${L.r2.toFixed(4)}`,
        data: [{ x: xmin, y: L.intercept + L.slope * xmin }, { x: xmax, y: L.intercept + L.slope * xmax }],
        type: "line", borderColor: colors[1], borderDash: [6, 4], pointRadius: 0, borderWidth: 2 });
    }
    data = { datasets: ds };
    options.scales.x.type = "linear";
  } else {
    const labels = spec.labels ? vals(spec.labels).map((v) => formatValue(v)) : vals(spec.ranges[0]).map((_, i) => String(i + 1));
    const pct = 1 / (1 + Math.max(0, Number(spec.gap ?? 150)) / 100);
    data = { labels, datasets: spec.ranges.map((rg, i) => ({ label: spec.names[i] || `Series ${i + 1}`, data: vals(rg).map((v) => (typeof v === "number" ? v : null)),
      backgroundColor: colors[i % 4], borderColor: colors[i % 4], categoryPercentage: 1, barPercentage: pct, tension: 0.2 })) };
    if (spec.trendline && spec.type === "line" && data.datasets[0]) {
      const ys = data.datasets[0].data, xs = ys.map((_, i) => i + 1);
      const ok = xs.filter((_, i) => typeof ys[i] === "number");
      const L = ST.fitLine(ok, ok.map((x) => ys[x - 1]));
      data.datasets.push({ label: `Trend  R² = ${L.r2.toFixed(4)}`, data: xs.map((x) => L.intercept + L.slope * x), borderColor: colors[1], borderDash: [6, 4], pointRadius: 0 });
    }
    options.plugins.legend.display = spec.ranges.length > 1 || !!spec.trendline;
  }
  return new window.Chart(canvas, { type: spec.type === "scatter" ? "scatter" : spec.type === "line" ? "line" : "bar", data, options });
}
