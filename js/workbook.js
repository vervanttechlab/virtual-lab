// Workbook model: HyperFormula engine + the day's template sheets + the student's edits, extra sheets
// (ToolPak outputs), charts, highlight rules and number formats. Serialises to a small JSON for saving/submitting.
const WB = window.WBCore;
// The spreadsheet engine (HyperFormula, ~0.7 MB) is loaded on demand by loadEngine() in app.js.

export class Workbook {
  constructor(day, saved) {
    this.day = day;
    saved = saved || {};
    this.extra = saved.extra || {};                 // {name: sheetJSON}
    this.specs = { charts: [], analyses: [], rules: [], ...(saved.specs || {}) };
    this.fmt = saved.fmt || {};                     // {"Sheet!A1": numFmt}
    this.edits = {};
    this.meta = {};
    for (const s of day.sheets) this.meta[s.name] = s;
    for (const [n, s] of Object.entries(this.extra)) this.meta[n] = s;
    const HFLib = window.HyperFormula;
    if (!HFLib) throw new Error("The spreadsheet engine is not loaded yet.");
    if (window.installHFExtras && !HFLib.HyperFormula?.__dal3Extras) window.installHFExtras(HFLib);
    this.hf = WB.buildHF(HFLib, day.sheets, saved.edits || {}, this.extra);
    this.edits = { ...(saved.edits || {}) };
    this.undoStack = [];
    this.listeners = new Set();
  }

  // ------------------------------------------------------------ basics
  get sheetNames() { return this.hf.getSheetNames(); }
  sid(name) { return this.hf.getSheetId(name); }
  onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  _emit(info) { this.listeners.forEach((fn) => fn(info)); }
  dims(sheet) {
    const d = this.hf.getSheetDimensions(this.sid(sheet));
    const m = this.meta[sheet] || {};
    return { rows: Math.max(d.height, m.rows || 0, 1), cols: Math.max(d.width, m.cols || 0, 1) };
  }
  style(sheet, r, c) {
    const m = this.meta[sheet];
    const cell = m && m.cells[WB.addr(r, c)];
    return (cell && cell[1]) || "";
  }
  numFmt(sheet, r, c) {
    const k = `${sheet}!${WB.addr(r, c)}`;
    if (k in this.fmt) return this.fmt[k];
    const m = this.meta[sheet];
    const cell = m && m.cells[WB.addr(r, c)];
    return (cell && cell[2]) || "";
  }
  raw(sheet, r, c) {
    const v = this.hf.getCellSerialized({ sheet: this.sid(sheet), row: r, col: c });
    return v == null ? "" : v;
  }
  value(sheet, r, c) { return this.hf.getCellValue({ sheet: this.sid(sheet), row: r, col: c }); }
  valueOfRef(ref) {
    const g = WB.parseRef(ref);
    return g ? this.value(g.sheet, g.r1, g.c1) : null;
  }
  isFormula(sheet, r, c) { return this.hf.doesCellHaveFormula({ sheet: this.sid(sheet), row: r, col: c }); }
  templateRaw(sheet, r, c) {
    const m = this.day.sheets.find((s) => s.name === sheet) || this.extra[sheet];
    const cell = m && m.cells[WB.addr(r, c)];
    const v = cell ? cell[0] : null;
    return v == null ? "" : v;
  }

  // ------------------------------------------------------------ editing
  _snapshot(sheet, g) {
    const cells = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) cells.push([r, c, this.raw(sheet, r, c)]);
    return { sheet, cells };
  }
  _record(sheet, g) {                                // sync this.edits for a rectangle after a change
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const k = `${sheet}!${WB.addr(r, c)}`;
      const now = this.raw(sheet, r, c), base = this.templateRaw(sheet, r, c);
      if (String(now) === String(base)) delete this.edits[k]; else this.edits[k] = now;
    }
  }
  _parseInput(text) {
    if (typeof text !== "string") return text;
    const t = text.trim();
    if (t === "") return null;
    if (t.startsWith("=")) return t;
    const n = window.DAL3Check.parseNum(t);
    if (/^[-+]?[\d,.]+%?$|^[-+]?\d*\.?\d+e[-+]?\d+$/i.test(t.replace(/[₱\s]/g, "")) && isFinite(n)) return n;
    if (/^true$/i.test(t)) return true;
    if (/^false$/i.test(t)) return false;
    return text;
  }
  setCell(sheet, r, c, text) {
    const g = { r1: r, c1: c, r2: r, c2: c };
    this.undoStack.push(this._snapshot(sheet, g));
    let err = null;
    try {
      this.hf.setCellContents({ sheet: this.sid(sheet), row: r, col: c }, [[this._parseInput(text)]]);
    } catch (e) {
      err = e.message || String(e);
      this.undoStack.pop();
    }
    this._record(sheet, g);
    this._emit({ sheet, range: g });
    return err;
  }
  clear(sheet, g) {
    this.undoStack.push(this._snapshot(sheet, g));
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++)
      this.hf.setCellContents({ sheet: this.sid(sheet), row: r, col: c }, [[null]]);
    this._record(sheet, g);
    this._emit({ sheet, range: g });
  }
  /** Copy the source rectangle and paste it at the top-left of target (relative refs shift, like Excel). */
  _copyPaste(sheet, src, tr, tc) {
    const sid = this.sid(sheet);
    this.hf.copy({ start: { sheet: sid, row: src.r1, col: src.c1 }, end: { sheet: sid, row: src.r2, col: src.c2 } });
    this.hf.paste({ sheet: sid, row: tr, col: tc });
  }
  fillDown(sheet, g) {                               // Ctrl + D
    if (g.r2 <= g.r1) return "Select at least two rows: the first row is copied down.";
    this.undoStack.push(this._snapshot(sheet, g));
    // (HyperFormula does not allow copy/paste inside hf.batch, so paste row by row)
    for (let r = g.r1 + 1; r <= g.r2; r++) this._copyPaste(sheet, { r1: g.r1, r2: g.r1, c1: g.c1, c2: g.c2 }, r, g.c1);
    this._record(sheet, g);
    this._emit({ sheet, range: g });
    return null;
  }
  fillRight(sheet, g) {                              // Ctrl + R
    if (g.c2 <= g.c1) return "Select at least two columns: the first column is copied right.";
    this.undoStack.push(this._snapshot(sheet, g));
    for (let c = g.c1 + 1; c <= g.c2; c++) this._copyPaste(sheet, { r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c1 }, g.r1, c);
    this._record(sheet, g);
    this._emit({ sheet, range: g });
    return null;
  }
  /** Fill down to the last row that has data in the column to the left (Excel: double-click the fill handle). */
  fillToEnd(sheet, g) {
    const { rows } = this.dims(sheet);
    const lc = g.c1 > 0 ? g.c1 - 1 : g.c2 + 1;
    let last = g.r1;
    for (let r = g.r1 + 1; r < rows; r++) { const v = this.raw(sheet, r, lc); if (v === "" || v == null) break; last = r; }
    if (last === g.r1) return "There is no data in the next column to fill along.";
    return this.fillDown(sheet, { r1: g.r1, r2: last, c1: g.c1, c2: g.c2 });
  }
  copy(sheet, g) { this.clip = { sheet, g: { ...g } }; }
  paste(sheet, r, c) {
    if (!this.clip) return "Copy something first.";
    const h = this.clip.g.r2 - this.clip.g.r1, w = this.clip.g.c2 - this.clip.g.c1;
    // Excel: a bigger target selection repeats the copied block across it
    const tg = this.pasteTarget || { r1: r, c1: c, r2: r + h, c2: c + w };
    const target = { r1: tg.r1, c1: tg.c1, r2: Math.max(tg.r2, tg.r1 + h), c2: Math.max(tg.c2, tg.c1 + w) };
    this.undoStack.push(this._snapshot(sheet, target));
    const sid = this.sid(sheet), csid = this.sid(this.clip.sheet);
    for (let rr = target.r1; rr + h <= target.r2; rr += h + 1) for (let cc = target.c1; cc + w <= target.c2; cc += w + 1) {
      this.hf.copy({ start: { sheet: csid, row: this.clip.g.r1, col: this.clip.g.c1 }, end: { sheet: csid, row: this.clip.g.r2, col: this.clip.g.c2 } });
      this.hf.paste({ sheet: sid, row: rr, col: cc });
    }
    this._record(sheet, target);
    this._emit({ sheet, range: target });
    return null;
  }
  undo() {
    const s = this.undoStack.pop();
    if (!s) return false;
    const sid = this.sid(s.sheet);
    if (sid === undefined) return false;
    this.hf.batch(() => { for (const [r, c, v] of s.cells) this.hf.setCellContents({ sheet: sid, row: r, col: c }, [[v === "" ? null : v]]); });
    const rs = s.cells.map((x) => x[0]), cs = s.cells.map((x) => x[1]);
    const g = { r1: Math.min(...rs), r2: Math.max(...rs), c1: Math.min(...cs), c2: Math.max(...cs) };
    this._record(s.sheet, g);
    this._emit({ sheet: s.sheet, range: g });
    return true;
  }
  setFormat(sheet, g, nf) {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const k = `${sheet}!${WB.addr(r, c)}`;
      if (nf === null) delete this.fmt[k]; else this.fmt[k] = nf;
    }
    this._emit({ sheet, range: g });
  }
  resetSheet(sheet) {
    const { rows, cols } = this.dims(sheet);
    const g = { r1: 0, c1: 0, r2: rows - 1, c2: cols - 1 };
    this.undoStack.push(this._snapshot(sheet, g));
    const tmpl = this.day.sheets.find((s) => s.name === sheet);
    if (!tmpl) return;
    this.hf.setSheetContent(this.sid(sheet), WB.sheetArray(tmpl, {}));
    for (const k of Object.keys(this.edits)) if (k.startsWith(sheet + "!")) delete this.edits[k];
    this._emit({ sheet, range: g, full: true });
  }

  // ------------------------------------------------------------ ranges / extra sheets
  /** Numbers in a ref like "Daily_Sales!C1:F91" -> [{label, values}] by column (labels = first row if asked). */
  columns(ref, labels, defSheet) {
    const g = WB.parseRef(ref, defSheet);
    if (!g || this.sid(g.sheet) === undefined) throw new Error(`Range not found: ${ref}`);
    const out = [];
    for (let c = g.c1; c <= g.c2; c++) {
      let label = `Column ${c - g.c1 + 1}`, r0 = g.r1;
      if (labels) { label = String(this.value(g.sheet, g.r1, c) ?? label); r0++; }
      const values = [];
      for (let r = r0; r <= g.r2; r++) { const v = this.value(g.sheet, r, c); if (typeof v === "number") values.push(v); }
      out.push({ label, values });
    }
    return { g, cols: out };
  }
  addOutputSheet(name, rows, colW) {
    let n = (name || "Output").replace(/[!'\[\]*?:/\\]/g, "").slice(0, 28) || "Output";
    let k = 2;
    const base = n;
    while (this.sid(n) !== undefined) n = `${base} (${k++})`;
    const cells = {};
    rows.forEach((row, r) => row.forEach((v, c) => { if (v !== null && v !== undefined) cells[WB.addr(r, c)] = [v]; }));
    const sh = { name: n, rows: rows.length, cols: Math.max(...rows.map((x) => x.length), 1), colW: colW || {}, cells };
    this.extra[n] = sh;
    this.meta[n] = sh;
    this.hf.addSheet(n);
    this.hf.setSheetContent(this.sid(n), WB.sheetArray(sh, {}));
    this._emit({ sheet: n, newSheet: true });
    return n;
  }
  removeExtraSheet(name) {
    if (!this.extra[name]) return;
    this.hf.removeSheet(this.sid(name));
    delete this.extra[name];
    delete this.meta[name];
    for (const k of Object.keys(this.edits)) if (k.startsWith(name + "!")) delete this.edits[k];
    this.specs.analyses = this.specs.analyses.filter((a) => a.out !== name);
    this._emit({ removed: name });
  }

  // ------------------------------------------------------------ specs
  addSpec(kind, spec) {
    spec.id = spec.id || Math.random().toString(36).slice(2, 9);
    const list = this.specs[kind];
    const i = list.findIndex((x) => x.id === spec.id);
    if (i >= 0) list[i] = spec; else list.push(spec);
    this._emit({ spec: kind });
    return spec;
  }
  removeSpec(kind, id) {
    this.specs[kind] = this.specs[kind].filter((x) => x.id !== id);
    this._emit({ spec: kind });
  }
  highlightOf(sheet, r, c) {
    for (const rule of this.specs.rules) {
      if (rule.sheet !== sheet) continue;
      const g = WB.parseRef(rule.range);
      if (!g || r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) continue;
      const v = this.value(sheet, r, c);
      if (window.DAL3Check.normText(v) === window.DAL3Check.normText(rule.value)) return rule.color || "red";
    }
    return null;
  }

  // ------------------------------------------------------------ save / submit
  serialize() { return { v: 1, edits: this.edits, extra: this.extra, specs: this.specs, fmt: this.fmt }; }
  /** {ref:{f,v}} for a list of refs (f = what the student typed, v = computed value) */
  cellsFor(refs) {
    const out = {};
    for (const ref of refs) {
      const g = WB.parseRef(ref);
      if (!g || this.sid(g.sheet) === undefined) continue;
      const raw = this.raw(g.sheet, g.r1, g.c1);
      let v = this.value(g.sheet, g.r1, g.c1);
      if (v && typeof v === "object") v = v.value;     // error -> "#DIV/0!" etc.
      out[ref] = { f: typeof raw === "string" ? raw : (raw === "" ? "" : String(raw)), v: v ?? "" };
    }
    return out;
  }
}

// ------------------------------------------------------------ display formatting (Excel-like)
export function formatValue(v, nf) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "object") return v.value || "#ERROR!";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v !== "number") return String(v);
  if (nf) {
    const pct = nf.endsWith("%");
    const dec = ((nf.split(".")[1] || "").match(/0/g) || []).length;
    const comma = nf.includes(",");
    const x = pct ? v * 100 : v;
    const s = x.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: comma });
    return pct ? s + "%" : s;
  }
  if (Number.isInteger(v)) return String(v);
  const a = Math.abs(v);
  if (a !== 0 && (a < 1e-4 || a >= 1e11)) return v.toExponential(4).replace("e", "E").replace(/E\+?/, "E+").replace("E+-", "E-");
  return String(Number(v.toPrecision(10)));
}
