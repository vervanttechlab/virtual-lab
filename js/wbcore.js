/* Workbook core shared by the browser app and the build scripts (classic script / CommonJS).
   Sheet JSON: {name, rows, cols, colW:{A:w}, cells:{A1:[valueOrFormula, style, numFmt]}}  style: i=input h=header g=given */
(function (root) {
  const colToNum = (s) => s.toUpperCase().split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); // A=1
  const numToCol = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  function parseAddr(a) {                  // "B4" -> {r:3, c:1} (0-based); tolerates $
    const m = String(a).replace(/\$/g, "").toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    return { r: Number(m[2]) - 1, c: colToNum(m[1]) - 1 };
  }
  const addr = (r, c) => numToCol(c + 1) + (r + 1);

  function parseRef(ref, defSheet) {       // "Stats!B4" | "'Daily Sales'!B4:C9" | "B4"
    ref = String(ref).trim();
    let sheet = defSheet || null, rest = ref;
    const i = ref.lastIndexOf("!");
    if (i >= 0) { sheet = ref.slice(0, i).replace(/^'|'$/g, ""); rest = ref.slice(i + 1); }
    const [a, b] = rest.split(":");
    const s = parseAddr(a), e = parseAddr(b || a);
    if (!s || !e) return null;
    return { sheet, r1: Math.min(s.r, e.r), c1: Math.min(s.c, e.c), r2: Math.max(s.r, e.r), c2: Math.max(s.c, e.c) };
  }
  const rangeStr = (g) => addr(g.r1, g.c1) + (g.r1 === g.r2 && g.c1 === g.c2 ? "" : ":" + addr(g.r2, g.c2));
  const contains = (outer, inner) => !!outer && !!inner && outer.r1 <= inner.r1 && outer.c1 <= inner.c1 && outer.r2 >= inner.r2 && outer.c2 >= inner.c2;

  /** 2-D array for HyperFormula from a sheet JSON plus edits {"A1": "=..."} (edits win). */
  function sheetArray(sh, edits) {
    let rows = sh.rows, cols = sh.cols;
    const put = [];
    for (const [a, v] of Object.entries(sh.cells)) put.push([a, v[0]]);
    for (const [a, v] of Object.entries(edits || {})) put.push([a, v]);
    for (const [a] of put) { const p = parseAddr(a); if (p) { rows = Math.max(rows, p.r + 1); cols = Math.max(cols, p.c + 1); } }
    const arr = Array.from({ length: rows }, () => Array(cols).fill(null));
    for (const [a, v] of put) { const p = parseAddr(a); if (p) arr[p.r][p.c] = v === undefined ? null : v; }
    return arr;
  }

  /** Build a HyperFormula instance for a day. edits = {"Sheet!A1": raw}; extra = {sheetName: sheetJSON} (ToolPak outputs). */
  function buildHF(lib, sheets, edits, extra) {
    const per = {};
    for (const [ref, v] of Object.entries(edits || {})) {
      const i = ref.lastIndexOf("!");
      (per[ref.slice(0, i)] = per[ref.slice(0, i)] || {})[ref.slice(i + 1)] = v;
    }
    const data = {};
    for (const sh of sheets) data[sh.name] = sheetArray(sh, per[sh.name]);
    for (const [name, sh] of Object.entries(extra || {})) data[name] = sheetArray(sh, per[name]);
    const ne = (root.installHFExtras && root.installHFExtras.namedExpressions) || (lib.__namedExpressions || []);
    return lib.HyperFormula.buildFromSheets(data, { licenseKey: "gpl-v3", useArrayArithmetic: false, maxRows: 5000, maxColumns: 60 }, ne);
  }

  const api = { colToNum, numToCol, parseAddr, addr, parseRef, rangeStr, contains, sheetArray, buildHF };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.WBCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
