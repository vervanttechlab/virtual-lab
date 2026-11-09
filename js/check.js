/* Answer normalisation, spec checks and scoring. Shared, byte-for-byte, by:
   the browser app (window.DAL3Check), the build step (require) and the Apps Script server (Check.gs, global DAL3Check).
   No dependencies; keep it ES2019 so Apps Script V8 runs it unchanged. */
(function (root) {
  // ------------------------------------------------------------- normalisation
  function normText(s) {
    return String(s == null ? "" : s).replace(/[‒-―−]/g, "-").replace(/\s+/g, " ").trim().toLowerCase();
  }
  /** Value of a spreadsheet cell -> comparison string (7 significant digits for numbers). */
  function normCell(v) {
    if (v && typeof v === "object" && "value" in v) v = v.value;         // HyperFormula DetailedCellError
    if (typeof v === "number") {
      if (!isFinite(v)) return "#num";
      if (v === 0) return "0";
      const r = Number(v.toPrecision(7));
      return String(r === 0 ? 0 : r);
    }
    if (typeof v === "boolean") return v ? "true" : "false";
    if (v == null) return "";
    const s = String(v);
    return s.charAt(0) === "#" ? s.toLowerCase() : normText(s);
  }
  /** Student-typed number: "1,234.5", "40%", "8/20", "₱ 3.87E-05", "−2.25" -> Number or NaN. */
  function parseNum(s) {
    if (typeof s === "number") return s;
    s = String(s == null ? "" : s).replace(/[−‒-―]/g, "-").replace(/[₱,\s]/g, "").replace(/^php/i, "");
    if (!s) return NaN;
    let pct = false;
    if (s.slice(-1) === "%") { pct = true; s = s.slice(0, -1); }
    let x;
    const f = s.match(/^(-?\d+(?:\.\d+)?)\/(\d+(?:\.\d+)?)$/);
    if (f) x = Number(f[1]) / Number(f[2]);
    else x = /^-?(\d+\.?\d*|\.\d+)(e-?\d+)?$/i.test(s) ? Number(s) : NaN;
    return pct ? x / 100 : x;
  }
  function normNum(x, part) {
    x = parseNum(x);
    if (!isFinite(x)) return "";
    if (part && part.sig) { const r = Number(x.toPrecision(part.sig)); return String(r === 0 ? 0 : r); }
    const dp = part && part.dp != null ? part.dp : 2;
    let r = Number(x.toFixed(dp));
    if (r === 0) r = 0;
    return r.toFixed(dp);
  }
  function normList(s) {
    const xs = Array.isArray(s) ? s : String(s == null ? "" : s).split(/[,;\s]+/).filter(Boolean);
    return xs.map((x) => normNum(x, { dp: 2 })).join(",");
  }
  function normMulti(arr) { return (arr || []).map(String).sort().join("|"); }
  function normFormula(s) { return String(s == null ? "" : s).toUpperCase().replace(/[\s$]/g, "").replace(/;/g, ",").replace(/[“”]/g, '"'); }

  /** Normalised answer for a part (the thing that is hashed / compared). */
  function normAnswer(part, ans) {
    switch (part.type) {
      case "num": return normNum(ans, part);
      case "list": return normList(ans);
      case "choice": return ans == null ? "" : String(ans);
      case "multi": return normMulti(ans);
      case "formula": return normFormula(ans);
      default: return normText(ans);
    }
  }
  /** Normalised KEY for a part (server side, plain keys). */
  function normKey(part) {
    switch (part.type) {
      case "num": return normNum(part.key, part);
      case "list": return normList(part.key);
      case "multi": return normMulti(part.key);
      case "formula": return (part.key || []).map(normFormula);
      default: return String(part.key);
    }
  }

  // ------------------------------------------------------------- ranges (no WBCore dependency, for Apps Script)
  function rect(ref) {
    const s = String(ref || "").replace(/\$/g, "").toUpperCase();
    const body = s.slice(s.lastIndexOf("!") + 1);
    const sheet = s.indexOf("!") >= 0 ? s.slice(0, s.lastIndexOf("!")).replace(/^'|'$/g, "") : "";
    const m = body.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) return null;
    const col = (t) => t.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const c1 = col(m[1]), r1 = +m[2], c2 = m[3] ? col(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    return { sheet, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
  }
  const within = (o, i) => !!o && !!i && o.r1 <= i.r1 && o.c1 <= i.c1 && o.r2 >= i.r2 && o.c2 >= i.c2;
  function sameRange(a, b, defSheet) {
    const x = rect(a), y = rect(b);
    if (!x || !y) return false;
    const sx = x.sheet || String(defSheet || "").toUpperCase(), sy = y.sheet || String(defSheet || "").toUpperCase();
    return sx === sy && x.r1 === y.r1 && x.r2 === y.r2 && x.c1 === y.c1 && x.c2 === y.c2;
  }

  /** Workbook "specs" made in the lab: charts, analyses (ToolPak runs) and highlight rules. */
  function specCheck(check, params, specs) {
    specs = specs || {};
    if (check === "chart") {
      return (specs.charts || []).some((c) => c.sheet === params.sheet && c.type === params.type &&
        (c.ranges || []).some((r) => within(rect(r), rect(params.covers))) &&
        (params.gapMax == null || Number(c.gap) <= params.gapMax) && (!params.trendline || !!c.trendline));
    }
    if (check === "toolpak") {
      return (specs.analyses || []).some((a) => a.tool === params.tool && sameRange(a.input, params.input, a.sheet) &&
        (!params.input2 || sameRange(a.input2, params.input2, a.sheet)) && (!params.labels || !!a.labels));
    }
    if (check === "highlight") {
      return (specs.rules || []).some((r) => r.sheet === params.sheet && within(rect(r.range), rect(params.covers)) &&
        normText(r.value) === normText(params.value));
    }
    return false;
  }

  /**
   * Score one activity submission.
   *   parts:    the activity parts (with keys on the server, or hashes on the client)
   *   sub:      {cells:{ref:{f,v}}, answers:{pid:value}, specs:{...}, text:{ref:value}}
   *   isRight:  (kind, id, norm) -> boolean   kind = "cell" | "part"
   * Returns {auto, autoMax, manualMax, parts:{pid:{frac, pts, max, detail}}}
   */
  function scoreActivity(parts, sub, isRight) {
    sub = sub || {};
    const out = { auto: 0, autoMax: 0, manualMax: 0, parts: {} };
    for (const p of parts) {
      let frac = 0, detail = null;
      if (p.manual) {
        out.manualMax += p.points;
        out.parts[p.id] = { frac: null, pts: null, max: p.points, manual: true };
        continue;
      }
      if (p.type === "cells") {
        detail = {};
        let ok = 0;
        for (const ref of p.cells) {
          const c = (sub.cells || {})[ref];
          const hasF = !!c && typeof c.f === "string" && c.f.charAt(0) === "=";
          const right = hasF && isRight("cell", ref, normCell(c.v));
          detail[ref] = right ? 1 : (c && c.f !== "" && c.f != null ? (hasF ? 0 : -1) : -2);   // 1 ok, 0 wrong value, -1 no formula, -2 empty
          if (right) ok++;
        }
        frac = p.cells.length ? ok / p.cells.length : 0;
      } else if (p.type === "spec") {
        frac = specCheck(p.check, p.params, sub.specs) ? 1 : 0;
      } else {
        const a = (sub.answers || {})[p.id];
        frac = a == null || a === "" ? 0 : (isRight("part", p.id, normAnswer(p, a)) ? 1 : 0);
      }
      const pts = Math.round(p.points * frac * 100) / 100;
      out.auto += pts;
      out.autoMax += p.points;
      out.parts[p.id] = { frac, pts, max: p.points, detail };
    }
    out.auto = Math.round(out.auto * 100) / 100;
    return out;
  }

  /** Quiz scoring on the server (plain keys). answers = {n: value}. */
  function scoreQuiz(quiz, answers) {
    let score = 0;
    const res = {};
    for (const it of quiz.items) {
      const a = (answers || {})[it.n];
      let ok = false;
      if (it.type === "choice") ok = String(a) === it.key;
      else if (it.type === "num") ok = a != null && a !== "" && normNum(a, it) === normNum(it.key, it);
      else if (it.type === "formula") ok = it.key.map(normFormula).indexOf(normFormula(a)) >= 0;
      res[it.n] = ok ? 1 : 0;
      if (ok) score++;
    }
    return { score, max: quiz.items.length, items: res };
  }

  const api = { normText, normCell, parseNum, normNum, normList, normMulti, normFormula, normAnswer, normKey,
    rect, within, sameRange, specCheck, scoreActivity, scoreQuiz };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.DAL3Check = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
