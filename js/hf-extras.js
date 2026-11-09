/* Extra Excel functions for HyperFormula, used by the DAL3 Virtual Lab.
   Adds MODE.SNGL/MODE, QUARTILE.INC/QUARTILE, PERCENTILE.INC/PERCENTILE, INTERCEPT, FORECAST.LINEAR/FORECAST,
   XLOOKUP, AVERAGEIFS. TRUE and FALSE also work without parentheses (named expressions).
   Classic script: defines installHFExtras(lib) where lib is the HyperFormula module namespace. */
(function (root) {
  function installHFExtras(lib) {
    const { HyperFormula, FunctionPlugin, FunctionArgumentType: T, CellError, ErrorType } = lib;
    if (HyperFormula.__dal3Extras) return;
    HyperFormula.__dal3Extras = true;

    const raw = (v) => (v && typeof v === "object" && "val" in v ? v.val : v);
    const isNum = (v) => typeof raw(v) === "number";
    const nums = (range) => range.valuesFromTopLeftCorner().filter(isNum).map(raw);
    const flat = (range) => range.valuesFromTopLeftCorner().map(raw);
    const err = (t) => new CellError(ErrorType[t]);

    function quantileInc(xs, p) {
      if (!xs.length || p < 0 || p > 1) return err("NUM");
      const s = [...xs].sort((a, b) => a - b);
      const h = (s.length - 1) * p;
      const lo = Math.floor(h);
      return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]);
    }
    function mode(xs) {
      const cnt = new Map();
      let best = null, bestC = 1;
      for (const x of xs) {                     // Excel: first value (in order) with the highest count > 1
        const c = (cnt.get(x) || 0) + 1;
        cnt.set(x, c);
      }
      for (const x of xs) {
        const c = cnt.get(x);
        if (c > bestC) { best = x; bestC = c; }
      }
      return best === null ? err("NA") : best;
    }
    function pairs(ys, xs) {
      const Y = flat(ys), X = flat(xs);
      if (Y.length !== X.length) return err("NA");
      const a = [], b = [];
      for (let i = 0; i < Y.length; i++) if (typeof Y[i] === "number" && typeof X[i] === "number") { a.push(Y[i]); b.push(X[i]); }
      return [a, b];
    }
    function line(ys, xs) {
      const p = pairs(ys, xs);
      if (p instanceof CellError) return p;
      const [y, x] = p, n = x.length;
      if (n < 2) return err("DIV_BY_ZERO");
      const mx = x.reduce((s, v) => s + v, 0) / n, my = y.reduce((s, v) => s + v, 0) / n;
      let sxy = 0, sxx = 0;
      for (let i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; }
      if (sxx === 0) return err("DIV_BY_ZERO");
      const slope = sxy / sxx;
      return { slope, intercept: my - slope * mx };
    }
    // Excel criteria: ">=10", "<5", "<>x", "=x", "Cash", "C*"
    function matcher(crit) {
      crit = raw(crit);
      if (typeof crit === "number") return (v) => raw(v) === crit;
      const m = String(crit).match(/^(<=|>=|<>|<|>|=)?(.*)$/);
      const op = m[1] || "=", rhs = m[2];
      const n = rhs.trim() !== "" && !isNaN(Number(rhs)) ? Number(rhs) : null;
      const re = new RegExp("^" + rhs.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
      return (v) => {
        v = raw(v);
        if (n !== null && typeof v === "number") {
          return { "=": v === n, "<>": v !== n, "<": v < n, ">": v > n, "<=": v <= n, ">=": v >= n }[op];
        }
        if (op === "=") return typeof v === "string" ? re.test(v) : rhs === "" && (v === null || v === undefined);
        if (op === "<>") return !(typeof v === "string" && re.test(v));
        return false;
      };
    }

    class DAL3Extras extends FunctionPlugin {
      modeSngl(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("MODE.SNGL"), (...rs) => mode(rs.flatMap(nums)));
      }
      quartileInc(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("QUARTILE.INC"), (r, q) => {
          q = Math.floor(q);
          return q < 0 || q > 4 ? err("NUM") : quantileInc(nums(r), q / 4);
        });
      }
      percentileInc(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("PERCENTILE.INC"), (r, p) => quantileInc(nums(r), p));
      }
      intercept(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("INTERCEPT"), (ys, xs) => {
          const l = line(ys, xs);
          return l instanceof CellError ? l : l.intercept;
        });
      }
      forecast(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("FORECAST.LINEAR"), (x, ys, xs) => {
          const l = line(ys, xs);
          return l instanceof CellError ? l : l.intercept + l.slope * x;
        });
      }
      xlookup(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("XLOOKUP"), (key, look, ret, notFound) => {
          const L = flat(look), R = flat(ret), k = raw(key);
          const i = L.findIndex((v) => (typeof v === "string" && typeof k === "string" ? v.toLowerCase() === k.toLowerCase() : v === k));
          if (i < 0) return notFound === undefined ? err("NA") : notFound;
          return R[i] ?? err("NA");
        });
      }
      averageifs(ast, state) {
        return this.runFunction(ast.args, state, this.metadata("AVERAGEIFS"), (avg, ...rest) => {
          const A = flat(avg);
          const conds = [];
          for (let i = 0; i + 1 < rest.length; i += 2) conds.push([flat(rest[i]), matcher(rest[i + 1])]);
          let s = 0, c = 0;
          A.forEach((v, i) => {
            if (typeof v === "number" && conds.every(([r, f]) => f(r[i]))) { s += v; c++; }
          });
          return c ? s / c : err("DIV_BY_ZERO");
        });
      }
    }
    const R = { argumentType: T.RANGE }, N = { argumentType: T.NUMBER };
    DAL3Extras.implementedFunctions = {
      "MODE.SNGL": { method: "modeSngl", parameters: [R], repeatLastArgs: 1 },
      "QUARTILE.INC": { method: "quartileInc", parameters: [R, N] },
      "PERCENTILE.INC": { method: "percentileInc", parameters: [R, N] },
      INTERCEPT: { method: "intercept", parameters: [R, R] },
      "FORECAST.LINEAR": { method: "forecast", parameters: [N, R, R] },
      XLOOKUP: { method: "xlookup", parameters: [{ argumentType: T.SCALAR }, R, R, { argumentType: T.SCALAR, optionalArg: true }] },
      AVERAGEIFS: { method: "averageifs", parameters: [R, R, { argumentType: T.SCALAR }], repeatLastArgs: 2 },
    };
    DAL3Extras.aliases = { MODE: "MODE.SNGL", QUARTILE: "QUARTILE.INC", PERCENTILE: "PERCENTILE.INC", FORECAST: "FORECAST.LINEAR" };
    const names = [...Object.keys(DAL3Extras.implementedFunctions), ...Object.keys(DAL3Extras.aliases)];
    const tr = Object.fromEntries(names.map((n) => [n, n]));
    HyperFormula.registerFunctionPlugin(DAL3Extras, { enGB: tr, enUS: tr });
  }

  // Named expressions so TRUE / FALSE work without parentheses, as in Excel.
  installHFExtras.namedExpressions = [{ name: "TRUE", expression: "=TRUE()" }, { name: "FALSE", expression: "=FALSE()" }];

  if (typeof module !== "undefined" && module.exports) module.exports = installHFExtras;
  else root.installHFExtras = installHFExtras;
})(typeof globalThis !== "undefined" ? globalThis : this);
