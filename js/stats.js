// Analysis ToolPak equivalents. Output layouts follow Excel's ToolPak, so trainees see the same tables.

// ---------------------------------------------------------------- special functions
function lgamma(x) {               // Lanczos
  const g = 7, c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - lgamma(1 - x);
  x -= 1;
  let a = c[0];
  const t = x + g + 0.5;
  for (let i = 1; i < g + 2; i++) a += c[i] / (x + i);
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}
function betacf(a, b, x) {
  const MAXIT = 300, EPS = 3e-16, FPMIN = 1e-300;
  let qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= MAXIT; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}
function ibeta(x, a, b) {          // regularized incomplete beta I_x(a,b)
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? bt * betacf(a, b, x) / a : 1 - bt * betacf(b, a, 1 - x) / b;
}
/** P(T > t) for t >= 0 (one tail) */
export function tTail(t, df) { return 0.5 * ibeta(df / (df + t * t), df / 2, 0.5); }
/** t such that P(T > t) = p */
export function tInv(p, df) {
  let lo = 0, hi = 1000;
  for (let i = 0; i < 200; i++) { const mid = (lo + hi) / 2; if (tTail(mid, df) > p) lo = mid; else hi = mid; }
  return (lo + hi) / 2;
}
/** P(F > f) */
export function fTail(f, d1, d2) { return ibeta(d2 / (d2 + d1 * f), d2 / 2, d1 / 2); }

// ---------------------------------------------------------------- basics
const sum = (xs) => xs.reduce((s, x) => s + x, 0);
const mean = (xs) => sum(xs) / xs.length;
function variance(xs) { const m = mean(xs); return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - 1); }
function median(xs) { const s = [...xs].sort((a, b) => a - b), n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }
function mode(xs) {
  const cnt = new Map();
  xs.forEach((x) => cnt.set(x, (cnt.get(x) || 0) + 1));
  let best = null, bc = 1;
  for (const x of xs) if (cnt.get(x) > bc) { best = x; bc = cnt.get(x); }
  return best === null ? "#N/A" : best;
}
function skew(xs) {
  const n = xs.length, m = mean(xs), s = Math.sqrt(variance(xs));
  return n / ((n - 1) * (n - 2)) * sum(xs.map((x) => ((x - m) / s) ** 3));
}
function kurt(xs) {
  const n = xs.length, m = mean(xs), s = Math.sqrt(variance(xs));
  return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * sum(xs.map((x) => ((x - m) / s) ** 4)) - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3));
}

// ---------------------------------------------------------------- ToolPak tools  (each returns a 2-D array)
/** cols = [{label, values:number[]}] */
export function descriptive(cols) {
  const names = ["Mean", "Standard Error", "Median", "Mode", "Standard Deviation", "Sample Variance", "Kurtosis", "Skewness",
    "Range", "Minimum", "Maximum", "Sum", "Count"];
  const out = [[], []];
  for (let i = 0; i < names.length; i++) out.push([]);
  cols.forEach((col) => {
    const x = col.values, n = x.length, v = variance(x);
    const vals = [mean(x), Math.sqrt(v / n), median(x), mode(x), Math.sqrt(v), v, n > 3 ? kurt(x) : "#DIV/0!",
      n > 2 ? skew(x) : "#DIV/0!", Math.max(...x) - Math.min(...x), Math.min(...x), Math.max(...x), sum(x), n];
    out[0].push(col.label, null);
    out[1].push(null, null);
    names.forEach((nm, k) => out[k + 2].push(nm, vals[k]));
  });
  return out;
}

/** Welch t-test, as "t-Test: Two-Sample Assuming Unequal Variances" (df rounded like Excel's ToolPak). */
export function tTestUnequal(a, b, la = "Variable 1", lb = "Variable 2", hyp = 0, alpha = 0.05) {
  const m1 = mean(a), m2 = mean(b), v1 = variance(a), v2 = variance(b), n1 = a.length, n2 = b.length;
  const se2 = v1 / n1 + v2 / n2;
  const dfExact = se2 ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1));
  const df = Math.round(dfExact);
  const t = (m1 - m2 - hyp) / Math.sqrt(se2);
  const p1 = tTail(Math.abs(t), df);
  return [
    ["t-Test: Two-Sample Assuming Unequal Variances", null, null], [null, null, null],
    [null, la, lb], ["Mean", m1, m2], ["Variance", v1, v2], ["Observations", n1, n2],
    ["Hypothesized Mean Difference", hyp, null], ["df", df, null], ["t Stat", t, null],
    ["P(T<=t) one-tail", p1, null], ["t Critical one-tail", tInv(alpha, df), null],
    ["P(T<=t) two-tail", 2 * p1, null], ["t Critical two-tail", tInv(alpha / 2, df), null],
  ];
}

/** Simple linear regression, as the ToolPak "SUMMARY OUTPUT". */
export function regression(y, x, ly = "Y", lx = "X Variable 1") {
  const n = y.length, mx = mean(x), my = mean(y);
  let sxx = 0, sxy = 0, syy = 0;
  for (let i = 0; i < n; i++) { sxx += (x[i] - mx) ** 2; sxy += (x[i] - mx) * (y[i] - my); syy += (y[i] - my) ** 2; }
  const b1 = sxy / sxx, b0 = my - b1 * mx;
  const ssr = b1 * sxy, sse = syy - ssr, df2 = n - 2;
  const r2 = ssr / syy, mse = sse / df2, se = Math.sqrt(mse);
  const F = ssr / mse;
  const seB1 = Math.sqrt(mse / sxx), seB0 = Math.sqrt(mse * (1 / n + mx * mx / sxx));
  const tc = tInv(0.025, df2);
  const row = (name, b, s) => { const t = b / s; return [name, b, s, t, 2 * tTail(Math.abs(t), df2), b - tc * s, b + tc * s]; };
  const pad = (r) => [...r, ...Array(7 - r.length).fill(null)];
  return [
    ["SUMMARY OUTPUT"], [], ["Regression Statistics"], ["Multiple R", Math.sqrt(r2)], ["R Square", r2],
    ["Adjusted R Square", 1 - (1 - r2) * (n - 1) / df2], ["Standard Error", se], ["Observations", n], [],
    ["ANOVA"], [null, "df", "SS", "MS", "F", "Significance F"],
    ["Regression", 1, ssr, ssr, F, fTail(F, 1, df2)], ["Residual", df2, sse, mse], ["Total", n - 1, syy], [],
    [null, "Coefficients", "Standard Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"],
    row("Intercept", b0, seB0), row(lx, b1, seB1),
  ].map(pad);
}

/** Least-squares line for chart trendlines. */
export function fitLine(x, y) {
  const mx = mean(x), my = mean(y);
  let sxx = 0, sxy = 0, syy = 0;
  for (let i = 0; i < x.length; i++) { sxx += (x[i] - mx) ** 2; sxy += (x[i] - mx) * (y[i] - my); syy += (y[i] - my) ** 2; }
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, r2: (sxy * sxy) / (sxx * syy) };
}
