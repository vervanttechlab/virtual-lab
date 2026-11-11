// DAL3 Virtual Lab: router and views (student + trainer).
import { CONFIG } from "./config.js";
import { call, trainerCall, sendOrQueue, session, trainer, demoMode, queue, onQueue, onBusy, apiUrl, NetError } from "./api.js";
import { Workbook } from "./workbook.js";
import { SheetUI } from "./grid.js";
import { buildSubmissionFile, downloadBlob } from "./export.js";
import { VIDEOS, embedUrl, driveId } from "./videos.js";
import { el, esc, modal, toast, store, sha, timeAgo, icon } from "./ui.js";
import { loadSql, runSql, fingerprintOf, sqlEditor, schemaPanel, insertAtCursor, resultTable, downloadCsv } from "./sqllab.js";

const C = window.DAL3Check;
const app = document.getElementById("app");
const cache = { days: {} };
let cleanup = [];
const onLeave = (fn) => cleanup.push(fn);

// ------------------------------------------------------------------ data
async function getIndex() {
  if (!cache.index) cache.index = await (await fetch("data/index.json", { cache: "no-cache" })).json();
  return cache.index;
}
async function getDay(n) {
  const tag = `day${String(n).padStart(2, "0")}`;
  if (!cache.days[tag]) cache.days[tag] = await (await fetch(`data/${tag}.json`, { cache: "no-cache" })).json();
  return cache.days[tag];
}
const uid = () => (session.get() || {}).id || "guest";

/** Load the spreadsheet engine + charts only when a workbook is opened (keeps sign-in and lists fast). */
const scripts = {};
function loadScript(src) {
  if (!scripts[src]) {
    scripts[src] = new Promise((ok, bad) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = ok;
      s.onerror = () => { delete scripts[src]; bad(new Error("Could not load the workbook. Check your connection and try again.")); };
      document.head.append(s);
    });
  }
  return scripts[src];
}
async function loadEngine() {
  await Promise.all([loadScript("vendor/hyperformula.full.min.js"), loadScript("vendor/chart.umd.js")]);
  if (window.installHFExtras && !window.HyperFormula.HyperFormula?.__dal3Extras) window.installHFExtras(window.HyperFormula);
}
const K = {
  wb: (tag) => `dal3_wb_${uid()}_${tag}`,
  ans: (tag) => `dal3_ans_${uid()}_${tag}`,
  res: (tag) => `dal3_res_${uid()}_${tag}`,
  quiz: (tag) => `dal3_quiz_${uid()}_${tag}`,
};

/** Student config: which activities are open + my results. Cached for offline use. */
async function loadConfig(force) {
  if (demoMode()) return { enabled: null, results: {} };
  const ck = `dal3_cfg_${uid()}`;
  const cached = store.get(ck);
  if (cached && !force && Date.now() - cached.at < 15000) return cached;
  try {
    const r = await call("config");
    const v = { enabled: r.enabled, unlocked: r.unlocked || {}, results: r.results || {}, links: r.links || {}, redo: r.redo || {}, at: Date.now() };
    store.set(ck, v);
    return v;
  } catch (e) {
    if (e instanceof NetError && cached) return cached;
    if (e instanceof NetError) return { enabled: null, results: {}, offline: true };
    if (/token|login/i.test(e.message)) { session.clear(); location.hash = "#/login"; }
    throw e;
  }
}
const isOpen = (cfg, tag, item) => !cfg.enabled || !!(cfg.enabled[tag] && cfg.enabled[tag][item]);
/** Activities the trainer returned to this student for a redo: [{tag, act, at, note}], oldest first. */
const redoList = (cfg) => Object.entries((cfg && cfg.redo) || {}).flatMap(([tag, acts]) => Object.entries(acts).map(([act, r]) => ({ tag, act, ...r })))
  .sort((a, b) => new Date(a.at) - new Date(b.at));
const redoOf = (cfg, tag, act) => (((cfg && cfg.redo) || {})[tag] || {})[act] || null;
/** Grades / feedback the student has not looked at yet (compared with what My progress last showed). */
const seenKey = () => `dal3_seengrades_${uid()}`;
const gradeSig = (tag, act, r) => `${tag}|${act}|${r.attempts}|${r.manual}|${r.feedback || ""}`;
function newGrades(cfg) {
  const seen = new Set(store.get(seenKey(), [])), out = [];
  for (const [tag, acts] of Object.entries((cfg && cfg.results) || {}))
    for (const [act, r] of Object.entries(acts)) {
      const graded = (r.manual !== "" && r.manual != null) || !!r.feedback;
      if (graded && !seen.has(gradeSig(tag, act, r))) out.push({ tag, act, r });
    }
  return out;
}
function markGradesSeen() {
  const cfg = store.get(`dal3_cfg_${uid()}`);
  if (!cfg) return;
  const all = [];
  for (const [tag, acts] of Object.entries(cfg.results || {})) for (const [act, r] of Object.entries(acts)) all.push(gradeSig(tag, act, r));
  store.set(seenKey(), all);
}
/** Tell the server the first time this student opens an activity or quiz (time on task). */
function noteOpened(tag, act) {
  const k = `dal3_opened_${uid()}_${tag}_${act}`;
  if (demoMode() || store.get(k)) return;
  call("opened", { day: tag, act }, { quiet: true }).then(() => store.set(k, 1)).catch(() => {});
}
/** Classmates for the partner picker (kept for a day on the phone). */
async function loadClassmates() {
  const k = `dal3_classmates_${uid()}`, c = store.get(k);
  if (c && Date.now() - c.at < 86400000) return c.list;
  const r = await call("classmates", {}, { quiet: true });
  store.set(k, { list: r.list || [], at: Date.now() });
  return r.list || [];
}

// ------------------------------------------------------------------ shell
function shell(title, back, ...body) {
  let cls = "page narrow";
  if (body.length && body[body.length - 1] && body[body.length - 1].__cls) cls = body.pop().__cls;
  const s = session.get();
  const qn = queue.list().length;
  const pend = el("span", { class: "tb-flag unsent" + (qn ? " show" : ""), title: "Work saved on this device, waiting for a connection" }, `${qn} UNSENT`);
  onLeave(onQueue((n) => { pend.textContent = `${n} UNSENT`; pend.classList.toggle("show", n > 0); }));
  const parts = String(title).split(" › ");
  const crumbHref = (t) => (t === "Home" ? "#/" : /^Day (\d+)$/.test(t) ? `#/day/${t.slice(4)}` : null);
  const trail = parts[0] === "Home" || parts[0] === "Sign in" || parts[0] === "Trainer board" ? parts : ["Home", ...parts];
  const crumbs = el("nav", { class: "crumbs", "aria-label": "You are here" }, ...trail.flatMap((t, i) => {
    const last = i === trail.length - 1, href = !last && crumbHref(t);
    return [i ? el("span", { class: "sep" }, "/") : "", href ? el("a", { href }, t) : el("span", { class: last ? "cur" : "" }, t)];
  }));
  const backLabel = back === "#/" ? "Home" : (back || "").replace(/^#\/day\/(\d+).*$/, "Day $1").replace(/^#\/lesson\/(\d+).*$/, "Lesson");
  const head = el("header", { class: "titlebar" },
    back ? el("a", { class: "tb-btn tb-back", href: back, title: `Back to ${backLabel}` }, icon("back"), el("span", {}, backLabel)) : "",
    el("a", { class: "tb-btn brand", href: "#/" }, "DAL3 LAB"),
    crumbs, pend,
    demoMode() ? el("span", { class: "tb-flag", title: "No class server configured: work stays on this device" }, "DEMO") : "",
    s ? el("button", { class: "tb-btn tb-user", type: "button", onclick: () => accountMenu() }, s.id) : "");
  const main = el("main", { class: cls }, ...body);
  // Students get the sidebar (PC/tablet) and bottom tabs (phone); sign-in and trainer pages keep the plain layout.
  const student = !!s && !/^(Sign in|Trainer)/.test(String(title)) && !location.hash.startsWith("#/trainer");
  const work = /\bwork\b/.test(cls);                 // activity / workbook / SQL: the sidebar shrinks to icons, no bottom tabs
  document.body.classList.toggle("has-snav", student);
  document.body.classList.toggle("has-bnav", student && !work);
  if (!student) app.replaceChildren(head, main);
  else {
    const rail = work || !!store.get("dal3_rail", false);
    const body_ = el("div", { class: "appbody" + (rail ? " rail" : "") }, sideNav(rail, work), main);
    document.body.classList.toggle("snav-rail", rail);
    app.replaceChildren(head, body_, work ? "" : bottomNav());
    if (!cache.index) getIndex().then(() => { const a = body_.querySelector(".sidenav"); if (a && a.isConnected) a.replaceWith(sideNav(rail, work)); }).catch(() => {});
  }
  window.scrollTo(0, 0);
}

// ------------------------------------------------------------------ student navigation: sidebar (PC/tablet) + bottom tabs (phone)
const hashDay = () => { const m = location.hash.match(/^#\/(?:day|lesson|sheet|sql|act|quiz|files)\/(\d+)/); return m ? Number(m[1]) : null; };
const navCfg = () => store.get(`dal3_cfg_${uid()}`) || null;
/** The day to send a student to: the one they are on, else the latest day with something open, else Day 1 of the list. */
function currentDay(cfg) {
  const days = (cache.index && cache.index.days) || [];
  const withOpen = cfg && cfg.enabled ? days.filter((d) => [...d.activities.map((a) => a.id), "quiz"].some((id) => isOpen(cfg, d.tag, id))) : [];
  return hashDay() || (withOpen.pop() || days[days.length - 1] || {}).day || null;
}
function sideNav(rail, work) {
  const cfg = navCfg(), s = session.get() || {}, h = location.hash || "#/", here = hashDay();
  const ng = cfg && !demoMode() ? newGrades(cfg).length : 0;
  const link = (href, ic, label, on, extra) => el("a", { class: "sn-link" + (on ? " on" : ""), href, title: label, "aria-current": on ? "page" : null },
    icon(ic), el("span", { class: "sn-t" }, label), extra || "");
  const days = (cache.index && cache.index.days) || [];
  const dayRows = days.map((d) => {
    const ids = [...d.activities.map((a) => a.id), "quiz"], total = ids.length;
    const res = (cfg && cfg.results && cfg.results[d.tag]) || {};
    const done = ids.filter((id) => res[id]).length;
    const openN = cfg && cfg.enabled ? ids.filter((id) => isOpen(cfg, d.tag, id)).length : total;
    const locked = !!(cfg && cfg.enabled) && !openN && !done;
    const on = here === d.day;
    return el("a", { class: "sn-day" + (on ? " on" : "") + (locked ? " locked" : "") + (openN && !locked ? " live" : ""), href: `#/day/${d.day}`,
      title: `Day ${d.day}: ${d.title}` + (locked ? " (not open yet)" : ` · ${done} of ${total} done`), "aria-current": on ? "page" : null },
      el("span", { class: "sn-dn mono" }, String(d.day).padStart(2, "0")),
      el("span", { class: "sn-t sn-dbody" }, el("span", { class: "sn-dt" }, d.title),
        el("span", { class: "sn-bar" }, el("span", { style: `width:${Math.round(100 * done / total)}%` }))),
      el("span", { class: "sn-t sn-c" }, locked ? icon("padlock") : `${done}/${total}`));
  });
  const initials = (s.name || s.id || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join("").toUpperCase();
  const toggle = work ? "" : el("button", { class: "sn-link sn-collapse", type: "button", title: rail ? "Show the full menu" : "Make the menu narrow",
    onclick: () => { store.set("dal3_rail", !rail); route(); } }, icon(rail ? "chev" : "back"), el("span", { class: "sn-t" }, "Narrow menu"));
  return el("aside", { class: "sidenav", "aria-label": "Lab menu" },
    el("button", { class: "sn-user", type: "button", title: "Account and display settings", onclick: () => accountMenu() },
      el("span", { class: "sn-av" }, initials), el("span", { class: "sn-t" }, el("b", {}, s.name || s.id || ""), el("span", { class: "muted" }, s.id && s.name ? s.id : ""))),
    el("nav", { class: "sn-main" },
      link("#/", "home", "Home", h === "#/" || h === "#" || h === ""),
      demoMode() ? "" : link("#/progress", "chart", "My progress", h.startsWith("#/progress"), ng ? el("span", { class: "sn-dot", title: `${ng} new grade(s)` }, String(ng)) : "")),
    days.length ? el("div", { class: "sn-label sn-t" }, "Training days") : "",
    el("nav", { class: "sn-days", "aria-label": "Training days" }, ...dayRows),
    el("div", { class: "sn-foot" },
      el("button", { class: "sn-link", type: "button", title: "Account, text size, theme, sign out", onclick: () => accountMenu() }, icon("gear"), el("span", { class: "sn-t" }, "Settings")),
      toggle));
}
function bottomNav() {
  const cfg = navCfg(), h = location.hash || "#/", n = currentDay(cfg);
  const ng = cfg && !demoMode() ? newGrades(cfg).length : 0;
  const tab = (href, ic, label, on, dot) => el(href ? "a" : "button", href ? { class: "bn" + (on ? " on" : ""), href, "aria-current": on ? "page" : null }
    : { class: "bn", type: "button", onclick: () => accountMenu() }, el("span", { class: "bn-ic" }, icon(ic), dot ? el("span", { class: "bn-dot" }, String(dot)) : ""), el("span", {}, label));
  return el("nav", { class: "bnav", "aria-label": "Lab menu" },
    tab("#/", "home", "Home", h === "#/" || h === "#" || h === ""),
    n ? tab(`#/day/${n}`, "list", `Day ${n}`, !!hashDay()) : "",
    demoMode() ? "" : tab("#/progress", "chart", "Progress", h.startsWith("#/progress"), ng),
    tab(null, "user", "Account"));
}
/** Tabs shown on every page of a day, so students can move between its parts in one tap. */
function dayTabs(n, active) {
  const d = cache.days[`day${String(n).padStart(2, "0")}`];
  const hasSheets = !d || (d.sheets || []).length > 0, hasSql = !!(d && d.sql);
  const tabs = [["day", "Activities", `#/day/${n}`, "list"], ["lesson", "Lesson", `#/lesson/${n}`, "book"],
    hasSql ? ["sql", "SQL", `#/sql/${n}`, "db"] : null, hasSheets ? ["sheet", "Workbook", `#/sheet/${n}`, "grid"] : null,
    ["quiz", "Quiz", `#/quiz/${n}`, "quiz"], ["files", "Files", `#/files/${n}`, "download"]].filter(Boolean);
  return el("nav", { class: "daytabs", "aria-label": `Day ${n}` }, el("span", { class: "dt-day" }, `Day ${n}`),
    ...tabs.map(([k, label, href, ic]) => el("a", { class: "dt" + (k === active ? " on" : ""), href, "aria-current": k === active ? "page" : null },
      icon(ic), el("span", {}, label))));
}

// ------------------------------------------------------------------ display settings (per device): text size + theme
const display = () => ({ size: "normal", theme: "auto", ...store.get("dal3_display", {}) });
function applyDisplay() {
  const d = display(), root = document.documentElement;
  if (d.size === "normal") delete root.dataset.size; else root.dataset.size = d.size;
  if (d.theme === "auto") delete root.dataset.theme; else root.dataset.theme = d.theme;
}
applyDisplay();
function segmented(options, current, onPick) {
  const g = el("div", { class: "seg3", role: "radiogroup" });
  for (const [val, label, cls] of options) {
    const b = el("button", { class: (val === current ? "on " : "") + (cls || ""), type: "button", role: "radio", "aria-checked": String(val === current) }, label);
    b.addEventListener("click", () => {
      g.querySelectorAll("button").forEach((x) => { x.classList.toggle("on", x === b); x.setAttribute("aria-checked", String(x === b)); });
      onPick(val);
    });
    g.append(b);
  }
  return g;
}

function accountMenu() {
  const s = session.get();
  const set = (k, v) => { store.set("dal3_display", { ...display(), [k]: v }); applyDisplay(); };
  const box = el("div", { class: "form" },
    el("p", {}, el("b", {}, s.name || s.id), el("span", { class: "muted" }, `  ${s.id}${s.section ? " · " + s.section : ""}`)),
    el("div", { class: "label", style: "padding-top:6px" }, "Display"),
    el("div", { class: "small muted" }, "Text size"),
    segmented([["normal", "Aa Normal", "aa1"], ["large", "Aa Large", "aa2"], ["larger", "Aa Larger", "aa3"]], display().size, (v) => set("size", v)),
    el("div", { class: "small muted" }, "Theme"),
    segmented([["auto", "Auto"], ["light", "Light"], ["dark", "Dark"]], display().theme, (v) => set("theme", v)),
    el("div", { class: "actions" },
    demoMode() ? "" : el("a", { class: "btn wide", href: "#/progress", onclick: () => m.close() }, icon("chart"), "My progress"),
    el("button", { class: "btn wide danger", type: "button", onclick: () => {
      if (queue.list().length && !confirm("Some work is still waiting to be sent. Sign out anyway? It will be sent after you sign in again on this phone.")) return;
      session.clear(); m.close(); location.hash = "#/login";
    } }, "Sign out")));
  const m = modal("Account", box);
}

// ------------------------------------------------------------------ views
async function viewLogin() {
  const f = el("form", { class: "form signin" });
  f.innerHTML = `
    <h2>Sign in</h2>
    <p class="sub">DAL3 Virtual Lab · TESDA Data Analytics Level III</p>
    ${demoMode() ? `<p class="note">Demo mode: no class server is set up yet. Your work is saved on this device only.</p>
      <label>Your name <input name="name" required autocomplete="name"></label>` : `
      <label>Class code <input name="cls" required value="${esc(CONFIG.CLASS_CODE_HINT)}" autocapitalize="characters"></label>
      <label>Student ID <input name="sid" required autocapitalize="characters" autocomplete="username"></label>
      <label>PIN <input name="pin" required inputmode="numeric" type="password" autocomplete="current-password" maxlength="8"></label>`}
    <button class="btn primary wide" type="submit">Sign in</button>
    <p class="small" style="margin:12px 0 0"><a href="guide.html">New here? How to install the app and sign in</a></p>`;
  shell("Sign in", null, f);
  if (!demoMode()) call("ping", {}, { quiet: true, auth: false }).catch(() => {});      // wake the Apps Script server while they type
  loadScript("vendor/hyperformula.full.min.js").catch(() => {});                       // and fetch the engine in the background
  f.addEventListener("submit", async (e) => {
    e.preventDefault();
    const btn = f.querySelector("button[type=submit]");
    if (demoMode()) {
      const name = f.name.value.trim();
      session.set({ id: "demo-" + name.toLowerCase().replace(/\W+/g, "-"), name, demo: true });
      location.hash = "#/";
      return;
    }
    btn.disabled = true;
    btn.textContent = "Signing in…";
    try {
      const r = await call("login", { cls: f.cls.value.trim(), sid: f.sid.value.trim(), pin: f.pin.value.trim() }, { auth: false });
      session.set({ id: r.id, name: r.name, token: r.token, section: r.section });
      if (r.config) store.set(`dal3_cfg_${r.id}`, { ...r.config, at: Date.now(), fromLogin: true });   // Home opens without another round trip
      queue.flush();
      location.hash = "#/";
    } catch (err) {
      toast(err.message, "bad");
      btn.disabled = false;
      btn.textContent = "Sign in";
    }
  });
}

/** Render with the cached config at once (instant page), then refresh from the server and re-render only if something changed. */
async function withConfig(render) {
  const cached = demoMode() ? null : store.get(`dal3_cfg_${uid()}`);
  if (!cached) return render(await loadConfig(true));
  const here = location.hash;
  await render(cached);
  if (cached.fromLogin && Date.now() - (cached.at || 0) < 10000) {   // the login reply just brought it: skip ONE refresh only
    delete cached.fromLogin;
    store.set(`dal3_cfg_${uid()}`, cached);
    return;
  }
  loadConfig(true).then(async (fresh) => {
    const sig = (c) => JSON.stringify([c.enabled, c.results, c.links, c.redo]);
    if (location.hash !== here || sig(fresh) === sig(cached)) return;
    const y = window.scrollY;
    await render(fresh);
    window.scrollTo(0, y);
  }).catch(() => {});
}

async function viewHome() {
  const idx = await getIndex();
  return withConfig(async (cfg) => {
    const rows = idx.days.map((d) => {
      const open = d.activities.filter((a) => isOpen(cfg, d.tag, a.id)).length;
      const res = cfg.results[d.tag] || {};
      const done = Object.keys(res).filter((k) => k !== "quiz").length;
      return el("a", { class: "row daysrow", href: `#/day/${d.day}` },
        el("div", { class: "num" }, `DAY ${String(d.day).padStart(2, "0")}`),
        el("div", {}, el("div", { class: "title" }, d.title), el("div", { class: "sub" }, d.unit)),
        el("div", { class: "status c-count" }, cfg.enabled ? [el("span", { class: "tag" + (open ? " good" : "") }, `${open}/${d.activities.length} OPEN`),
          done ? el("span", { class: "tag info" }, `${done} SUBMITTED`) : ""] : el("span", { class: "tag" }, `${d.activities.length} ACTIVITIES`)),
        el("div", { class: "chev" }, icon("chev")));
    });
    shell("Home", null,
      el("div", { class: "pagehead" }, el("div", { class: "kicker" }, "Data Analytics Level III"),
        el("h2", {}, `Welcome, ${(session.get().name || "").split(" ")[0] || "trainee"}`),
        el("div", { class: "meta" }, cfg.offline ? "Offline. Keep working: answers are sent when the connection returns." :
          "Activities appear as your trainer opens them.")),
      (() => {
        const ng = demoMode() ? [] : newGrades(cfg);
        return ng.length ? el("a", { class: "newgrades", href: "#/progress" }, el("span", { class: "dot" }),
          el("span", {}, el("b", {}, `${ng.length} new ${ng.length === 1 ? "grade" : "grades"} or feedback`), " from your trainer"), el("span", { class: "go" }, "View", icon("chev"))) : "";
      })(),
      (() => {
        const rl = redoList(cfg);
        if (!rl.length) return "";
        const first = rl[0], dn = Number(first.tag.replace(/\D/g, ""));
        return el("a", { class: "redobanner", href: first.act === "quiz" ? `#/quiz/${dn}` : `#/act/${dn}/${first.act}` }, icon("undo"),
          el("span", {}, el("b", {}, rl.length === 1 ? `Your trainer asked you to redo Activity ${first.act}` : `Your trainer asked you to redo ${rl.length} activities`),
            first.note ? el("span", { class: "nt" }, `“${first.note}”`) : ""), el("span", { class: "go" }, "Open", icon("chev")));
      })(),
      el("div", { class: "label" }, "Training days"),
      el("div", { class: "list" }, ...rows),
      el("p", { class: "small muted", style: "margin-top:10px" }, "Days 4 to 15 are added as the course progresses."));
  });
}

// ------------------------------------------------------------------ my progress (the student's own record: next step, trainer feedback, days, log)
async function viewProgress() {
  const idx = await getIndex();
  const days = await Promise.all(idx.days.map((d) => getDay(d.day)));
  const ck = `dal3_prog_${uid()}`;
  const stamp = el("span", {});
  const refreshBtn = el("button", { class: "btn small", type: "button" }, "Refresh");
  const body = el("div", { class: "prog" });
  shell("My progress", "#/",
    el("div", { class: "pagehead" }, el("div", { class: "kicker" }, "Data Analytics Level III"), el("h2", {}, "My progress"),
      el("div", { class: "meta" }, stamp, " ", refreshBtn)), body);
  if (demoMode()) {
    body.replaceChildren(el("p", { class: "note" }, "Demo mode: there is no class server, so there is no progress record yet."));
    return;
  }
  const here = location.hash;
  // grades and feedback this student has not seen yet: marked NEW on this visit (then they count as seen)
  let fresh = new Set(newGrades(store.get(`dal3_cfg_${uid()}`)).map((x) => `${x.tag}|${x.act}`));
  const openDays = store.get(`dal3_progdays_${uid()}`, null);           // days the student opened or closed by hand
  let shown = store.get(ck);
  if (shown) draw(shown);
  else body.replaceChildren(el("p", { class: "muted" }, "Loading your progress…"));
  refreshBtn.addEventListener("click", () => load(false));
  load(!!shown);

  async function load(quiet) {
    refreshBtn.disabled = true;
    try {
      const cfgP = loadConfig(true).catch(() => null);
      let p;
      try {
        const r = await call("myProgress", {}, { quiet });
        p = { subs: r.subs || [], live: r.live || [], enabled: r.enabled || null, at: Date.now() };
      } catch (e) {
        if (!/unknown action/i.test(e.message)) throw e;
        // The Apps Script server has not been updated yet: show the latest result of each activity only.
        const cfg = (await cfgP) || { results: {} };
        const subs = [];
        for (const [day, acts] of Object.entries(cfg.results || {}))
          for (const [act, r] of Object.entries(acts)) subs.push({ ...r, day, act, attempt: r.attempts, late: false });
        p = { subs, live: [], enabled: cfg.enabled, at: Date.now(), partial: true };
      }
      const cfg = await cfgP;
      if (cfg) for (const x of newGrades(cfg)) fresh.add(`${x.tag}|${x.act}`);
      store.set(ck, p);
      shown = p;
      if (location.hash === here) { draw(p); if (cfg) markGradesSeen(); }
    } catch (e) {
      if (location.hash !== here) return;
      if (/token|login/i.test(e.message) && !(e instanceof NetError)) { session.clear(); location.hash = "#/login"; return; }
      toast(e instanceof NetError ? "Offline: showing your progress from the last time you were online." : e.message, "bad");
      if (!shown) body.replaceChildren(el("p", { class: "note" }, "Your progress cannot be loaded right now. Connect to the internet and tap Refresh."));
    } finally {
      refreshBtn.disabled = false;
    }
  }

  function draw(p) {
    stamp.textContent = `Updated ${timeAgo(p.at)}` + (p.partial ? " · latest attempts only" : "");
    const t = (x) => new Date(x).getTime() || 0;
    const subs = [...p.subs].sort((a, b) => t(a.ts) - t(b.ts));
    const latest = {}, count = {};
    for (const s of subs) { const k = `${s.day}|${s.act}`; latest[k] = s; count[k] = (count[k] || 0) + 1; }
    const working = {};
    for (const l of p.live) {
      const k = `${l.day}|${l.act}`;
      if (l.status === "draft" && (!latest[k] || t(l.updated) > t(latest[k].ts))) working[k] = l;
    }
    const open = (tag, id) => !p.enabled || !!(p.enabled[tag] && p.enabled[tag][id]);
    const pointsOf = (s) => {
      const graded = !Number(s.manualMax) || (s.manual !== "" && s.manual != null);
      return { pts: Number(s.auto || 0) + Number(s.manual || 0), max: Number(s.autoMax || 0) + Number(s.manualMax || 0), graded };
    };
    const band = (f) => (f >= 0.8 ? "good" : f >= 0.5 ? "half" : "bad");
    const href = (n, id) => (id === "quiz" ? `#/quiz/${n}` : `#/act/${n}/${id}`);
    const itemsOf = (day) => [...day.activities.map((a) => ({ id: a.id, title: a.title })), { id: "quiz", title: day.quiz.title || `Day ${day.day} quiz` }];
    const nameOf = (day, id) => (id === "quiz" ? `Day ${day.day} quiz` : `${id} ${(day.activities.find((a) => a.id === id) || {}).title || ""}`);

    // The status of every item, in plain words (no legend needed).
    const rcfg = navCfg();
    const statusOf = (day, it) => {
      const k = `${day.tag}|${it.id}`, s = latest[k], l = working[k], isOpen = open(day.tag, it.id);
      const rd = redoOf(rcfg, day.tag, it.id);
      if (rd && isOpen) return { kind: "redo", s, q: s ? pointsOf(s) : null, text: "Redo", note: "Your trainer asked for a redo" + (rd.note ? `: ${rd.note}` : "") };
      if (s) {
        const q = pointsOf(s), f = q.pts / Math.max(1, q.max);
        if (!q.graded) return { kind: "wait", s, q, text: Number(s.autoMax) ? `${round(s.auto)}/${round(s.autoMax)} so far` : "Sent", note: "Waiting for your trainer" };
        return { kind: band(f), s, q, text: `${round(q.pts)}/${round(q.max)}`, note: `${Math.round(100 * f)}%` };
      }
      if (l) return { kind: "draft", l, text: `${l.pct}%`, note: "In progress" };
      if (isOpen) return { kind: "open", text: "Start", note: "Open now" };
      return { kind: "closed", text: "", note: "Not open yet" };
    };

    // ---- 1. summary: activities done, score on graded work, waiting, last active
    let gotPts = 0, gradedMax = 0, done = 0, waiting = 0, last = 0;
    for (const s of Object.values(latest)) {
      const q = pointsOf(s);
      if (q.graded) { gotPts += q.pts; gradedMax += q.max; } else waiting++;
      done++;
    }
    for (const s of subs) last = Math.max(last, t(s.ts));
    for (const l of p.live) last = Math.max(last, t(l.updated));
    const totalItems = days.reduce((n, d) => n + d.activities.length + 1, 0);
    const pct = gradedMax ? Math.round(100 * gotPts / gradedMax) : null;
    const summary = el("div", { class: "psum" },
      el("div", { class: "psum-main" },
        el("div", { class: "k" }, "Activities and quizzes done"),
        el("div", { class: "v" }, el("b", {}, String(done)), ` of ${totalItems}`),
        el("div", { class: "pbar", role: "progressbar", "aria-valuemin": 0, "aria-valuemax": totalItems, "aria-valuenow": done },
          el("span", { style: `width:${Math.round(100 * done / Math.max(1, totalItems))}%` }))),
      el("div", { class: "psum-stats" },
        el("div", {}, el("div", { class: "k" }, "Score"), el("div", { class: "v " + (pct == null ? "" : band(pct / 100)) }, pct == null ? "—" : `${pct}%`),
          el("div", { class: "s" }, gradedMax ? `${round(gotPts)} of ${round(gradedMax)} graded points` : "No graded work yet")),
        el("div", {}, el("div", { class: "k" }, "Waiting"), el("div", { class: "v" + (waiting ? " wait" : "") }, String(waiting)),
          el("div", { class: "s" }, waiting ? "for your trainer to grade" : "nothing to grade")),
        el("div", {}, el("div", { class: "k" }, "Last active"), el("div", { class: "v sm" }, last ? timeAgo(last) : "—"))));

    // ---- 2. next step: continue work in progress, else the first open item not done yet
    // next step, in this order: a redo the trainer asked for, then work in progress, then the first open item not started
    let next = null;
    const rank = { redo: 3, draft: 2, open: 1 };
    for (const day of days) for (const it of itemsOf(day)) {
      if (!open(day.tag, it.id)) continue;
      const st = statusOf(day, it), r = rank[st.kind] || 0;
      if (r && (!next || r > rank[next.st.kind])) next = { day, it, st };
    }
    const nextCard = next
      ? el("a", { class: "pnext", href: href(next.day.day, next.it.id) },
        el("div", { class: "k" }, next.st.kind === "redo" ? "Your trainer asked for a redo" : next.st.kind === "draft" ? "Continue where you left off" : "Up next"),
        el("div", { class: "t" }, nameOf(next.day, next.it.id)),
        el("div", { class: "s" }, `Day ${next.day.day}` + (next.st.kind === "redo" ? ` · ${(redoOf(rcfg, next.day.tag, next.it.id) || {}).note || "your answers are still there"}` :
          next.st.kind === "draft" ? ` · ${next.st.l.pct}% correct so far` : " · open now")),
        el("span", { class: "go" }, next.st.kind === "redo" ? "Redo it" : next.st.kind === "draft" ? "Continue" : "Start", icon("chev")))
      : el("div", { class: "pnext done" }, el("div", { class: "k" }, "All caught up"),
        el("div", { class: "t" }, "You have done everything that is open."), el("div", { class: "s" }, "New activities appear here when your trainer opens them."));

    // ---- 3. from your trainer: grades and feedback, newest first (NEW = not seen before)
    const notes = Object.values(latest).filter((s) => s.feedback || (Number(s.manualMax) && s.manual !== "" && s.manual != null))
      .map((s) => ({ s, at: t(s.gradedAt || s.ts), day: days.find((d) => d.tag === s.day) })).filter((x) => x.day).sort((a, b) => b.at - a.at);
    const noteRow = (x) => {
      const q = pointsOf(x.s), isNew = fresh.has(`${x.s.day}|${x.s.act}`);
      return el("a", { class: "pnote" + (isNew ? " new" : ""), href: href(x.day.day, x.s.act) },
        el("div", { class: "top" }, isNew ? el("span", { class: "tag solid" }, "NEW") : "",
          el("span", { class: "t" }, nameOf(x.day, x.s.act)), el("span", { class: "pts " + band(q.pts / Math.max(1, q.max)) }, `${round(q.pts)}/${round(q.max)}`)),
        x.s.feedback ? el("div", { class: "fb" }, x.s.feedback) : el("div", { class: "fb muted" }, "Graded, no comment."),
        el("div", { class: "when" }, x.s.gradedAt ? `Graded ${timeAgo(x.s.gradedAt)}` : ""));
    };
    const noteBox = el("div", { class: "pnotes" });
    const drawNotes = (all) => noteBox.replaceChildren(...notes.slice(0, all ? notes.length : 3).map(noteRow),
      notes.length > 3 && !all ? el("button", { class: "btn small pmore", type: "button", onclick: () => drawNotes(true) }, `Show all ${notes.length}`) : "");
    drawNotes(false);

    // ---- 4. one card per day: progress bar + every activity in plain words. Tap a row to open it.
    const relevant = days.filter((day) => itemsOf(day).some((it) => open(day.tag, it.id) || latest[`${day.tag}|${it.id}`] || working[`${day.tag}|${it.id}`]));
    const current = relevant.filter((d) => itemsOf(d).some((it) => open(d.tag, it.id))).pop() || relevant[relevant.length - 1];
    const dayCards = relevant.map((day) => {
      const items = itemsOf(day), sts = items.map((it) => statusOf(day, it));
      const dDone = sts.filter((x) => x.s && x.kind !== "redo").length;
      let dPts = 0, dMax = 0;
      for (const x of sts) if (x.s && x.q && x.q.graded) { dPts += x.q.pts; dMax += x.q.max; }
      const rows = items.map((it, i) => {
        const st = sts[i], isOpen = open(day.tag, it.id), k = `${day.tag}|${it.id}`;
        const tags = [st.s && st.s.late ? el("span", { class: "tag warn" }, "LATE") : "", st.s && count[k] > 1 ? el("span", { class: "tag" }, `${count[k]} tries`) : ""];
        const inner = [el("span", { class: "id mono" }, it.id === "quiz" ? "Quiz" : it.id),
          el("span", { class: "nm" }, el("span", { class: "t" }, it.id === "quiz" ? (it.title || "Quiz") : it.title), el("span", { class: "n" }, st.note, ...tags)),
          el("span", { class: "st " + st.kind }, st.kind === "closed" ? icon("padlock") : st.text, st.kind === "open" ? icon("chev") : "")];
        return isOpen ? el("a", { class: "pitem " + st.kind, href: href(day.day, it.id) }, ...inner) : el("div", { class: "pitem " + st.kind }, ...inner);
      });
      const isOpenCard = openDays && day.tag in openDays ? openDays[day.tag] : day === current;
      const card = el("details", { class: "pdaycard", open: isOpenCard || null },
        el("summary", {},
          el("span", { class: "dnum mono" }, `DAY ${String(day.day).padStart(2, "0")}`),
          el("span", { class: "dtx" }, el("span", { class: "t" }, day.title),
            el("span", { class: "pbar sm" }, el("span", { style: `width:${Math.round(100 * dDone / items.length)}%` }))),
          el("span", { class: "dmeta" }, el("b", {}, `${dDone}/${items.length}`), " done", dMax ? el("span", { class: "s" }, `${Math.round(100 * dPts / dMax)}% score`) : ""),
          el("span", { class: "chev" }, icon("chev"))),
        el("div", { class: "pitems" }, ...rows));
      card.addEventListener("toggle", () => {
        const m = store.get(`dal3_progdays_${uid()}`, {}) || {};
        m[day.tag] = card.open;
        store.set(`dal3_progdays_${uid()}`, m);
      });
      return card;
    });

    // ---- 5. activity log (newest first, grouped by date) with filters
    const dayOf = (tag) => days.find((x) => x.tag === tag);
    const titleOf = (tag, id) => { const d = dayOf(tag); return d ? nameOf(d, id) : id; };
    const events = [
      ...subs.map((s) => ({ at: t(s.ts), s, f: s.act === "quiz" ? "quiz" : "sub" })),
      ...Object.values(latest).filter((s) => s.gradedAt).map((s) => ({ at: t(s.gradedAt), g: s, f: "graded" })),
      ...Object.values(working).map((l) => ({ at: t(l.updated), l, f: "sub" })),
    ].sort((a, b) => b.at - a.at);
    const logList = el("div", { class: "list plog" });
    const dayKey = (ms) => new Date(ms).toDateString();
    const dayName = (ms) => {
      const k = dayKey(ms);
      if (k === new Date().toDateString()) return "Today";
      if (k === new Date(Date.now() - 86400000).toDateString()) return "Yesterday";
      return new Date(ms).toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" });
    };
    let filter = "all", limit = 15;
    const drawLog = () => {
      const list = events.filter((e) => filter === "all" || e.f === filter);
      let prevDay = "";
      const rows = [];
      for (const ev of list.slice(0, limit)) {
        if (dayKey(ev.at) !== prevDay) { prevDay = dayKey(ev.at); rows.push(el("div", { class: "plog-date" }, dayName(ev.at))); }
        const time = el("div", { class: "when mono" }, new Date(ev.at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }));
        if (ev.g) {
          const q = pointsOf(ev.g);
          rows.push(el("div", { class: "row plogrow" }, time, el("span", { class: "tag good" }, "GRADED"),
            el("div", {}, el("div", { class: "title" }, titleOf(ev.g.day, ev.g.act)), el("div", { class: "sub" }, `Your trainer graded it: ${round(q.pts)}/${round(q.max)} points`),
              ev.g.feedback ? el("div", { class: "feedback" }, ev.g.feedback) : "")));
          continue;
        }
        if (ev.l) {
          rows.push(el("div", { class: "row plogrow" }, time, el("span", { class: "tag info" }, "WORKING"),
            el("div", {}, el("div", { class: "title" }, titleOf(ev.l.day, ev.l.act)), el("div", { class: "sub" }, `Saved a draft · ${ev.l.pct}% correct so far`))));
          continue;
        }
        const s = ev.s, q = pointsOf(s), quiz = s.act === "quiz", isLatest = latest[`${s.day}|${s.act}`] === s;
        const f = q.pts / Math.max(1, q.max);
        const tagEl = el("span", { class: "tag " + (!isLatest ? "" : !q.graded ? "warn" : f >= 0.8 ? "good" : f >= 0.5 ? "warn" : "bad") }, quiz ? "QUIZ" : "SUBMITTED");
        const mm = Number(s.manualMax);
        const detail = [`Attempt ${s.attempt || 1}`,
          !isLatest ? "replaced by a later attempt" :
            q.graded ? `${round(q.pts)}/${round(q.max)} points` : `${round(s.auto)}/${round(s.autoMax)} auto-checked · ${mm} ${mm === 1 ? "pt" : "pts"} waiting for the trainer`,
          s.late ? "sent after it closed (LATE)" : ""].filter(Boolean).join(" · ");
        rows.push(el("div", { class: "row plogrow" + (isLatest ? "" : " old") }, time, tagEl,
          el("div", {}, el("div", { class: "title" }, titleOf(s.day, s.act)), el("div", { class: "sub" }, detail))));
      }
      if (!list.length) rows.push(el("div", { class: "row" }, el("span", { class: "muted" },
        events.length ? "Nothing of this kind yet." : "No activity yet. Open an activity and start working: it will appear here.")));
      if (list.length > limit) rows.push(el("div", { class: "row" }, el("button", { class: "btn small", type: "button", onclick: () => { limit = list.length; drawLog(); } },
        `Show all ${list.length} entries`)));
      logList.replaceChildren(...rows);
    };
    const chips = el("div", { class: "pfilters", role: "group", "aria-label": "Show in the log" },
      ...[["all", "All"], ["sub", "Activities"], ["quiz", "Quizzes"], ["graded", "Graded"]].map(([v, txt]) => {
        const n = v === "all" ? events.length : events.filter((e) => e.f === v).length;
        const b = el("button", { type: "button", class: "chip" + (v === filter ? " on" : ""), "aria-pressed": String(v === filter) }, txt, el("span", { class: "c" }, String(n)));
        b.addEventListener("click", () => {
          filter = v; limit = 15;
          chips.querySelectorAll(".chip").forEach((c) => { c.classList.toggle("on", c === b); c.setAttribute("aria-pressed", String(c === b)); });
          drawLog();
        });
        return b;
      }));
    drawLog();

    body.replaceChildren(summary, nextCard,
      notes.length ? el("div", { class: "label" }, "From your trainer", fresh.size ? el("span", { class: "tag solid lbltag" }, `${[...fresh].filter((k) => latest[k]).length || fresh.size} NEW`) : "") : "",
      notes.length ? noteBox : "",
      el("div", { class: "label" }, "My days"),
      dayCards.length ? el("div", { class: "pdays" }, ...dayCards) : el("p", { class: "muted" }, "No activities are open yet."),
      el("div", { class: "label" }, "Activity log"), chips, logList,
      p.partial ? el("p", { class: "small muted", style: "margin-top:10px" }, "The full history of every attempt appears after your trainer updates the lab server.") : "");
  }
}

async function viewDay(n) {
  const day = await getDay(n);
  return withConfig(async (cfg) => {
    const tag = day.tag, res = cfg.results[tag] || {}, local = store.get(K.res(tag), {});
    const statusOf = (open, r, l, rd) => {
      if (rd && open) return [el("span", { class: "tag warn" }, "REDO"), el("span", {}, "Fix it and submit again")];
      if (!open) return [el("span", { class: "tag" }, "CLOSED"), el("span", { class: "muted" }, "Opens when the trainer starts it")];
      if (r) return [el("span", { class: "tag good" }, "SUBMITTED"), el("span", {}, fmtScore(r))];
      if (l && l.pending) return [el("span", { class: "tag warn" }, "UNSENT"), el("span", { class: "muted" }, "Sends when online")];
      if (l && l.checked) return [el("span", { class: "tag info" }, "IN PROGRESS"), el("span", {}, `${Math.round(100 * l.checked.auto / Math.max(1, l.checked.autoMax))}% correct`)];
      return [el("span", { class: "tag" }, "NOT STARTED")];
    };
    const rows = day.activities.map((a) => {
      const open = isOpen(cfg, tag, a.id), r = res[a.id];
      return el(open ? "a" : "div", { class: "row actsrow" + (open ? "" : " closed"), href: open ? `#/act/${n}/${a.id}` : null },
        el("div", { class: "num" }, a.id),
        el("div", { class: "title" }, a.title),
        el("div", { class: "sub c-mode" }, a.mode), el("div", { class: "sub c-min" }, `${a.minutes} min`), el("div", { class: "sub c-pts" }, `${a.points} pts`),
        el("div", { class: "status" }, ...statusOf(open, r, local[a.id], redoOf(cfg, tag, a.id))),
        el("div", { class: "chev" }, open ? icon("chev") : icon("padlock")),
        redoOf(cfg, tag, a.id) ? el("div", { class: "feedback redo" }, "Your trainer asked you to redo this" +
          (redoOf(cfg, tag, a.id).note ? `: ${redoOf(cfg, tag, a.id).note}` : ". Your answers are still there: fix them and submit again.")) :
        cfg.unlocked && cfg.unlocked[tag] && cfg.unlocked[tag][a.id] ? el("div", { class: "feedback unl" }, "Your trainer unlocked this activity for you.") : "",
        r && r.feedback ? el("div", { class: "feedback" }, "Trainer feedback: " + r.feedback) : "");
    });
    const qOpen = isOpen(cfg, tag, "quiz"), qr = res.quiz;
    const head = el("div", { class: "listhead actsrow acts" }, el("div", {}, "No."), el("div", {}, "Activity"), el("div", {}, "Mode"),
      el("div", {}, "Time"), el("div", {}, "Points"), el("div", {}, "Status"), el("div", {}));
    shell(`Day ${n}`, "#/", dayTabs(n, "day"),
      el("div", { class: "pagehead" }, el("div", { class: "kicker" }, `Day ${n} · ${day.unit}`), el("h2", {}, day.title)),
      VIDEOS[n] ? el("div", { class: "toolrow" }, el("a", { class: "btn", href: `#/lesson/${n}/0` }, icon("play"), "Watch the demo video")) : "",
      el("div", { class: "label" }, "Activities"),
      day.activities.some((a) => a.sheet) ? el("p", { class: "small muted daynote" }, "Activities with a Workbook save your work in the lab. Tap ",
        el("b", {}, "Submit"), " when done; no need to copy it to Excel or upload a file.") : "",
      head, el("div", { class: "list" }, ...rows),
      el("div", { class: "label" }, "Self-check"),
      el("div", { class: "list", style: "border-top:1px solid var(--line)" },
        el(qOpen ? "a" : "div", { class: "row actsrow" + (qOpen ? "" : " closed"), href: qOpen ? `#/quiz/${n}` : null },
          el("div", { class: "num" }, "Q"), el("div", { class: "title" }, day.quiz.title),
          el("div", { class: "sub c-mode" }, "Individual"), el("div", { class: "sub c-min" }, "20 min"), el("div", { class: "sub c-pts" }, `${day.quiz.items.length} items`),
          el("div", { class: "status" }, ...(!qOpen ? [el("span", { class: "tag" }, "CLOSED")] : qr ? [el("span", { class: qr.auto / qr.autoMax >= 0.8 ? "tag good" : "tag warn" },
            `${qr.auto}/${qr.autoMax}`), el("span", { class: "muted" }, "last attempt")] : [el("span", { class: "tag" }, "PASS MARK 80%")])),
          el("div", { class: "chev" }, qOpen ? icon("chev") : icon("padlock")))));
  });
}
const fmtScore = (r) => {
  const graded = r.manualMax ? (r.manual !== "" && r.manual != null) : true;
  const total = Number(r.auto || 0) + Number(r.manual || 0);
  return graded ? `${round(total)}/${round(Number(r.autoMax) + Number(r.manualMax || 0))}` :
    `${round(r.auto)}/${round(r.autoMax)} auto · ${r.manualMax} pts waiting for trainer`;
};
const round = (x) => Math.round(Number(x) * 10) / 10;

/** A video box that loads the Google Drive player only when tapped (saves mobile data); one player at a time. */
function videoBox(url, title, sub, chapterList, start = 0) {
  const stage = el("div", { class: "vstage" });
  const play = el("button", { class: "vplay", type: "button", "aria-label": `Play: ${title}` },
    el("span", { class: "vtri" }), el("span", {}, start ? `Play from ${fmtSec(start)}` : "Play video"));
  stage.append(play);
  const open = (at) => {
    if (!navigator.onLine) return toast("You are offline. Connect to the internet to watch the video.", "bad");
    document.querySelectorAll(".vstage iframe").forEach((f) => {                 // stop any other player (or this one, to jump)
      const s = f.parentElement; f.remove(); s.classList.remove("on"); s.querySelector(".vplay").hidden = false;
    });
    play.hidden = true;
    stage.classList.add("on");
    // Drive's player begins at ?t=<seconds>, so the video opens right at the chapter
    stage.append(el("iframe", { src: embedUrl(url, at), title, allow: "autoplay; fullscreen", allowfullscreen: true, loading: "lazy", referrerpolicy: "no-referrer" }));
  };
  play.addEventListener("click", () => open(start));
  // chapter rows (Demo video page): tap one to play the video from that chapter
  if (chapterList) chapterList.querySelectorAll("tr[data-start]").forEach((tr) => {
    const go = () => { open(Number(tr.dataset.start)); stage.scrollIntoView({ block: "nearest", behavior: "smooth" }); };
    tr.addEventListener("click", go);
    tr.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
  });
  return el("div", { class: "vbox" },
    el("div", { class: "vhead" }, el("div", { class: "vlabel" }, "Watch it in Excel"), el("div", { class: "vtitle" }, title), sub ? el("div", { class: "vsub" }, sub) : ""),
    stage, chapterList || "",
    el("p", { class: "vnote" }, "If the video does not play, ", el("a", { href: url, target: "_blank", rel: "noopener" }, "open it in Google Drive"), "."));
}

/** The video link for a day: the one the trainer saved on the board (Google Sheet), else the default in videos.js. */
const videoUrl = (cfg, n) => (((cfg && cfg.links) || {})[`day${String(n).padStart(2, "0")}`] || {}).video || (VIDEOS[n] && VIDEOS[n].url) || "";

async function viewLesson(n, k) {
  const day = await getDay(n);
  if (demoMode()) return renderLesson(n, day, {}, k);
  return withConfig((cfg) => renderLesson(n, day, cfg, k));
}
function renderLesson(n, day, cfg, k) {
  const vid = VIDEOS[n] ? { ...VIDEOS[n], url: videoUrl(cfg, n) } : null;
  const chapterFor = (h) => vid && vid.chapters.find((c) => h.startsWith(c.part + " ") || h === c.part);
  // the parts of the lesson: the demo video first (if any), then every section of the handout
  const items = [];
  if (vid) items.push({ title: "Demo video", video: true, short: "Demo video", body: () =>
    videoBox(vid.url, `Day ${n} demonstration: ${day.title}`, "The whole demonstration. Tap a chapter below to jump to it. The lesson parts that have a VIDEO tag also play their own chapter.",
      el("table", { class: "vchapters" }, el("tbody", {}, ...vid.chapters.slice().sort((x, y) => toSec(x.start) - toSec(y.start)).map((c) =>
        el("tr", { "data-start": toSec(c.start), tabindex: 0, role: "button", title: `Play from ${c.start}` },
          el("td", { class: "mono" }, el("span", { class: "vtri-sm" }), `${c.start}${c.end ? "–" + c.end : ""}`), el("td", {}, c.title), el("td", { class: "muted" }, c.part)))))) });
  for (const sec of day.lesson) {
    const ch = chapterFor(sec.h);
    items.push({ title: sec.h, video: !!ch, short: sec.h.replace(/^Part (\d+) — /, "$1 · "), body: () => [el("div", { class: "prose", html: sec.html }),
      ch ? videoBox(vid.url, ch.title, `This part of the demo plays from ${ch.start}` + (ch.end ? ` and ends at ${ch.end}: stop the video there, or keep watching the next part` : "") +
        ".", null, toSec(ch.start)) : ""] });
  }
  const tag = day.tag, rk = `dal3_read_${uid()}_${tag}`, lk = `dal3_lastpart_${uid()}_${tag}`;
  const read = new Set(store.get(rk, []));
  if (!(k >= 0 && k < items.length)) k = Math.min(Math.max(0, store.get(lk, 0)), items.length - 1);
  read.add(k);
  store.set(rk, [...read]);
  store.set(lk, k);
  const it = items[k];

  const toc = el("ol", { class: "toc" }, ...items.map((x, i) => el("li", {},
    el("a", { href: `#/lesson/${n}/${i}`, class: (i === k ? "on " : "") + (read.has(i) ? "read" : ""), "aria-current": i === k ? "page" : null },
      el("span", { class: "tk" }, read.has(i) ? icon("check") : ""), el("span", { class: "tt" }, x.short),
      x.video ? el("span", { class: "tag vtag" }, "VIDEO") : ""))));
  const done = read.size, pct = Math.round(100 * done / items.length);
  const progress = el("div", { class: "lprog" }, el("div", { class: "lbar" }, el("span", { style: `width:${pct}%` })),
    el("span", { class: "small muted" }, `${done} of ${items.length} parts read`));
  const contents = el("details", { class: "tocbox" }, el("summary", {}, icon("book"), el("span", {}, "Contents"),
    el("span", { class: "small muted" }, `Part ${k + 1} of ${items.length}`)), toc);
  if (window.innerWidth > 900) contents.open = true;

  const prev = k > 0 ? el("a", { class: "btn", href: `#/lesson/${n}/${k - 1}` }, icon("back"), el("span", {}, "Previous")) : el("span", {});
  const next = k < items.length - 1
    ? el("a", { class: "btn primary", href: `#/lesson/${n}/${k + 1}` }, el("span", {}, "Next: "), el("span", { class: "nx" }, items[k + 1].short), icon("chev"))
    : el("a", { class: "btn primary", href: `#/day/${n}` }, el("span", {}, "Finish: go to the activities"), icon("chev"));
  const article = el("article", { class: "lpart" },
    el("div", { class: "lpart-head" }, el("div", { class: "kicker" }, `Part ${k + 1} of ${items.length}`), el("h2", {}, it.title)),
    ...[].concat(it.body()),
    el("div", { class: "lnav" }, prev, next));
  shell(`Day ${n} › Lesson`, `#/day/${n}`, dayTabs(n, "lesson"),
    el("div", { class: "pagehead" }, el("div", { class: "kicker" }, `Day ${n} lesson`), el("h2", {}, day.title), progress),
    el("div", { class: "lreader" }, el("aside", { class: "lside" }, contents), article));
}
const toSec = (t) => String(t).split(":").reduce((a, x) => a * 60 + Number(x), 0);
const fmtSec = (n) => `${Math.floor(n / 60)}:${String(Math.round(n % 60)).padStart(2, "0")}`;

const kb = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1024)) + " KB");
/** One downloadable file: PDFs can also be opened in the browser. */
function fileRow(f) {
  const kind = f.type === "pdf" ? "PDF" : f.type.toUpperCase();
  return el("div", { class: "frow" },
    el("span", { class: "ftype ft-" + f.type }, kind),
    el("div", { class: "fmain" }, el("div", { class: "fname" }, f.label), el("div", { class: "fsub" }, `${f.name} · ${kb(f.size)}`),
      f.note ? el("div", { class: "fsub" }, f.note) : ""),
    el("div", { class: "fact" },
      f.type === "pdf" ? el("a", { class: "btn small", href: f.path, target: "_blank", rel: "noopener" }, "Open") : "",
      el("a", { class: "btn small primary", href: f.path, download: f.name }, icon("download"), el("span", { class: "hide-sm" }, "Download"))));
}
async function viewFiles(n) {
  const day = await getDay(n);
  const files = day.files || [];
  const pdfs = files.filter((f) => f.group === "pdf"), data = files.filter((f) => f.group !== "pdf");
  shell(`Day ${n} › Files`, `#/day/${n}`, dayTabs(n, "files"),
    el("div", { class: "pagehead" }, el("div", { class: "kicker" }, `Day ${n} files`), el("h2", {}, day.title),
      el("div", { class: "meta" }, "Download the files to keep, print or use in Microsoft Excel. You do not need them for the lab: the lab's Workbook already has the same data.")),
    pdfs.length ? el("div", { class: "label" }, "Reading and activity sheets") : "", pdfs.length ? el("div", { class: "flist" }, ...pdfs.map(fileRow)) : "",
    data.length ? el("div", { class: "label" }, "Excel files for the activities") : "", data.length ? el("div", { class: "flist" }, ...data.map(fileRow)) : "",
    data.length ? el("p", { class: "small muted", style: "margin-top:10px" }, "On a phone, Excel files open in the Microsoft Excel or Google Sheets app. " +
      "Use them when your trainer asks you to work in Excel; otherwise do the activities in the lab.") : "",
    !files.length ? el("p", { class: "muted" }, "No files for this day yet.") : "");
}

async function viewSheet(n) {
  const day = await getDay(n);
  const host = el("div", { class: "sheethost" });
  shell(`Day ${n} › Workbook`, `#/day/${n}`, dayTabs(n, "sheet"),
    el("p", { class: "small muted" }, `Your Day ${n} workbook. The activities use this same workbook, so formulas you enter here count there too.`),
    host);
  await loadEngine();
  const wb = new Workbook(day, store.get(K.wb(day.tag)));
  const save = debounce(() => store.set(K.wb(day.tag), wb.serialize()), 600);
  const ui = new SheetUI(host, wb, { onChange: save });
  onLeave(() => { save.flush(); ui.destroy(); });
}

// ------------------------------------------------------------------ SQL playground (free practice on aling_nena.db)
async function viewSql(n) {
  const day = await getDay(n);
  const key = `dal3_sqlpad_${uid()}_${day.tag}`;
  const csvBtn = el("button", { class: "btn small", type: "button" }, icon("download"), "Download CSV");
  const saveSql = debounce((t) => store.set(key, t), 400);
  const ed = sqlEditor({ value: store.get(key, ""), rows: 6, onInput: saveSql, extra: [csvBtn] });
  onLeave(() => saveSql.flush());
  csvBtn.addEventListener("click", () => {
    const r = ed.lastResult();
    if (!r) return toast("Run a query first.");
    downloadCsv(r, `Day${String(n).padStart(2, "0")}_query_result.csv`);
  });
  shell(`Day ${n} › SQL`, `#/day/${n}`, dayTabs(n, "sql"),
    el("p", { class: "small muted" }, "Practice any query on Aling Nena's database: the same aling_nena.db the PC students open in VS Code. Nothing you run can damage it."),
    schemaPanel((t) => insertAtCursor(ed.textarea, t + " ")), ed);
  loadSql().catch((e) => toast(e.message, "bad"));
}

// ------------------------------------------------------------------ activity runner
async function viewActivity(n, actId) {
  const day = await getDay(n);
  let cfg = await loadConfig();
  const tag = day.tag, act = day.activities.find((a) => a.id === actId);
  if (!act) return viewDay(n);
  if (!isOpen(cfg, tag, actId)) cfg = await loadConfig(true);      // the cache may be older than the trainer's switch
  if (!isOpen(cfg, tag, actId)) { toast("This activity is not open yet."); location.hash = `#/day/${n}`; return; }
  noteOpened(tag, actId);
  const allAns = store.get(K.ans(tag), {});
  const ans = (allAns[actId] = allAns[actId] || {});
  const results = store.get(K.res(tag), {});
  const needsSheet = act.sheet || act.parts.some((p) => ["cells", "spec", "textcells"].includes(p.type));
  // Nothing on this phone for an activity the student already worked on (another phone, cleared browser): bring back the last
  // saved answers and workbook from the class server, so a redo starts from their own work.
  const redo = redoOf(cfg, tag, actId);
  let restored = false;
  if (!demoMode() && !Object.keys(ans).some((k) => k !== "partners") && (((cfg.results || {})[tag] || {})[actId] || redo)) {
    try {
      const r = await call("myWork", { day: tag, act: actId }, { quiet: true });
      const pl = r && r.payload;
      if (pl && pl.sub) {
        Object.assign(ans, pl.sub.answers || {});
        store.set(K.ans(tag), allAns);
        if (pl.wb) {
          const local = store.get(K.wb(tag));
          if (!local) store.set(K.wb(tag), pl.wb);
          else store.set(K.wb(tag), { ...local, edits: { ...(pl.wb.edits || {}), ...(local.edits || {}) } });   // this phone's edits win
        }
        restored = true;
      }
    } catch { /* offline, or the server is not updated yet: start from what is on this phone */ }
  }
  if (needsSheet) await loadEngine();
  const wb = needsSheet ? new Workbook(day, store.get(K.wb(tag))) : null;
  let dirty = false, ui = null;
  const saveLocal = debounce(() => {
    if (wb) store.set(K.wb(tag), wb.serialize());
    store.set(K.ans(tag), allAns);
  }, 500);
  const touch = () => { dirty = true; saveLocal(); };

  // ---- left: steps + answers
  const EQUIV = [["F4", "Formulas › Lock Reference ($), or the $ F4 key"], ["Ctrl + D", "Home › Fill Down"], ["Ctrl + R", "Home › Fill Right"],
    ["Double-click fill handle", "Home › Fill to End"], ["Shift + click", "Home › Select range, then select the last cell"],
    ["Alt + =", "Formulas › AutoSum"], ["Insert › Chart", "Insert › Chart"], ["Data › Data Analysis", "Data › Data Analysis"],
    ["Conditional Formatting", "Home › Highlight Cells"], ["Ctrl + Shift + %", "Home › Number Format › Percent"]];
  const isSqlAct = act.parts.some((p) => p.type === "sql") || !!act.import;
  const SQL_EQUIV = [["Select the query, Ctrl + E, Ctrl + E", "Run (or Ctrl + Enter on a keyboard)"], ["SQLTools tree: Aling Nena › Tables", "Tables (tap a name to insert it)"],
    ["Results › EXPORT › Save results as CSV", "Download it as CSV"], ["Excel: Data › From Text/CSV › Load", "Import my Q11 result"], ["Ctrl + S", "Saved automatically"]];
  const actFiles = (day.files || []).filter((f) => f.group !== "pdf" || /Activities/i.test(f.name));
  const filesBox = actFiles.length ? el("details", { class: "actfiles" }, el("summary", {}, icon("download"), el("span", {}, "Files for this activity"),
    el("span", { class: "small muted" }, `${actFiles.length} file${actFiles.length > 1 ? "s" : ""}`)), el("div", { class: "flist" }, ...actFiles.map(fileRow))) : "";
  const steps = el("div", { class: "task-sec steps" }, el("div", { class: "label" }, "Instructions"), filesBox,
    el("div", { class: "prose", html: act.html }),
    el("details", {}, el("summary", { class: "small", style: "cursor:pointer;margin-top:8px;color:var(--info)" }, isSqlAct ? "VS Code keys and their lab equivalents" : "Excel keys and their lab equivalents"),
      el("table", { class: "equiv" }, el("thead", {}, el("tr", {}, el("th", {}, isSqlAct ? "In VS Code / Excel" : "In Excel"), el("th", {}, "In the lab"))),
        el("tbody", {}, ...(isSqlAct ? SQL_EQUIV : EQUIV).map(([x, y]) => el("tr", {}, el("td", { class: "mono" }, x), el("td", {}, y)))))));
  const partEls = {};
  const form = el("div", { class: "parts" });
  const hasSql = act.parts.some((p) => p.type === "sql");
  if (hasSql) {
    form.append(schemaPanel((t) => insertAtCursor(null, t + " ")));
    loadSql().catch((e) => toast(e.message, "bad"));
  }
  if (act.import && wb) {
    // Day 5: "export" the joined dataset from SQL and "import" it into the workbook (Excel's Data › From Text/CSV)
    const imp = act.import;
    const resultOf = async () => {
      const src = (allAns[imp.from] || {})[imp.part];
      if (!src || !src.sql || !String(src.sql).trim()) throw new Error(`Write ${imp.part} in Activity ${imp.from} first.`);
      return runSql(src.sql);
    };
    const bImp = el("button", { class: "btn small primary", type: "button" }, icon("grid"), `Import my ${imp.part} result`);
    const bCsv = el("button", { class: "btn small", type: "button" }, icon("download"), "Download it as CSV");
    bImp.addEventListener("click", async () => {
      try {
        const r = await resultOf();
        wb.removeExtraSheet(imp.sheet);
        wb.addOutputSheet(imp.sheet, [r.columns, ...r.rows], {});
        // The imported sheet is too big for the class server (a Sheets cell holds ~50,000 characters), so it is
        // dropped from the saved copy. Keep the query that made it: the trainer's view runs it again to rebuild the sheet.
        wb.specs.imports = { ...(wb.specs.imports || {}), [imp.sheet]: String((allAns[imp.from] || {})[imp.part].sql) };
        // formulas typed before the import pointed at a missing sheet: enter them again so they recalculate
        for (const [k, v] of Object.entries(wb.edits)) if (typeof v === "string" && v.startsWith("=") && v.includes(imp.sheet)) {
          const g = window.WBCore.parseRef(k);
          if (g) wb.setCell(g.sheet, g.r1, g.c1, v);
        }
        touch();
        toast(`Imported ${r.rows.length} rows and ${r.columns.length} columns into the sheet ${imp.sheet}.`, "good");
        showPane("sheet");
        ui && ui.switchSheet(imp.sheet);
      } catch (e) { toast(e.message, "bad"); }
    });
    bCsv.addEventListener("click", async () => {
      try { downloadCsv(await resultOf(), `${imp.sheet}.csv`); } catch (e) { toast(e.message, "bad"); }
    });
    form.append(el("div", { class: "part importbox" }, el("div", { class: "phead" }, el("span", { class: "pprompt" }, imp.prompt)),
      el("p", { class: "small muted" }, `Runs your ${imp.part} from Activity ${imp.from} and loads the result into a new sheet called ${imp.sheet}, like Excel's Data › From Text/CSV.`),
      el("div", { class: "btnrow" }, bImp, bCsv)));
  }
  for (const p of act.parts) {
    const box = el("div", { class: "part", "data-part": p.id });
    const head = el("div", { class: "phead" }, el("span", { class: "pprompt" }, p.prompt || p.label || ""),
      el("span", { class: "ppts" }, `${p.points} ${p.points === 1 ? "pt" : "pts"}${p.manual ? " · graded by trainer" : ""}`));
    const fb = el("div", { class: "pfb" });
    box.append(head);
    if (p.type === "cells") {
      const refs = summarizeRefs(p.cells);
      box.append(el("p", { class: "small muted" }, `${p.cells.length} cells: ${refs}`));
      const go = el("button", { class: "btn small", type: "button" }, icon("grid"), "Open in workbook");
      go.addEventListener("click", () => { showPane("sheet"); ui && ui.switchSheet(p.cells[0].split("!")[0]); });
      box.append(go);
    } else if (p.type === "sql") {
      const cur = ans[p.id] && typeof ans[p.id] === "object" ? ans[p.id] : {};
      const ed = sqlEditor({ value: cur.sql || "", rows: 3,
        onInput: (t) => { ans[p.id] = { sql: t, fp: "" }; touch(); } });
      box.append(ed);
    } else if (p.type === "num" || p.type === "list") {
      const i = el("input", { type: "text", inputmode: p.type === "num" ? "decimal" : "text", value: ans[p.id] ?? "", placeholder: p.type === "list" ? "e.g. 30, 115, 120" : "Type a number" });
      i.addEventListener("input", () => { ans[p.id] = i.value; touch(); });
      box.append(i);
    } else if (p.type === "choice" || p.type === "multi") {
      const g = el("div", { class: "choices", role: p.type === "choice" ? "radiogroup" : "group" });
      for (const [val, label] of p.options) {
        const inp = el("input", { type: p.type === "choice" ? "radio" : "checkbox", name: `${actId}-${p.id}`, value: val });
        if (p.type === "choice") inp.checked = ans[p.id] === val;
        else inp.checked = (ans[p.id] || []).includes(val);
        inp.addEventListener("change", () => {
          if (p.type === "choice") ans[p.id] = val;
          else ans[p.id] = [...g.querySelectorAll("input:checked")].map((x) => x.value);
          touch();
        });
        g.append(el("label", { class: "choice" }, inp, el("span", {}, label)));
      }
      box.append(g);
    } else if (p.type === "text") {
      const t = el("textarea", { rows: p.rows || 3, placeholder: "Write your answer" });
      t.value = ans[p.id] || "";
      t.addEventListener("input", () => { ans[p.id] = t.value; touch(); });
      box.append(t);
    } else if (p.type === "textcells") {
      for (const ref of p.cells) {
        const g = window.WBCore.parseRef(ref);
        const t = el("textarea", { rows: 2, placeholder: `${ref}: one insight sentence` });
        t.value = String(wb.raw(g.sheet, g.r1, g.c1) || "");
        t.addEventListener("change", () => { wb.setCell(g.sheet, g.r1, g.c1, t.value); touch(); });
        box.append(el("label", { class: "small muted" }, ref), t);
      }
      if (act.rubric) box.append(el("p", { class: "small muted" }, "Rubric: " + act.rubric));
    } else if (p.type === "spec") {
      box.append(el("p", { class: "small muted" }, { chart: "Workbook: Insert › Chart.", toolpak: "Workbook: Data › Data Analysis.", highlight: "Workbook: Home › Highlight Cells." }[p.check] || ""));
    }
    box.append(fb);
    partEls[p.id] = { box, fb };
    form.append(box);
  }
  // Pair / Group activities: who did you work with? (the trainer sees partners side by side and can grade them together)
  if (/pair|group/i.test(act.mode) && !demoMode()) {
    const max = /group/i.test(act.mode) ? 3 : 1;
    const list = el("div", { class: "choices partnerlist" }, el("span", { class: "small muted" }, "Loading your classmates…"));
    const box = el("div", { class: "part partners", "data-part": "partners" },
      el("div", { class: "phead" }, el("span", { class: "pprompt" }, `Who did you work with? (${act.mode})`), el("span", { class: "ppts" }, "optional")),
      el("p", { class: "small muted" }, `Tick your ${max > 1 ? "group mates" : "partner"}. Your trainer sees your work side by side and can grade you together. Each of you still taps Submit.`), list);
    form.insertBefore(box, form.firstChild);
    loadClassmates().then((mates) => {
      const pick = new Set(ans.partners || []);
      list.replaceChildren(...mates.map((c) => {
        const inp = el("input", { type: "checkbox", value: c.id });
        inp.checked = pick.has(c.id);
        inp.addEventListener("change", () => {
          if (inp.checked && pick.size >= max) { inp.checked = false; return toast(max > 1 ? `Tick up to ${max} group mates.` : "Tick one partner.", "bad"); }
          inp.checked ? pick.add(c.id) : pick.delete(c.id);
          ans.partners = [...pick];
          touch();
        });
        return el("label", { class: "choice" }, inp, el("span", {}, c.name, el("span", { class: "muted mono" }, `  ${c.id}`)));
      }));
    }).catch(() => list.replaceChildren(el("span", { class: "small muted" }, "Your classmates cannot be loaded right now (offline?). Try again later.")));
  }
  // Optional: the student worked in Google Sheets / Excel instead and shares a link (the trainer grades it).
  const hasSheetWork = act.parts.some((p) => p.type === "cells" || p.type === "spec");
  let linkInput = null, lastScore = null;
  const hasLink = () => /^https:\/\//i.test(String(ans.worklink || "").trim());
  /** Points as the server records them: with a shared link, the unearned Workbook points move to the trainer to grade. */
  const asSent = (score, linked = hasLink()) => {
    let linkPts = 0;
    if (linked) for (const p of act.parts) {
      const r = score.parts[p.id];
      if ((p.type === "cells" || p.type === "spec") && r && !r.manual) linkPts += (Number(r.max) || 0) - (Number(r.pts) || 0);
    }
    linkPts = Math.round(linkPts * 100) / 100;
    return { auto: score.auto, autoMax: Math.round((score.autoMax - linkPts) * 100) / 100, manualMax: score.manualMax + linkPts, linkPts };
  };
  if (hasSheetWork) {
    linkInput = el("input", { type: "url", inputmode: "url", value: ans.worklink || "", placeholder: "https://docs.google.com/spreadsheets/d/…" });
    linkInput.addEventListener("input", () => {
      const v = linkInput.value.trim(), had = hasLink();
      if (v) ans.worklink = v; else delete ans.worklink;
      touch();
      if (had !== hasLink() && lastScore) showFeedback(lastScore);       // switch the Workbook feedback to / from "graded from your link"
    });
    form.append(el("div", { class: "part worklink", "data-part": "worklink" },
      el("div", { class: "phead" }, el("span", { class: "pprompt" }, "Optional: did you work in Google Sheets or Excel instead?"), el("span", { class: "ppts" }, "optional")),
      el("p", { class: "small muted" }, "Leave this empty if you worked in the lab's Workbook. Otherwise paste the share link of your file here. First set its sharing to ",
        el("b", {}, "Anyone with the link · Viewer"), ". Your trainer opens it and grades the workbook part."), linkInput));
  }
  const syncKey = `dal3_syncnote_${uid()}`;
  const syncNote = needsSheet && !store.get(syncKey) ? el("div", { class: "syncnote" },
    el("div", {}, el("b", {}, "Your Workbook is saved in the lab. "), "Everything you do in the Workbook tab (formulas, charts, highlights) is saved as you work and recorded when you tap ",
      el("b", {}, "Submit"), ". You do not need to copy it to Excel or upload a file. Worked in Google Sheets or Excel instead? Paste your file's link at the end of the Answers."),
    el("button", { class: "btn small", type: "button", onclick: (e) => { store.set(syncKey, true); e.currentTarget.parentElement.remove(); } }, "Got it")) : "";
  const redoBox = redo ? el("div", { class: "redobox" },
    el("div", { class: "k" }, icon("undo"), "Your trainer asked you to redo this activity"),
    redo.note ? el("div", { class: "nt" }, redo.note) : "",
    el("div", { class: "small" }, (restored ? "Your last submitted answers were brought back. " : "Your answers are still here. ") +
      "Change what you need, then tap Submit again." + (hasSheetWork ? " Shared a Google Sheets / Excel link? You can keep it or replace it in the Submit window." : ""))) : "";
  const left = el("div", { class: "pane pane-task" }, redoBox, syncNote, steps, el("div", { class: "task-sec" }, el("div", { class: "label" }, "Answers"), form));
  const right = needsSheet ? el("div", { class: "pane pane-sheet" }, el("div", { class: "sheethost" })) : null;

  // ---- bottom bar
  const status = el("span", { class: "stxt" }, results[actId]?.pending ? "Submitted, waiting for a connection" :
    results[actId]?.server ? `Submitted: ${fmtScore(results[actId].server)}` : "Ready");
  const btnCheck = el("button", { class: "btn", type: "button" }, icon("check"), "Check");
  const btnSubmit = el("button", { class: "btn primary", type: "button" }, "Submit");
  const btnTurnIn = el("button", { class: "btn", type: "button", title: "Download your answers as an Excel file and turn it in on Google Classroom" },
    icon("download"), el("span", { class: "hide-sm" }, "Turn in to "), "Classroom");
  const bar = el("div", { class: "statusbar" }, status, el("div", { class: "grow" }), btnTurnIn, btnCheck, btnSubmit);

  const seg = el("div", { class: "switch", role: "tablist" });
  const panes = { steps: left, sheet: right };
  const showPane = (k) => {
    if (!panes[k]) return;
    seg.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.k === k));
    Object.entries(panes).forEach(([kk, p]) => p && p.classList.toggle("show", kk === k));
  };
  if (right) {
    for (const [k, label] of [["steps", "Task"], ["sheet", "Workbook"]]) {
      const b = el("button", { type: "button", "data-k": k }, label);
      b.addEventListener("click", () => showPane(k));
      seg.append(b);
    }
  }
  shell(`Day ${n} › Activity ${act.id}`, `#/day/${n}`,
    el("div", { class: "workhead" }, el("h2", {}, `${act.id}  ${act.title}`), el("span", { class: "meta" }, `${act.mode} · ${act.minutes} min · ${act.points} pts`)),
    right ? seg : "", el("div", { class: "panes" + (right ? "" : " single") }, left, right || ""), bar, { __cls: "page work" });
  showPane("steps");
  if (right) {
    ui = new SheetUI(right.firstChild, wb, { sheet: act.sheet || undefined, onChange: touch });
    onLeave(() => ui.destroy());
  }

  // ---- check & submit
  async function evaluate() {
    for (const p of act.parts) if (p.type === "sql") {             // run every query on a fresh database copy
      const cur = ans[p.id];
      if (!cur || typeof cur !== "object" || !String(cur.sql || "").trim()) continue;
      const f = await fingerprintOf(cur.sql, p.ordered);
      Object.assign(cur, { fp: f.fp, n: f.rows ?? null, c: f.cols ?? null, err: f.error || "" });
    }
    const sub = buildSub(act, wb, ans);
    const want = new Map();
    for (const p of act.parts) {
      if (p.type === "cells") for (const ref of p.cells) {
        const c = sub.cells[ref];
        if (c) want.set(`cell|${ref}|${C.normCell(c.v)}`, sha(`${tag}|cell|${ref}|${C.normCell(c.v)}`).then((h) => h === p.hashes[ref]));
      } else if (!p.manual && p.type !== "spec") {
        const a = ans[p.id];
        if (a != null && a !== "") { const nm = C.normAnswer(p, a); want.set(`part|${p.id}|${nm}`, sha(`${tag}|${act.id}|${p.id}|${nm}`).then((h) => h === p.hash)); }
      }
    }
    const keys = [...want.keys()], vals = await Promise.all(want.values());
    const ok = new Map(keys.map((k, i) => [k, vals[i]]));
    const score = C.scoreActivity(act.parts, sub, (kind, id, norm) => !!ok.get(`${kind}|${id}|${norm}`));
    return { sub, score };
  }
  function showFeedback(score) {
    const marks = {};
    lastScore = score;
    const linked = hasLink();
    for (const p of act.parts) {
      const r = score.parts[p.id], { box, fb } = partEls[p.id];
      box.classList.remove("good", "bad", "half", "linked");
      if (r.manual) { fb.textContent = "Your trainer will grade this."; continue; }
      // Worked in Google Sheets / Excel and shared the link: the Workbook part is graded by the trainer from that file.
      if (linked && (p.type === "cells" || p.type === "spec") && r.frac < 1) {
        const right = p.type === "cells" ? Object.values(r.detail).filter((x) => x === 1).length : 0;
        fb.textContent = "You shared a link to your Google Sheets / Excel file, so your trainer grades this part from your file." +
          (right ? ` (${right} of ${p.cells.length} cells done here in the lab already count.)` : "");
        box.classList.add("linked");
        if (p.type === "cells") for (const [ref, v] of Object.entries(r.detail)) if (v === 1) marks[ref] = "ok";
        continue;
      }
      if (p.type === "cells") {
        const d = r.detail, right = Object.values(d).filter((x) => x === 1).length;
        const noF = Object.values(d).filter((x) => x === -1).length, empty = Object.values(d).filter((x) => x === -2).length;
        fb.textContent = `${right} of ${p.cells.length} cells correct.` + (noF ? ` ${noF} cell(s) have a typed number: use a formula (start with =).` : "") +
          (empty ? ` ${empty} still empty.` : "") + (right < p.cells.length ? " Wrong cells are red on the sheet." : " Great work!");
        for (const [ref, v] of Object.entries(d)) if (v !== -2) marks[ref] = v === 1 ? "ok" : "bad";
      } else if (p.type === "sql") {
        const a = ans[p.id] || {}, x = p.expect || {};
        fb.textContent = r.frac === 1 ? `Correct: ${a.n} row${a.n === 1 ? "" : "s"}.` : !String(a.sql || "").trim() ? "Write your query, then tap Check." :
          a.err ? "Your query has an error: " + a.err :
          a.n === x.rows && a.c === x.cols ? `Almost: the size is right (${a.c} columns, ${a.n} rows) but some values are different.` +
            (p.ordered ? " Check the ORDER BY too: the order of the rows matters here." : " Check your conditions and columns.") :
          `Not correct yet. The answer has ${x.cols} column${x.cols === 1 ? "" : "s"} and ${x.rows} row${x.rows === 1 ? "" : "s"}; yours has ${a.c} and ${a.n}.`;
      } else {
        fb.textContent = r.frac === 1 ? "Correct." : p.type === "spec" ? "Not found yet. Check the range and the settings." : "Not correct yet. Try again.";
      }
      box.classList.add(r.frac === 1 ? "good" : r.frac > 0 ? "half" : "bad");
    }
    ui && ui.setMarks(marks);
  }
  // ---- turn in to Google Classroom: download the answers file, then open the assignment the trainer linked
  const ios = /iPhone|iPad|iPod/i.test(navigator.userAgent);
  /** The next place to go after this activity: the next OPEN activity, else the open quiz, else the Day page. */
  async function nextTarget() {
    const c = await loadConfig(true).catch(() => cfg);
    const i = day.activities.findIndex((a) => a.id === actId);
    const nxt = day.activities.slice(i + 1).find((a) => isOpen(c, tag, a.id));
    if (nxt) return { href: `#/act/${n}/${nxt.id}`, label: `Activity ${nxt.id}`, title: nxt.title };
    if (isOpen(c, tag, "quiz")) return { href: `#/quiz/${n}`, label: "the self-check quiz", title: day.quiz.title };
    return { href: `#/day/${n}`, label: `Day ${n}`, title: "You have finished all the open activities for today.", done: true };
  }
  /** "Going to … in 5 s" with Go now / Stay here. Stops if the window is closed or the student leaves the page. */
  function countdown(holder, target, lead) {
    let left = 5, timer = null;
    const msg = el("span", {});
    const go = el("a", { class: "btn primary small", href: target.href }, "Go now");
    const stay = el("button", { class: "btn small", type: "button" }, "Stay here");
    const box = el("div", { class: "nextbar" }, el("div", {}, lead ? el("b", {}, lead + " ") : "", msg,
      target.done ? el("div", { class: "small muted" }, target.title) : el("div", { class: "small muted" }, target.title)), el("div", { class: "btnrow" }, go, stay));
    const paint = () => { msg.textContent = `Going to ${target.label} in ${left} s…`; };
    const stop = (text) => { clearInterval(timer); timer = null; msg.textContent = text; stay.remove(); };
    stay.addEventListener("click", () => stop(`Staying here. Tap Go now when you are ready for ${target.label}.`));
    paint();
    timer = setInterval(() => {
      if (!document.body.contains(box) || !location.hash.includes(`/${n}/${actId}`)) return clearInterval(timer);
      left -= 1;
      if (left <= 0) { clearInterval(timer); document.querySelectorAll(".modal-back").forEach((m) => m.remove()); location.hash = target.href; return; }
      paint();
    }, 1000);
    holder.replaceChildren(box);
    return box;
  }
  /** The Classroom link the trainer saved for this activity ("" if none). */
  const classroomLink = async () => {
    const fresh = await loadConfig(true).catch(() => cfg);           // the trainer may have added the link just now
    return ((fresh.links || cfg.links || {})[tag] || {})[actId] || "";
  };
  btnTurnIn.addEventListener("click", () => openTurnIn(false));
  /** Turn in to Google Classroom: download the answers file, then open the assignment the trainer linked. */
  async function openTurnIn(afterSubmit, knownLink) {
    btnTurnIn.disabled = true;
    const label = btnTurnIn.lastChild.textContent;
    btnTurnIn.lastChild.textContent = "Preparing…";
    let file;
    try {
      saveLocal.flush();
      const { sub, score } = await evaluate();
      showFeedback(score);
      file = await buildSubmissionFile({ day, act, wb, ans, sub, score, student: session.get() || {} });
    } catch (e) {
      toast(e.message || String(e), "bad");
      return;
    } finally {
      btnTurnIn.disabled = false;
      btnTurnIn.lastChild.textContent = label;
    }
    const link = knownLink ?? (await classroomLink());
    const submitted = !!(results[actId] && (results[actId].server || results[actId].pending));
    const dl = el("button", { class: "btn primary", type: "button" }, icon("download"), "Download my answers");
    dl.addEventListener("click", () => { downloadBlob(file.blob, file.filename); if (open_) { dl.classList.remove("primary"); open_.classList.add("primary"); } });
    const open_ = link ? el("a", { class: "btn", href: link, target: "_blank", rel: "noopener" }, "Open the Google Classroom assignment") : null;
    const step = (n, title, ...kids) => el("div", { class: "tstep" }, el("span", { class: "tn" }, String(n)), el("div", {}, el("b", { class: "tt" }, title), ...kids));
    const box = el("div", { class: "turnin" },
      afterSubmit ? el("div", { class: "tdone" }, icon("check"), el("div", {}, el("b", {}, "Submitted to the lab."),
        " Your trainer can see your score. Now turn in your work on Google Classroom:")) : "",
      step(1, "Download your answers", el("p", { class: "small muted" }, `File: ${file.filename} (Excel). It has your answers and your workbook.`),
        ios ? el("p", { class: "small" }, "On iPhone: after you tap Download, tap the ", el("b", {}, "Share"), " button, then ", el("b", {}, "Save to Files"), ".") : "", dl),
      link ? step(2, "Open the assignment in Google Classroom", el("p", { class: "small muted" }, "Classroom opens on the assignment for this activity."), open_)
        : step(2, "Open Google Classroom", el("p", { class: "small" }, "Your trainer has not added a Classroom link for this activity yet. Open the assignment in Google Classroom yourself, or ask your trainer.")),
      step(3, "Attach the file and turn in", el("p", { class: "small" }, "In the assignment, tap ", el("b", {}, "Add or create"), " (or ", el("b", {}, "+ Add"),
        "), then ", el("b", {}, "File"), ". Choose the file from your ", el("b", {}, "Downloads"), ios ? " (in the Files app)" : " folder", ", then tap ", el("b", {}, "Turn in"), ".")),
      submitted ? "" : el("p", { class: "note small" }, "Also tap Submit here in the lab, so your trainer sees your score on the class board."));
    if (afterSubmit) {
      const target = await nextTarget();
      const holder = el("div", {}, el("a", { class: "btn", href: target.href }, target.done ? `Back to Day ${n}` : `Next: ${target.label}`, icon("chev")));
      box.append(step(4, target.done ? "All done for today" : "Go to the next activity", el("p", { class: "small muted" },
        target.done ? target.title : `When you have turned it in on Classroom: ${target.label} · ${target.title}`), holder));
      if (open_) {
        open_.addEventListener("click", () => {
          // when the student comes back from Classroom, move on (with a 5-second chance to stay)
          const back = () => {
            if (document.visibilityState !== "visible") return;
            document.removeEventListener("visibilitychange", back);
            if (document.body.contains(holder)) countdown(holder, target, "Welcome back! Turned it in on Classroom?");
          };
          setTimeout(() => document.addEventListener("visibilitychange", back), 300);
        });
      }
    }
    modal(afterSubmit ? `Activity ${act.id}: turn in on Google Classroom` : `Turn in Activity ${act.id}`, box);
  }

  btnCheck.addEventListener("click", async () => {
    const { score } = await evaluate();
    showFeedback(score);
    results[actId] = { ...(results[actId] || {}), checked: { auto: score.auto, autoMax: score.autoMax } };
    store.set(K.res(tag), results);
    const x = asSent(score);
    toast(`Auto-checked: ${round(x.auto)} of ${round(x.autoMax)} pts` + (x.manualMax ? ` (+${round(x.manualMax)} ${x.manualMax === 1 ? "pt" : "pts"} graded by the trainer` +
      (x.linkPts ? ", including your shared file" : "") + ")" : ""));
    saveDraft(true);
  });
  /** True when nothing was done in the Workbook for this activity (no answer cells filled, no chart / ToolPak / highlight made). */
  const sheetEmpty = (sub, score) => act.parts.every((p) =>
    p.type === "cells" ? p.cells.every((r) => !sub.cells[r] || String(sub.cells[r].f ?? "").trim() === "") :
    p.type === "spec" ? !(score.parts[p.id] || {}).frac : true);
  /** The Submit window for activities with a Workbook: the student chooses how the Workbook part is turned in,
   *  their work in the lab's Workbook or a link to their Google Sheets / Excel file. Resolves { changed } or null (cancelled). */
  const askSubmit = (sub, score) => new Promise((resolve) => {
    let done = false, iv = null, m = null;
    const finish = (v) => { if (done) return; done = true; clearInterval(iv); m && m.close(); resolve(v); };
    const empty = sheetEmpty(sub, score);
    const cellParts = act.parts.filter((p) => p.type === "cells" && !p.manual), specParts = act.parts.filter((p) => p.type === "spec" && !p.manual);
    const nCells = cellParts.reduce((a, p) => a + p.cells.length, 0);
    const filled = cellParts.reduce((a, p) => a + p.cells.filter((r) => sub.cells[r] && String(sub.cells[r].f ?? "").trim() !== "").length, 0);
    const right = cellParts.reduce((a, p) => a + Object.values((score.parts[p.id] || {}).detail || {}).filter((x) => x === 1).length, 0);
    const specDone = specParts.filter((p) => (score.parts[p.id] || {}).frac === 1).length;
    const sheetPts = [...cellParts, ...specParts].reduce((a, p) => a + p.points, 0);
    const labNote = empty ? "Your Workbook is still empty, so there is nothing to check yet." :
      [nCells ? `${filled} of ${nCells} cells filled, ${right} correct` : "",
        specParts.length ? `${specDone} of ${specParts.length} chart / ToolPak / highlight step${specParts.length === 1 ? "" : "s"} found` : ""].filter(Boolean).join(" · ") +
      ". The lab checks it now.";
    const input = el("input", { type: "url", inputmode: "url", name: "worklink", value: ans.worklink || "", placeholder: "https://docs.google.com/spreadsheets/d/…" });
    const err = el("p", { class: "small bad" });
    const rLab = el("input", { type: "radio", name: "how", value: "lab" }), rLink = el("input", { type: "radio", name: "how", value: "link" });
    (hasLink() || empty ? rLink : rLab).checked = true;
    const opt = (r, title, note) => el("label", { class: "howopt" }, r, el("span", { class: "howtxt" }, el("b", {}, title), el("span", { class: "small muted" }, note)));
    const opts = [opt(rLab, "My work in the lab's Workbook", labNote),
      opt(rLink, "A link to my Google Sheets / Excel file", "Worked outside the lab? Your trainer opens your file and grades the Workbook part.")];
    const linkBox = el("div", { class: "howlink" }, el("label", {}, "Link to my file", input),
      el("p", { class: "small muted" }, "First set the file's sharing to ", el("b", {}, "Anyone with the link · Viewer"),
        " (Google Sheets: Share › General access; Excel: Share › Copy link)."), err);
    const summary = el("div", { class: "subsum" });
    const manualEmpty = act.parts.filter((p) => p.manual && p.type === "text" && !(ans[p.id] || "").trim()).length;
    const refresh = () => {
      const link = rLink.checked, x = asSent(score, link);
      linkBox.hidden = !link;
      opts.forEach((o) => o.classList.toggle("on", o.querySelector("input").checked));
      summary.replaceChildren(
        el("div", { class: "ln" }, el("span", {}, "Checked by the lab now"), el("b", { class: "mono" }, `${round(x.auto)} / ${round(x.autoMax)} pts`)),
        x.manualMax ? el("div", { class: "ln" }, el("span", {}, link ? "Graded by your trainer (includes your file)" : "Graded by your trainer"),
          el("b", { class: "mono" }, `${round(x.manualMax)} pts`)) : "",
        !link && empty ? el("p", { class: "warnline" }, "Your Workbook is empty, so the Workbook part gets 0 points.") : "",
        manualEmpty ? el("p", { class: "warnline" }, `${manualEmpty} written answer${manualEmpty === 1 ? " is" : "s are"} still empty.`) : "");
    };
    [rLab, rLink].forEach((r) => r.addEventListener("change", () => { refresh(); if (rLink.checked) input.focus(); }));
    input.addEventListener("input", () => { err.textContent = ""; });
    const form = el("form", { class: "form submitbox", novalidate: true },
      el("p", { class: "howq" }, `How are you turning in the Workbook part (${round(sheetPts)} pts)?`), ...opts, linkBox, summary,
      el("p", { class: "small muted" }, "You can submit again later; the latest submission counts."),
      el("div", { class: "actions" },
        el("button", { class: "btn", type: "button", onclick: () => finish(null) }, "Cancel"),
        empty ? el("button", { class: "btn", type: "button", onclick: () => { finish(null); showPane("sheet"); } }, "Go to the Workbook") : "",
        el("button", { class: "btn primary", type: "submit", name: "submitgo" }, "Submit")));
    form.addEventListener("submit", (e) => {
      e.preventDefault();
      const before = ans.worklink || "";
      if (rLink.checked) {
        const v = input.value.trim();
        if (!/^https:\/\//i.test(v)) { err.textContent = "Paste the full link of your file. It starts with https://"; input.focus(); return; }
        ans.worklink = v;
      } else delete ans.worklink;                    // turning in the lab's Workbook: no link
      if (linkInput) linkInput.value = ans.worklink || "";
      const changed = (ans.worklink || "") !== before;
      if (changed) touch();
      finish({ changed });
    });
    refresh();
    m = modal(`Submit Activity ${act.id}`, form);
    iv = setInterval(() => { if (!document.body.contains(m.box)) finish(null); }, 400);   // closed with X / Esc
  });
  btnSubmit.addEventListener("click", async () => {
    let { sub, score } = await evaluate();
    if (hasSheetWork) {
      const choice = await askSubmit(sub, score);
      if (!choice) return;
      if (choice.changed) ({ sub, score } = await evaluate());
    } else {
      const manualEmpty = act.parts.filter((p) => p.manual && p.type === "text" && !(ans[p.id] || "").trim()).length;
      const msg = `Submit Activity ${act.id}?\n\nAuto-checked part: ${round(score.auto)} / ${round(score.autoMax)} pts` +
        (score.manualMax ? `\nTrainer-graded part: ${round(score.manualMax)} pts` : "") +
        (manualEmpty ? `\n\nNote: ${manualEmpty} written answer(s) are still empty.` : "") +
        "\n\nYou can submit again later; the latest submission counts.";
      if (!confirm(msg)) return;
    }
    showFeedback(score);
    saveLocal.flush();
    const body = { day: tag, act: act.id, final: true, sub, wb: wb ? wb.serialize() : null, client: { auto: score.auto, autoMax: score.autoMax } };
    if (demoMode()) {
      results[actId] = { ...(results[actId] || {}), server: { auto: score.auto, autoMax: score.autoMax, manualMax: score.manualMax }, pending: false };
      store.set(K.res(tag), results);
      status.textContent = "Saved on this device (demo mode)";
      return toast("Demo mode: saved on this device only.");
    }
    btnSubmit.disabled = true;
    btnSubmit.textContent = "Submitting…";
    try {
      const r = await sendOrQueue("submit", body, { onDoneKey: null });
      if (r.sent) {
        results[actId] = { ...(results[actId] || {}), server: r.data.result, pending: false };
        if (redo) {                                                     // the trainer's redo request is answered
          const c = store.get(`dal3_cfg_${uid()}`);
          if (c && c.redo && c.redo[tag]) { delete c.redo[tag][actId]; store.set(`dal3_cfg_${uid()}`, c); }
          if (redoBox && redoBox.isConnected) redoBox.replaceWith(el("div", { class: "tdone" }, icon("check"), el("span", {}, "Redo sent to your trainer.")));
        }
        status.textContent = `Submitted: ${fmtScore(r.data.result)}`;
        store.set(K.res(tag), results);
        const link = await classroomLink().catch(() => "");
        if (link) openTurnIn(true, link);                                 // next step: Google Classroom
        else {
          const target = await nextTarget();
          const holder = el("div", {});
          const done = el("div", { class: "turnin" }, el("div", { class: "tdone" }, icon("check"), el("div", {}, el("b", {}, "Submitted. "),
            `Your trainer can see your work. Auto-checked score: ${fmtScore(r.data.result)}.`)), holder);
          modal(`Activity ${act.id} submitted`, done);
          countdown(holder, target);
        }
      } else {
        results[actId] = { ...(results[actId] || {}), pending: true };
        status.textContent = "Submitted, waiting for a connection";
        toast("You are offline. Your work is saved and will be sent automatically.");
        nextTarget().then((target) => {                                   // offline: no automatic move, just offer the next step
          modal(`Activity ${act.id} saved`, el("div", { class: "turnin" },
            el("div", { class: "note" }, "You are offline. Your work is saved on this phone and will be sent when you are back online."),
            el("div", { class: "btnrow" }, el("a", { class: "btn primary", href: target.href }, target.done ? `Back to Day ${n}` : `Next: ${target.label}`, icon("chev")))));
        });
      }
      store.set(K.res(tag), results);
      dirty = false;
    } catch (e) {
      toast(e.message, "bad");
    } finally {
      btnSubmit.disabled = false;
      btnSubmit.textContent = "Submit";
    }
  });

  // ---- live autosave to the trainer board
  async function saveDraft(force) {
    if (demoMode() || (!dirty && !force)) return;
    dirty = false;
    const { sub, score } = await evaluate();
    sendOrQueue("save", { day: tag, act: act.id, final: false, sub, wb: wb ? wb.serialize() : null,
      client: { auto: score.auto, autoMax: score.autoMax } }, { replaceKey: `save|${tag}|${act.id}`, quiet: true }).catch(() => {});
  }
  const timer = setInterval(() => saveDraft(false), CONFIG.AUTOSAVE_SECONDS * 1000);
  saveDraft(true);
  onLeave(() => { clearInterval(timer); saveLocal.flush(); if (dirty) saveDraft(true); });
}

function buildSub(act, wb, ans) {
  const refs = act.parts.flatMap((p) => (p.type === "cells" || p.type === "textcells" ? p.cells : []));
  const answers = {};
  for (const p of act.parts) if (ans[p.id] != null) answers[p.id] = ans[p.id];
  if (ans.worklink) answers.worklink = ans.worklink;                   // worked in Google Sheets / Excel instead
  if (ans.partners && ans.partners.length) answers.partners = ans.partners;
  const textcells = act.parts.filter((p) => p.type === "textcells").flatMap((p) => p.cells);
  const sub = { cells: wb ? wb.cellsFor(refs) : {}, answers, specs: wb ? wb.specs : {} };
  if (textcells.length && wb) sub.answers.insights = textcells.map((r) => String(sub.cells[r]?.v ?? "")).join("\n");
  return sub;
}
function summarizeRefs(refs) {
  const by = {};
  refs.forEach((r) => { const [s, a] = r.split("!"); (by[s] = by[s] || []).push(a); });
  return Object.entries(by).map(([s, a]) => `${s}!${a[0]}${a.length > 1 ? "…" + a[a.length - 1] : ""}`).join(", ");
}

// ------------------------------------------------------------------ quiz
async function viewQuiz(n) {
  const day = await getDay(n);
  let cfg = await loadConfig();
  const tag = day.tag;
  if (!isOpen(cfg, tag, "quiz")) cfg = await loadConfig(true);
  if (!isOpen(cfg, tag, "quiz")) { toast("The quiz is not open yet."); location.hash = `#/day/${n}`; return; }
  noteOpened(tag, "quiz");
  const saved = store.get(K.quiz(tag), { answers: {} });
  const form = el("form", { class: "quiz" });
  for (const it of day.quiz.items) {
    const box = el("fieldset", { class: "qitem", "data-n": it.n }, el("legend", {}, `${it.n}. ${it.q}`));
    if (it.type === "choice") {
      it.options.forEach((o, i) => {
        const L = "ABCD"[i];
        const inp = el("input", { type: "radio", name: `q${it.n}`, value: L });
        inp.checked = saved.answers[it.n] === L;
        box.append(el("label", { class: "choice" }, inp, el("span", {}, `${L}. ${o}`)));
      });
    } else {
      const inp = el("input", { type: "text", name: `q${it.n}`, inputmode: it.type === "num" ? "decimal" : "text",
        placeholder: it.type === "formula" ? "=FORMULA(...)" : "Your answer", value: saved.answers[it.n] || "" });
      box.append(inp);
    }
    form.append(box);
  }
  const result = el("div", { class: "result", hidden: true });
  const submit = el("button", { class: "btn primary", type: "submit" }, "Submit quiz");
  form.append(el("div", { class: "submitrow" }, submit));
  shell(`Day ${n} › Quiz`, `#/day/${n}`, dayTabs(n, "quiz"), el("div", { class: "pagehead" }, el("div", { class: "kicker" }, `Day ${n} · ${day.quiz.items.length} items · pass mark 80%`), el("h2", {}, day.quiz.title)), result, form);
  const collect = () => {
    const a = {};
    for (const it of day.quiz.items) {
      const v = it.type === "choice" ? form.querySelector(`input[name=q${it.n}]:checked`)?.value : form[`q${it.n}`].value.trim();
      if (v) a[it.n] = v;
    }
    return a;
  };
  form.addEventListener("input", () => store.set(K.quiz(tag), { ...saved, answers: collect() }));
  const show = (r) => {
    result.hidden = false;
    const pass = r.score / r.max >= day.quiz.pass;
    result.className = "result " + (pass ? "good" : "bad");
    result.innerHTML = `<h3>Score ${r.score} / ${r.max}</h3><p class="small muted" style="margin:2px 0 0">${pass ? "Passed (80% or higher)." :
      "Below the 80% pass mark. Review the lesson, then try again."} Items marked red were incorrect.</p>`;
    for (const [nn, ok] of Object.entries(r.items || {})) form.querySelector(`[data-n="${nn}"]`)?.classList.toggle("wrong", !ok);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  if (saved.result) show(saved.result);
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const answers = collect();
    const missing = day.quiz.items.length - Object.keys(answers).length;
    if (missing && !confirm(`${missing} item(s) have no answer. Submit anyway?`)) return;
    if (demoMode()) return toast("Quiz scoring needs the class server (demo mode).", "bad");
    submit.disabled = true;
    try {
      const r = await sendOrQueue("quiz", { day: tag, answers });
      if (r.sent) { store.set(K.quiz(tag), { answers, result: r.data.result }); show(r.data.result); }
      else toast("You are offline. The quiz will be sent and scored when you are back online.");
    } catch (err) { toast(err.message, "bad"); } finally { submit.disabled = false; }
  });
}

// ------------------------------------------------------------------ trainer board
async function viewTrainer() {
  if (!trainer.get()) return trainerLogin();
  const idx = await getIndex();
  const daySel = el("select", { "aria-label": "Day" }, ...idx.days.map((d) => el("option", { value: d.day }, `Day ${d.day}: ${d.title}`)));
  daySel.value = store.get("dal3_board_day") || idx.days[0].day;
  const live = el("span", { class: "livedot off" }, "Connecting");
  const toggles = el("div", { class: "switches" });
  const table = el("div", { class: "board" });
  const legend = el("div", { class: "legend" });
  const linksBtn = el("button", { class: "btn small", type: "button", title: "Google Classroom assignment link for each activity" }, "Classroom links");
  const videoBtn = el("button", { class: "btn small", type: "button", title: "Google Drive link of this day's demo video" }, icon("play"), "Video link");
  const exportBtn = el("button", { class: "btn small", type: "button", title: "This day's scores as a CSV file" }, icon("download"), "Export CSV");
  const gradesBtn = el("button", { class: "btn small", type: "button", title: "Every day's final scores in one file (for Classroom / TESDA records)" }, icon("download"), "Export grades");
  const queueBtn = el("button", { class: "btn small", type: "button", title: "Open the next submission that needs your grade (amber cells)" }, "Grade next");
  const baseTitle = document.title;
  onLeave(() => { document.title = baseTitle; });
  let prevPending = null;                          // pending keys at the last refresh (to spot new arrivals)
  const arrivals = new Set(), justGraded = new Set();
  const pkey = (x) => `${x.id}|${x.day}|${x.act}`;
  const settings = el("button", { class: "btn small icon", type: "button", title: "Board settings", "aria-label": "Board settings" }, icon("gear"));
  const logout = el("button", { class: "btn small", type: "button", title: "Log out of the trainer board on this browser" }, icon("logout"), "Log out");
  logout.addEventListener("click", trainerLogout);
  shell("Trainer board", "#/", el("div", { class: "boardbar" }, daySel, live, el("div", { class: "grow" }), queueBtn, videoBtn, linksBtn, exportBtn, gradesBtn, settings, logout),
    toggles, table, legend, { __cls: "page work" });
  let data = null;
  settings.addEventListener("click", trainerSettings);
  daySel.addEventListener("change", () => { store.set("dal3_board_day", daySel.value); render(); });
  exportBtn.addEventListener("click", () => data && exportCsv(data, Number(daySel.value)));
  gradesBtn.addEventListener("click", () => data && exportGrades(data));
  queueBtn.addEventListener("click", () => data && openQueue());
  /** Everything waiting for a grade (latest attempts with ungraded written answers or a shared file), oldest first. */
  const pendingList = () => data.subs.filter((x) => Number(x.manualMax) > 0 && (x.manual === "" || x.manual == null))
    .sort((a, b) => new Date(a.ts) - new Date(b.ts));
  const dayNum = (tag) => Number(String(tag).replace(/\D/g, ""));
  /** Next item to grade after x: same activity first (grading one question across the class), then the same day, then any day. */
  const nextAfter = (x) => {
    const q = pendingList().filter((y) => !x || pkey(y) !== pkey(x));
    return (x && (q.find((y) => y.day === x.day && y.act === x.act) || q.find((y) => y.day === x.day))) || q[0] || null;
  };
  /** Options for the detail popup: Save / Save & next, and the table updates at once after a grade. */
  const gradeOpts = (x) => (!x ? {} : {
    nextCount: () => pendingList().filter((y) => pkey(y) !== pkey(x)).length,
    onRedo: () => { render(); setTimeout(() => refresh(true), 1200); },
    onGraded: (goNext, ids) => {
      (ids || [x.id]).forEach((id) => { const k = `${id}|${x.day}|${x.act}`; justGraded.add(k); arrivals.delete(k); setTimeout(() => justGraded.delete(k), 2500); });
      render();
      setTimeout(() => refresh(true), 1200);                      // then sync with the server
      if (!goNext) return;
      const nx = nextAfter(x);
      if (nx) gradeFrom(nx); else toast("Everything is graded.", "good");
    },
  });
  async function gradeFrom(x) {
    const st = data.roster.find((r) => r.id === x.id) || { id: x.id, name: x.id };
    const day = await getDay(dayNum(x.day));
    if (String(day.day) !== String(daySel.value)) { daySel.value = String(day.day); store.set("dal3_board_day", daySel.value); await render(); }
    const it = x.act === "quiz" ? { id: "quiz", title: day.quiz.title } : { id: x.act, title: (day.activities.find((a) => a.id === x.act) || {}).title || "" };
    const lv = data.live.find((l) => l.id === x.id && l.day === x.day && l.act === x.act);
    arrivals.delete(pkey(x));
    openDetail(st, day, it, x, lv, gradeOpts(x));
  }
  /** "Grade next": the oldest waiting item of the day on screen, else of any day (the board switches to that day). */
  function openQueue() {
    const cur = `day${String(daySel.value).padStart(2, "0")}`, q = pendingList();
    if (!q.length) return toast("Nothing is waiting for a grade.", "good");
    gradeFrom(q.find((y) => y.day === cur) || q[0]);
  }
  /** Student tools: unlock activities for this student only, and clear a sign-in lock. */
  async function studentTools(st) {
    const idx = await getIndex();
    const sel = el("select", {}, ...idx.days.map((d) => el("option", { value: d.day }, `Day ${d.day}: ${d.title}`)));
    sel.value = daySel.value;
    const sws = el("div", { class: "switches stools" });
    const lockBox = el("div", {});
    const drawLock = () => {
      const locked = (data.locked || []).includes(st.id);
      lockBox.replaceChildren(locked ? el("div", { class: "note lockbox" }, el("div", {}, el("b", {}, "Sign-in locked. "),
        "Too many wrong PINs. It unlocks by itself after 10 minutes, or unlock it now."),
        el("button", { class: "btn small primary", type: "button", onclick: async (e) => {
          e.currentTarget.disabled = true;
          try { await trainerCall("unlockLogin", { sid: st.id }); data.locked = (data.locked || []).filter((x) => x !== st.id); toast(`${st.name} can sign in again.`, "good"); drawLock(); render(); }
          catch (err) { toast(err.message, "bad"); }
        } }, "Unlock sign-in")) : "");
    };
    const drawSw = async () => {
      const day = await getDay(Number(sel.value)), tag = day.tag;
      const en = data.config[tag] || {}, acc = ((data.access || {})[st.id.toUpperCase()] || {})[tag] || {};
      const items = [...day.activities.map((a) => ({ id: a.id, title: a.title })), { id: "quiz", title: day.quiz.title }];
      sws.replaceChildren(...items.map((it) => {
        if (en[it.id]) return el("span", { class: "sw on classopen", title: `${it.title}: open for the whole class` }, it.id === "quiz" ? "QUIZ" : it.id);
        const on = !!acc[it.id];
        const b = el("button", { class: "sw" + (on ? " on" : ""), type: "button", "aria-pressed": String(on),
          title: `${it.title}: ${on ? "unlocked for " + st.name + " (click to lock again)" : "closed (click to unlock for " + st.name + " only)"}` }, it.id === "quiz" ? "QUIZ" : it.id);
        b.addEventListener("click", async () => {
          b.classList.add("saving"); b.disabled = true;
          try {
            const r = await trainerCall("setAccess", { sid: st.id, day: tag, items: [it.id], enabled: !on });
            data.access = r.access || data.access;
            toast(`${it.id === "quiz" ? "The quiz" : it.id} is now ${!on ? "UNLOCKED" : "LOCKED"} for ${st.name}.`, "good");
            render();
          } catch (err) { toast(err.message, "bad"); }
          drawSw();
        });
        return b;
      }));
    };
    sel.addEventListener("change", drawSw);
    drawLock(); drawSw();
    modal(`${st.name} (${st.id})`, el("div", { class: "form" }, lockBox,
      el("div", { class: "label" }, "Unlock activities for this student only"),
      el("p", { class: "small muted" }, "For make-up work or a student who was absent: the activity opens for this student even while it is closed for the class. " +
        "It stays open until you lock it again. Green with no button = already open for the whole class."),
      sel, sws));
  }
  videoBtn.addEventListener("click", async () => {
    if (!data) return;
    const n = Number(daySel.value), day = await getDay(n), tag = day.tag, v = VIDEOS[n];
    if (!v) return toast(`Day ${n} has no demo video places set up yet.`, "bad");
    const saved = ((data.links || {})[tag] || {}).video || "";
    const f = el("form", { class: "form" });
    const input = el("input", { type: "url", name: "video", value: saved || v.url, placeholder: "https://drive.google.com/file/d/…/view?usp=sharing" });
    const check = el("p", { class: "small" });
    const show = () => {
      const id = driveId(input.value.trim());
      check.textContent = !input.value.trim() ? "Empty: the lesson will use the original video." :
        id ? `Google Drive file found (id ${id.slice(0, 8)}…).` : "This is not a Google Drive file link. Copy it from Drive › Share › Copy link.";
      check.className = "small " + (!input.value.trim() || id ? "muted" : "bad");
    };
    input.addEventListener("input", show);
    show();
    const preview = el("button", { class: "btn small", type: "button" }, icon("play"), "Preview");
    const stage = el("div", { class: "vstage", hidden: true });
    preview.addEventListener("click", () => {
      const src = embedUrl(input.value.trim() || v.url);
      if (!src) return toast("Paste a Google Drive file link first.", "bad");
      stage.hidden = false;
      stage.replaceChildren(el("iframe", { src, title: "Preview", allow: "autoplay; fullscreen", allowfullscreen: true }));
    });
    const places = el("table", { class: "vchapters" }, el("tbody", {},
      el("tr", {}, el("td", { class: "mono" }, `0:00–${v.length || ""}`), el("td", {}, "Full video"), el("td", { class: "muted" }, "Top of the lesson")),
      ...v.chapters.map((c) => {
        const sec = day.lesson.find((x) => x.h.startsWith(c.part + " ") || x.h === c.part);
        return el("tr", {}, el("td", { class: "mono" }, `${c.start}–${c.end}`), el("td", {}, c.title),
          el("td", { class: sec ? "muted" : "bad" }, sec ? sec.h : `${c.part}: section not found`));
      })));
    f.append(el("p", { class: "hint" }, "Paste the Google Drive share link of the Day " + n + " demo video. In Drive, set it to ",
      el("b", {}, "Anyone with the link · Viewer"), " first, or students see a Google sign-in box instead of the video."),
      el("label", {}, "Google Drive link", input), check, el("div", { class: "btnrow" }, preview), stage,
      el("div", { class: "label" }, "Where it appears in the lesson"), places,
      el("p", { class: "small muted" }, "These places and times match the current demo recording. If you upload a different recording " +
        "with other chapter times, the times above need updating too."));
    const save = el("button", { class: "btn primary", type: "submit" }, "Save video link");
    f.append(el("div", { class: "actions" }, el("button", { class: "btn", type: "button", "data-cancel": "" }, "Cancel"), save));
    const m = modal(`Day ${n}: demo video`, f, { wide: true });
    f.addEventListener("submit", async (e) => {
      e.preventDefault();
      let url = input.value.trim();
      if (url && !driveId(url)) return toast("That is not a Google Drive file link.", "bad");
      if (url === v.url) url = "";                                         // same as the original: no override needed
      if (url === saved) return m.close();
      save.disabled = true;
      save.classList.add("busy");
      save.textContent = "Saving…";
      try {
        const r = await trainerCall("setLink", { day: tag, item: "video", url });
        if (r.links) data.links = r.links;
        toast(url ? `Day ${n} lessons now use the new video.` : `Day ${n} lessons use the original video again.`, "good");
        m.close();
      } catch (err) {
        toast(err.message, "bad");
        save.disabled = false; save.classList.remove("busy"); save.textContent = "Save video link";
      }
    });
  });

  linksBtn.addEventListener("click", async () => {
    if (!data) return;
    const day = await getDay(Number(daySel.value)), tag = day.tag, cur = (data.links || {})[tag] || {};
    const f = el("form", { class: "form" });
    f.append(el("p", { class: "hint" }, "In Google Classroom, open the assignment, copy its link from the browser address bar (it starts with " +
      "https://classroom.google.com/), and paste it here. Students then see a button that opens that assignment directly. Leave a box empty to remove a link."));
    const inputs = day.activities.map((a) => {
      const i = el("input", { type: "url", name: a.id, value: cur[a.id] || "", placeholder: "https://classroom.google.com/c/…/a/…/details" });
      f.append(el("label", {}, `${a.id}  ${a.title}`, i));
      return i;
    });
    const save = el("button", { class: "btn primary", type: "submit" }, "Save links");
    f.append(el("div", { class: "actions" }, el("button", { class: "btn", type: "button", "data-cancel": "" }, "Cancel"), save));
    const m = modal(`Day ${day.day}: Google Classroom links`, f);
    f.addEventListener("submit", async (e) => {
      e.preventDefault();
      const bad = inputs.find((i) => i.value.trim() && !/^https:\/\//i.test(i.value.trim()));
      if (bad) return toast(`Activity ${bad.name}: the link must start with https://`, "bad");
      const changed = inputs.filter((i) => (cur[i.name] || "") !== i.value.trim());
      if (!changed.length) return m.close();
      save.disabled = true;
      save.textContent = "Saving…";
      try {
        for (const i of changed) await trainerCall("setLink", { day: tag, item: i.name, url: i.value.trim() });
        const odd = changed.filter((i) => i.value.trim() && !/classroom\.google\.com/i.test(i.value)).map((i) => i.name);
        toast(`Saved ${changed.length} link(s).` + (odd.length ? ` Note: ${odd.join(", ")} is not a classroom.google.com link.` : ""), "good");
        m.close();
        refresh();
      } catch (err) { toast(err.message, "bad"); save.disabled = false; save.textContent = "Save links"; }
    });
  });

  let refreshing = false, pendingSwitch = 0;
  async function refresh(quiet = true) {
    if (refreshing) return;
    refreshing = true;
    live.textContent = "Updating…";
    try {
      const fresh = await trainerCall("progress", {}, { quiet });
      if (pendingSwitch) fresh.config = data ? data.config : fresh.config;   // never undo a switch that is still saving
      data = fresh;
      boardData = fresh;
      const keys = new Set(pendingList().map(pkey));
      if (prevPending) {
        const added = pendingList().filter((x) => !prevPending.has(pkey(x)));
        added.forEach((x) => arrivals.add(pkey(x)));
        if (added.length === 1) {
          const x = added[0], st = data.roster.find((r) => r.id === x.id) || { name: x.id };
          toast(`${st.name} submitted ${x.act === "quiz" ? "the quiz" : x.act} (Day ${dayNum(x.day)}): it needs your grade.`, "good");
        } else if (added.length > 1) toast(`${added.length} new submissions need your grade.`, "good");
      }
      prevPending = keys;
      live.textContent = `Live · updated ${new Date().toLocaleTimeString()}`;
      live.className = "livedot";
      render();
    } catch (e) {
      live.textContent = e.message;
      live.className = "livedot off";
      if (/trainer key/i.test(e.message)) { trainer.clear(); toast("The trainer key was not accepted. Please enter it again.", "bad"); route(); }
    } finally {
      refreshing = false;
    }
  }
  /** Switch activities on/off: the switch changes at once, shows "saving", then confirms or rolls back. */
  async function setOpen(tag, ids, enabled, buttons) {
    const en = (data.config[tag] = data.config[tag] || {});
    const before = ids.map((id) => en[id]);
    ids.forEach((id) => { en[id] = enabled; });
    buttons.forEach((b) => { b.classList.toggle("on", enabled); b.classList.add("saving"); b.disabled = true; b.setAttribute("aria-pressed", String(enabled)); });
    pendingSwitch++;
    try {
      const r = await trainerCall("setConfig", ids.length > 1 ? { day: tag, item: "*", enabled, items: ids } : { day: tag, item: ids[0], enabled });
      if (r.config) data.config = r.config;
      const names = ids.map((id) => (id === "quiz" ? "the quiz" : id));
      toast(`${names.length > 2 ? "All activities" : names.join(" and ")} ${names.length > 1 ? "are" : "is"} now ${enabled ? "OPEN" : "CLOSED"} to students.`, "good");
    } catch (e) {
      ids.forEach((id, i) => { en[id] = before[i]; });
      toast(`Not saved: ${e.message}`, "bad");
    } finally {
      pendingSwitch--;
      render();
    }
  }
  async function render() {
    if (!data) return;
    const day = await getDay(Number(daySel.value)), tag = day.tag;
    const items = [...day.activities.map((a) => ({ id: a.id, title: a.title, points: a.points })), { id: "quiz", title: "Quiz", points: day.quiz.items.length }];
    const en = (data.config[tag] || {});
    toggles.replaceChildren(el("span", { class: "lbl" }, "Open to students"), ...items.map((it) => {
      const b = el("button", { class: "sw" + (en[it.id] ? " on" : ""), type: "button", title: `${it.title}: ${en[it.id] ? "open (click to close)" : "closed (click to open)"}`,
        "aria-pressed": en[it.id] ? "true" : "false" }, it.id === "quiz" ? "QUIZ" : it.id);
      b.addEventListener("click", () => setOpen(tag, [it.id], !en[it.id], [b]));
      return b;
    }), el("button", { class: "btn small", type: "button", onclick: (e) => {
      const btn = e.currentTarget;
      toggles.querySelectorAll(".btn").forEach((x) => { x.disabled = true; });
      btn.classList.add("busy");
      btn.textContent = "Opening…";
      setOpen(tag, items.map((i) => i.id), true, [...toggles.querySelectorAll(".sw")]);
    } }, "Open all"),
    el("button", { class: "btn small", type: "button", onclick: (e) => {
      const btn = e.currentTarget;
      toggles.querySelectorAll(".btn").forEach((x) => { x.disabled = true; });
      btn.classList.add("busy");
      btn.textContent = "Closing…";
      setOpen(tag, items.map((i) => i.id), false, [...toggles.querySelectorAll(".sw")]);
    } }, "Close all"));

    const sub = indexBy(data.subs.filter((s) => s.day === tag), (s) => `${s.id}|${s.act}`);
    const lv = indexBy(data.live.filter((s) => s.day === tag), (s) => `${s.id}|${s.act}`);
    const pend = pendingList(), nPend = pend.length;
    const perAct = {}, perDay = {};
    for (const x of pend) { perDay[x.day] = (perDay[x.day] || 0) + 1; if (x.day === tag) perAct[x.act] = (perAct[x.act] || 0) + 1; }
    const thead = el("tr", {}, el("th", {}, "Student"), ...items.map((it) => el("th", { title: it.title + (perAct[it.id] ? ` · ${perAct[it.id]} waiting for your grade` : "") },
      it.id === "quiz" ? "Quiz" : it.id, perAct[it.id] ? el("span", { class: "gbadge" }, perAct[it.id]) : "")), el("th", {}, "Total"), el("th", {}, "Last active"));
    queueBtn.replaceChildren(nPend ? `Grade next` : "Nothing to grade", nPend ? el("span", { class: "gcount" }, nPend) : "");
    queueBtn.classList.toggle("primary", !!nPend);
    queueBtn.disabled = !nPend;
    document.title = (nPend ? `(${nPend}) ` : "") + baseTitle;
    for (const o of daySel.options) {
      const d = idx.days.find((x) => String(x.day) === o.value), n = perDay[d.tag] || 0;
      o.textContent = `Day ${d.day}: ${d.title}` + (n ? ` · ${n} to grade` : "");
    }
    const nameOf = (id) => (data.roster.find((r) => r.id.toUpperCase() === String(id).toUpperCase()) || { name: id }).name;
    const hm = (t) => new Date(t).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    const rows = data.roster.map((st) => {
      let total = 0, last = null;
      const acc = ((data.access || {})[st.id.toUpperCase()] || {})[tag] || {};
      const tds = items.map((it) => {
        const s = sub[`${st.id}|${it.id}`], l = lv[`${st.id}|${it.id}`];
        const first = (data.opened || {})[`${st.id}|${tag}|${it.id}`];
        if (l && (!last || l.updated > last)) last = l.updated;
        if (s && (!last || s.ts > last)) last = s.ts;
        let txt = "—", cls = "none", title = "";
        if (s) {
          const graded = !s.manualMax || (s.manual !== "" && s.manual != null);
          const pts = Number(s.auto) + Number(s.manual || 0), max = Number(s.autoMax) + Number(s.manualMax || 0);
          total += pts;
          txt = `${round(pts)}/${round(max)}`;
          cls = pts / Math.max(1, max) >= 0.8 ? "good" : pts / Math.max(1, max) >= 0.5 ? "half" : "bad";
          if (!graded) { cls = "needgrade"; if (!Number(s.autoMax)) txt = "GRADE"; }   // waiting for the trainer: amber, not a fail
          if (arrivals.has(`${st.id}|${tag}|${it.id}`)) cls += " arrived";
          if (justGraded.has(`${st.id}|${tag}|${it.id}`)) cls += " justgraded";
          if (s.late) txt += " LATE";
          title = (graded ? "" : "NEEDS YOUR GRADE: click to grade · ") + (first ? `Started ${hm(first)} · ` : "") + `Submitted ${timeAgo(s.ts)} · attempt ${s.attempts}` +
            (s.partner ? ` · with ${s.partner.split(",").map(nameOf).join(", ")}` : "") + (s.link ? " · shared a Google Sheets / Excel file" : "");
        } else if (l) {
          const fresh = Date.now() - new Date(l.updated).getTime() < 3 * 60000;
          txt = `${l.pct}%`;
          cls = "draft" + (fresh ? " fresh" : "");
          title = (first ? `Started ${hm(first)} · ` : "") + `Working · last saved ${timeAgo(l.updated)}`;
        } else if (first) {
          txt = "OPENED"; cls = "seen"; title = `Opened at ${hm(first)} · nothing saved yet`;
        }
        const rq = ((((data.redo || {})[st.id.toUpperCase()] || {})[tag]) || {})[it.id];
        if (rq) { cls = "redo"; txt = "REDO"; title = `Returned for redo ${timeAgo(rq.at)}` + (rq.note ? `: ${rq.note}` : "") + " · waiting for the student to submit again"; }
        else if (acc[it.id] && !en[it.id]) { cls += " unl"; title = (title ? title + " · " : "") + "Unlocked for this student only"; }
        const td = el("td", { class: cls, title }, txt);
        if (s || l) {
          td.tabIndex = 0;
          const open = () => { arrivals.delete(`${st.id}|${tag}|${it.id}`); openDetail(st, day, it, s, l, gradeOpts(s)); };
          td.addEventListener("click", open);
          td.addEventListener("keydown", (e) => { if (e.key === "Enter") open(); });
        }
        return td;
      });
      const locked = (data.locked || []).includes(st.id);
      const nameBtn = el("button", { class: "sbtn", type: "button", title: `${st.name}: unlock activities for this student, sign-in lock` }, st.name,
        el("span", { class: "sid" }, st.id), locked ? el("span", { class: "tag bad" }, "LOCKED") : "");
      nameBtn.addEventListener("click", () => studentTools(st));
      return el("tr", {}, el("th", { class: "sname" }, nameBtn), ...tds,
        el("td", { class: "total" }, round(total)), el("td", { class: "when" }, last ? timeAgo(last) : "—"));
    });
    table.replaceChildren(el("table", { class: "boardtable" }, el("thead", {}, thead), el("tbody", {}, ...rows)));
    legend.replaceChildren(el("span", { class: "lg-good" }, "80% or more"), el("span", { class: "lg-half" }, "50–79%"), el("span", { class: "lg-bad" }, "Below 50%"),
      el("span", { class: "lg-draft" }, "Working now: % of auto-checked points correct"), el("span", { class: "lg-need" }, "Amber = needs your grade (click it). The number on a column = how many are waiting"),
      el("span", {}, "LATE: submitted after you closed it"), el("span", { class: "lg-unl" }, "Unlocked for this student only"),
      el("span", { class: "lg-redo" }, "REDO: returned to the student, waiting for their new submission"),
      el("span", {}, "OPENED: started, nothing saved yet"), el("span", {}, "Select a cell to open the student's work; select a name to unlock activities for one student"));
  }
  refresh(false);
  const t = setInterval(() => refresh(true), CONFIG.BOARD_REFRESH_SECONDS * 1000);
  onLeave(() => clearInterval(t));
}
const indexBy = (arr, f) => Object.fromEntries(arr.map((x) => [f(x), x]));

/** Forget the trainer key on this browser and show the key screen again (stops the board's auto-refresh). */
function trainerLogout() {
  trainer.clear();
  toast("Logged out of the trainer board on this browser.", "good");
  if (location.hash === "#/trainer") route(); else location.hash = "#/trainer";
}

function trainerLogin() {
  const f = el("form", { class: "form signin" });
  f.innerHTML = `<h2>Trainer board</h2><p class="sub">Enter your trainer key. To see it: in your Google Sheet, open the <b>DAL3 Lab</b> menu › <b>Show class code and trainer key</b>.</p>
    ${demoMode() ? `<p class="note">No class server is configured yet. Follow SETUP.md, put the web app URL in js/config.js, or paste it below.</p>
      <label>Web app URL <input name="api" placeholder="https://script.google.com/macros/s/.../exec"></label>` : ""}
    <label>Trainer key <input name="key" type="password" required autocomplete="current-password"></label>
    <button class="btn primary wide" type="submit">Open board</button>`;
  shell("Trainer", "#/", f);
  f.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (f.api && f.api.value.trim()) store.set("dal3_api_override", f.api.value.trim());
    trainer.set(f.key.value.trim());
    try { await trainerCall("progress"); viewTrainer(); }
    catch (err) { trainer.clear(); toast(err.message, "bad"); }
  });
}
function trainerSettings() {
  const box = el("div", { class: "stack" },
    el("p", { class: "small" }, `Server: ${apiUrl() || "(none)"}`),
    el("button", { class: "btn wide", type: "button", onclick: () => {
      const u = prompt("Web app URL (leave empty to use js/config.js):", store.get("dal3_api_override") || "");
      if (u === null) return;
      if (u.trim()) store.set("dal3_api_override", u.trim()); else store.del("dal3_api_override");
      location.reload();
    } }, "Change server URL"),
    el("button", { class: "btn wide danger", type: "button", onclick: () => { m.close(); trainerLogout(); } }, "Log out of the trainer board"));
  const m = modal("Board settings", box);
}

let boardData = null;                       // the trainer board's latest data (roster, subs) for the detail popup
async function openDetail(st, day, item, s, l, opts = {}) {
  const body = el("div", { class: "detail" }, el("p", { class: "muted" }, "Loading…"));
  const m = modal(`${st.name} · ${item.id === "quiz" ? "Quiz" : "Activity " + item.id}`, body, { wide: true });
  const nameOf = (id) => ((boardData && boardData.roster.find((r) => r.id.toUpperCase() === String(id).toUpperCase())) || { name: id }).name;
  const idOf = (id) => ((boardData && boardData.roster.find((r) => r.id.toUpperCase() === String(id).toUpperCase())) || { id }).id;
  const cache = {};
  async function load(attempt) {
    let d = cache[attempt || "latest"];
    if (!d) {
      try { d = await trainerCall("submission", { sid: st.id, day: day.tag, act: item.id, attempt }); }
      catch (e) { body.replaceChildren(el("p", { class: "bad" }, e.message)); return; }
      cache[attempt || "latest"] = d;
    }
    render(d);
  }
  function history(d) {
    const h = d.history || [];
    if (!h.length) return "";
    const cur = d.final ? d.final.attempt || d.final.attempts : 0;
    const fmt = (x) => {
      const graded = !Number(x.manualMax) || (x.manual !== "" && x.manual != null);
      const pts = Number(x.auto || 0) + Number(x.manual || 0), max = Number(x.autoMax || 0) + Number(x.manualMax || 0);
      return graded ? `${round(pts)}/${round(max)}` : `${round(x.auto)}/${round(x.autoMax)} + ${x.manualMax} to grade`;
    };
    const rows = [...h].reverse().map((x) => {
      const tr = el("tr", { class: x.n === cur ? "on" : "", tabindex: 0, title: x.n === cur ? "Showing this attempt" : "Open this attempt" },
        el("td", { class: "mono" }, `#${x.n}${x.n === h.length ? " (latest)" : ""}`),
        el("td", {}, new Date(x.ts).toLocaleString()),
        el("td", { class: "mono" }, fmt(x)),
        el("td", {}, x.late ? el("span", { class: "tag warn" }, "LATE") : ""),
        el("td", { class: "small" }, x.feedback || ""));
      const go = () => { if (x.n !== cur) { body.replaceChildren(el("p", { class: "muted" }, "Loading…")); load(x.n === h.length ? undefined : x.n); } };
      tr.addEventListener("click", go);
      tr.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
      return tr;
    });
    return el("div", { class: "hist" }, el("h3", {}, `Attempt history (${h.length})`),
      el("div", { class: "histwrap" }, el("table", { class: "histtable" },
        el("thead", {}, el("tr", {}, el("th", {}, "Attempt"), el("th", {}, "Sent"), el("th", {}, "Score"), el("th", {}, ""), el("th", {}, "Feedback"))),
        el("tbody", {}, ...rows))),
      h.length > 1 ? el("p", { class: "small muted" }, "Select a row to see what the student sent in that attempt.") : "");
  }
  load();
  /** "Return for redo": the student keeps their answers, fixes them and submits again (a shared link can be replaced). */
  function redoPanel() {
    const panel = el("div", { class: "redoask" });
    const cur = () => ((((boardData && boardData.redo) || {})[st.id.toUpperCase()] || {})[day.tag] || {})[item.id] || null;
    const send = async (body, btn) => {
      btn.disabled = true;
      try {
        const r = await trainerCall("redo", { sid: st.id, day: day.tag, act: item.id, ...body });
        if (boardData) { boardData.redo = r.all || boardData.redo; boardData.access = r.access || boardData.access; }
        toast(body.cancel ? "Redo request taken back." : r.updated ? "Redo note updated." :
          `Returned to ${st.name} for a redo.` + (r.unlocked ? " The activity is open for this student only." : ""), "good");
        draw();
        if (opts.onRedo) opts.onRedo();
      } catch (e) { toast(/unknown action/i.test(e.message) ? "Update the Apps Script server first (SETUP.md: Redo update)." : e.message, "bad"); btn.disabled = false; }
    };
    const draw = () => {
      const rq = cur();
      panel.classList.toggle("on", !!rq);
      const note = el("textarea", { name: "redonote", rows: 2, placeholder: "What should the student fix? (optional, they see this)" }, rq ? rq.note : "");
      if (rq) {
        const upd = el("button", { class: "btn small", type: "button", name: "redoupdate" }, "Update note");
        const back = el("button", { class: "btn small", type: "button", name: "redocancel" }, "Take back the redo");
        upd.addEventListener("click", () => send({ note: note.value.trim() }, upd));
        back.addEventListener("click", () => send({ cancel: true }, back));
        panel.replaceChildren(el("h3", {}, icon("undo"), "Returned for redo ", el("span", { class: "small muted" }, timeAgo(rq.at))),
          el("p", { class: "small muted" }, "Waiting for the student to submit again. Their answers are still there, and the activity is open for them. " +
            "When they submit, the cell turns amber (if it needs your grade) and this request closes by itself."),
          note, el("div", { class: "btnrow" }, upd, back));
      } else {
        const go = el("button", { class: "btn", type: "button", name: "redogo" }, icon("undo"), "Return for redo");
        go.addEventListener("click", () => send({ note: note.value.trim() }, go));
        panel.replaceChildren(el("h3", {}, "Ask for a redo"),
          el("p", { class: "small muted" }, "The student keeps all their answers, fixes what you ask and submits again. A shared Google Sheets / Excel link can be kept or replaced. " +
            "If the activity is closed, it opens for this student only (not marked LATE)."),
          note, el("div", { class: "btnrow" }, go));
      }
    };
    draw();
    return panel;
  }
  function render(d) {
  const rec = d.final || d.live;
  if (!rec) { body.replaceChildren(el("p", {}, "No work yet.")); return; }
  const payload = rec.payload || {};
  const act = day.activities.find((a) => a.id === item.id);
  const older = d.final && d.final.latest === false;
  const mins = d.opened && d.final ? Math.round((new Date(d.final.ts) - new Date(d.opened)) / 60000) : null;
  const kids = [el("p", { class: "small muted" },
    (d.opened ? `Started ${new Date(d.opened).toLocaleString()} · ` : "") +
    (d.final ? `${older ? "Attempt" : "Final submission"} ${new Date(d.final.ts).toLocaleString()} · attempt ${d.final.attempt || d.final.attempts} of ${d.final.attempts}` : "Not submitted yet") +
    (mins != null && mins >= 0 ? ` · ${mins < 120 ? mins + " min" : Math.round(mins / 60) + " h"} from start to this submission` : "") +
    (d.live ? ` · live work saved ${timeAgo(d.live.updated)}` : "")), history(d)];
  const partners = String((d.final && d.final.partner) || "").split(",").map((x) => x.trim()).filter(Boolean);
  if (partners.length) kids.push(el("div", { class: "partnerbox" }, el("b", {}, "Worked with: "),
    ...partners.map((pid) => {
      const b = el("button", { class: "btn small", type: "button", title: "Open their work for this activity" }, `${nameOf(pid)} (${idOf(pid)})`);
      b.addEventListener("click", () => {
        const ps = boardData && boardData.subs.find((x) => x.id.toUpperCase() === pid && x.day === day.tag && x.act === item.id);
        const pl = boardData && boardData.live.find((x) => x.id.toUpperCase() === pid && x.day === day.tag && x.act === item.id);
        m.close();
        openDetail({ id: idOf(pid), name: nameOf(pid) }, day, item, ps, pl);
      });
      return b;
    })));
  const hasText = item.id !== "quiz" && (day.activities.find((a) => a.id === item.id) || { parts: [] }).parts.some((p) => p.type === "text");
  if (hasText && d.final && !older) {
    const cc = el("div", { class: "copycheck" }, el("span", { class: "small muted" }, "Copy check: comparing the written answers with the class…"));
    kids.push(cc);
    trainerCall("similar", { sid: st.id, day: day.tag, act: item.id }, { quiet: true }).then((r) => {
      const ms = r.matches || [];
      if (!ms.length) { cc.replaceChildren(el("span", { class: "small muted" }, `Copy check: no matching written answers among ${r.checked} other student(s)${partners.length ? " (partners not counted)" : ""}.`)); return; }
      cc.classList.add("hit");
      cc.replaceChildren(el("b", {}, "Copy check: very similar written answers"),
        el("ul", { class: "small" }, ...ms.map((x) => el("li", {}, `${x.part}) `, el("b", {}, x.same ? "word-for-word the same" : `${x.pct}% alike`), ` as ${x.name} (${x.sid})`))),
        el("div", { class: "small muted" }, "Partners who worked together are not counted. Read both answers before deciding."));
    }).catch(() => cc.replaceChildren(el("span", { class: "small muted" }, "Copy check is not available until the Apps Script server is updated.")));
  }
  if (older) kids.push(el("p", { class: "note" }, `You are looking at an older attempt (#${d.final.attempt}). Grades and feedback go on the latest attempt.`));
  const worklink = String(((payload.sub || {}).answers || {}).worklink || "").trim();
  if (/^https:\/\//i.test(worklink)) kids.push(el("div", { class: "worklinkbox" },
    el("div", {}, el("b", {}, "The student worked in Google Sheets or Excel and shared this file:"), el("div", { class: "small mono wl" }, worklink),
      el("div", { class: "small muted" }, "Open it, check the workbook part, then give those points in the grading box below.")),
    el("a", { class: "btn primary", href: worklink, target: "_blank", rel: "noopener noreferrer" }, "Open the file")));
  if (item.id === "quiz") {
    const res = d.final && d.final.detail ? d.final.detail.items || {} : {};
    kids.push(el("h3", {}, `Score ${d.final ? d.final.auto + " / " + d.final.autoMax : "—"}`));
    kids.push(el("ol", { class: "small" }, ...day.quiz.items.map((it) => el("li", { class: res[it.n] ? "ok" : "no" },
      `${(payload.answers || {})[it.n] ?? "—"}  ${res[it.n] ? "correct" : "incorrect"}`))));
  } else {
    const det = (d.final && d.final.detail) || (d.live && d.live.detail) || {};
    const parts = (det.parts) || {};
    for (const p of act.parts) {
      const r = parts[p.id] || {};
      const a = (payload.sub && payload.sub.answers || {})[p.id];
      const row = el("div", { class: "part " + (r.manual ? "" : r.frac === 1 ? "good" : r.frac > 0 ? "half" : "bad") },
        el("div", { class: "phead" }, el("span", { class: "pprompt" }, p.prompt || p.label), el("span", { class: "ppts" },
          r.manual ? `${p.points} PTS · TRAINER` : `${r.pts ?? 0}/${p.points}`)));
      if (p.type === "cells") {
        const dd = r.detail || {};
        const right = Object.values(dd).filter((x) => x === 1).length;
        row.append(el("p", { class: "small" }, `${right} of ${p.cells.length} cells correct`));
        const wrong = Object.entries(dd).filter(([, v]) => v !== 1).slice(0, 12);
        if (wrong.length) row.append(el("ul", { class: "small mono" }, ...wrong.map(([ref, v]) => {
          const c = (payload.sub.cells || {})[ref] || {};
          return el("li", {}, `${ref}: ${c.f || "(empty)"} → ${c.v ?? ""}${v === -1 ? "  (typed, not a formula)" : ""}`);
        })));
      } else if (p.type === "textcells") {
        row.append(el("pre", { class: "answer" }, (payload.sub.answers || {}).insights || "(empty)"));
      } else if (p.type !== "spec") {
        row.append(el("pre", { class: "answer" + (p.type === "sql" ? " mono" : "") }, Array.isArray(a) ? a.join(", ") :
          a && typeof a === "object" ? (String(a.sql || "").trim() || "(no query)") + (a.err ? `\n-- error: ${a.err}` : a.n != null ? `\n-- ${a.n} row(s)` : "") :
          (a ?? "(no answer)")));
      } else row.append(el("p", { class: "small" }, r.frac === 1 ? "Done." : "Not found."));
      kids.push(row);
    }
    // grading form
    const manual = act.parts.filter((p) => p.manual);
    const writtenMax = manual.reduce((s, p) => s + p.points, 0);
    const gradeMax = Math.max(Number(d.final && d.final.manualMax) || 0, writtenMax);
    if (gradeMax && d.final && !older) {
      const what = gradeMax > writtenMax ? "the written answers and the linked workbook" : "the written answers";
      const g = el("form", { class: "form grade" + (d.final.manual === "" || d.final.manual == null ? " sticky" : "") });
      const left = opts.nextCount ? opts.nextCount() : 0;
      g.innerHTML = `<h3>Grade ${what} (${gradeMax} pts)</h3>
        <label>Points for this part <input name="pts" type="number" step="0.5" min="0" max="${gradeMax}" value="${d.final.manual ?? ""}"></label>
        <label>Feedback to the student <textarea name="fb" rows="2">${esc(d.final.feedback || "")}</textarea></label>
        ${partners.length ? `<label class="chk"><input type="checkbox" name="together" checked> Give the same points and feedback to ${esc(partners.map((p) => nameOf(p)).join(", "))}</label>` : ""}
        <div class="btnrow">${opts.onGraded && left ? `<button class="btn primary" type="submit" name="next">Save &amp; next <span class="small">(${left} more)</span></button>
          <button class="btn" type="submit" name="save">Save</button>` : `<button class="btn primary" type="submit" name="save">Save grade</button>`}</div>`;
      g.addEventListener("submit", async (e) => {
        e.preventDefault();
        const goNext = !!(e.submitter && e.submitter.name === "next");
        g.querySelectorAll("button[type=submit]").forEach((b) => { b.disabled = true; });
        const btn = { set disabled(v) { g.querySelectorAll("button[type=submit]").forEach((b) => { b.disabled = v; }); } };
        const also = partners.length && g.together && g.together.checked ? partners.map(idOf) : [];
        try {
          const r = await trainerCall("grade", { sid: st.id, day: day.tag, act: item.id, manual: Number(g.pts.value), feedback: g.fb.value, alsoSids: also });
          const done = (r.graded || [st.id]).length, skipped = r.skipped || [];
          toast(done > 1 ? `Grade saved for ${done} students.` : "Grade saved." + (skipped.length ? ` ${skipped.map(nameOf).join(", ")} has not submitted yet.` : ""), "good");
          if (boardData) for (const id of r.graded || [st.id]) {
            const x = boardData.subs.find((y) => y.id.toUpperCase() === String(id).toUpperCase() && y.day === day.tag && y.act === item.id);
            if (x) { x.manual = Math.min(Number(g.pts.value) || 0, Number(x.manualMax) || 0); x.feedback = g.fb.value; }
          }
          m.close();
          if (opts.onGraded) opts.onGraded(goNext, r.graded);
        } catch (err) { toast(err.message, "bad"); btn.disabled = false; }
      });
      kids.push(g);
    }
    if (d.final && !older) kids.push(redoPanel());
    if (payload.wb) {
      const host = el("div", { class: "sheethost ro" });
      const btn = el("button", { class: "btn", type: "button" }, "Show the student's workbook");
      btn.addEventListener("click", async () => {
        btn.disabled = true;
        btn.textContent = "Loading the workbook…";
        try { await loadEngine(); } catch (e) { toast(e.message, "bad"); btn.disabled = false; btn.textContent = "Show the student's workbook"; return; }
        btn.remove();
        const missing = await restoreImports(payload.wb);
        if (missing.length) host.before(el("p", { class: "note small" }, `The sheet ${missing.join(", ")} could not be rebuilt here, so formulas that use it ` +
          "show errors in this view. The scores above were checked on the student's phone with the full sheet."));
        const wb = new Workbook(day, payload.wb);
        const marks = {};
        const cp = act.parts.find((p) => p.type === "cells");
        if (cp && parts.cells && parts.cells.detail) for (const [ref, v] of Object.entries(parts.cells.detail)) if (v !== -2) marks[ref] = v === 1 ? "ok" : "bad";
        new SheetUI(host, wb, { readOnly: true, sheet: act.sheet || undefined, marks });
      });
      kids.push(btn, host);
    }
  }
  body.replaceChildren(...kids);
  }
}

/** Put back the big imported sheets (Day 5.5) that the server could not store, by running the student's saved query again. */
async function restoreImports(wbData) {
  const missing = [];
  const extra = (wbData.extra = wbData.extra || {});
  const imports = (wbData.specs && wbData.specs.imports) || {};
  for (const [name, sql] of Object.entries(imports)) {
    if (extra[name]) continue;
    try {
      const r = await runSql(sql);
      const rows = [r.columns, ...r.rows], cells = {};
      rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined) cells[window.WBCore.addr(i, j)] = [v]; }));
      extra[name] = { name, rows: rows.length, cols: Math.max(1, ...rows.map((x) => x.length)), colW: {}, cells };
    } catch (e) { missing.push(name); }
  }
  // older submissions (before the query was saved): a formula points at a sheet that is not there
  const known = new Set([...Object.keys(extra), ...Object.keys(imports)]);
  for (const v of Object.values(wbData.edits || {})) {
    const m = typeof v === "string" && v.startsWith("=") && v.match(/(Day\d+_[A-Za-z_]+)(?=\[|!)/);
    if (m && !known.has(m[1]) && !missing.includes(m[1])) missing.push(m[1]);
  }
  return missing;
}

/** One CSV with every day's final scores: a max-points row, then one row per student (Excel-friendly UTF-8). */
async function exportGrades(data) {
  const idx = await getIndex();
  const days = await Promise.all(idx.days.map((d) => getDay(d.day)));
  const q = (v) => (/[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v));
  const head = ["Student ID", "Name"], maxRow = ["", "Max points"], cols = [];
  for (const d of days) {
    for (const a of d.activities) { head.push(`D${d.day} ${a.id}`); maxRow.push(a.points); cols.push({ tag: d.tag, id: a.id }); }
    head.push(`D${d.day} Quiz`); maxRow.push(d.quiz.items.length); cols.push({ tag: d.tag, id: "quiz" });
    head.push(`Day ${d.day} total`); maxRow.push(d.activities.reduce((s, a) => s + a.points, 0) + d.quiz.items.length); cols.push({ tag: d.tag, total: true });
  }
  head.push("Grand total", "Waiting for grade"); maxRow.push(maxRow.slice(2).filter((_, i) => cols[i].total).reduce((s, v) => s + Number(v), 0), "");
  const sub = indexBy(data.subs, (s) => `${s.id}|${s.day}|${s.act}`);
  const lines = [head.map(q).join(","), maxRow.map(q).join(",")];
  for (const st of data.roster) {
    let dayTot = 0, grand = 0, waiting = 0;
    const row = [st.id, st.name];
    for (const c of cols) {
      if (c.total) { row.push(round(dayTot)); grand += dayTot; dayTot = 0; continue; }
      const s = sub[`${st.id}|${c.tag}|${c.id}`];
      if (!s) { row.push(""); continue; }
      const v = Number(s.auto || 0) + Number(s.manual || 0);
      if (Number(s.manualMax) > 0 && (s.manual === "" || s.manual == null)) waiting++;
      dayTot += v;
      row.push(round(v));
    }
    row.push(round(grand), waiting);
    lines.push(row.map(q).join(","));
  }
  const a = el("a", { href: URL.createObjectURL(new Blob(["﻿" + lines.join("\r\n")], { type: "text/csv" })),
    download: `DAL3_grades_${new Date().toISOString().slice(0, 10)}.csv` });
  document.body.append(a); a.click(); a.remove();
  toast("Grades exported. Blank = not submitted. 'Waiting for grade' counts items you still need to grade.", "good");
}

function exportCsv(data, n) {
  getDay(n).then((day) => {
    const tag = day.tag, ids = [...day.activities.map((a) => a.id), "quiz"];
    const sub = indexBy(data.subs.filter((s) => s.day === tag), (s) => `${s.id}|${s.act}`);
    const lines = [["Student ID", "Name", ...ids.map((i) => (i === "quiz" ? "Quiz" : "Act " + i)), "Total"].join(",")];
    for (const st of data.roster) {
      let tot = 0;
      const cells = ids.map((i) => { const s = sub[`${st.id}|${i}`]; if (!s) return ""; const v = Number(s.auto) + Number(s.manual || 0); tot += v; return round(v); });
      lines.push([st.id, `"${st.name.replace(/"/g, '""')}"`, ...cells, round(tot)].join(","));
    }
    const a = el("a", { href: URL.createObjectURL(new Blob([lines.join("\n")], { type: "text/csv" })), download: `DAL3_${tag}_scores.csv` });
    document.body.append(a); a.click(); a.remove();
  });
}

// ------------------------------------------------------------------ helpers & router
function debounce(fn, ms) {
  let t, args = null;
  const d = (...a) => { args = a; clearTimeout(t); t = setTimeout(() => { const x = args; args = null; fn(...x); }, ms); };
  // flush: run the waiting call now (with its arguments); a function without arguments always runs, as before
  d.flush = () => { clearTimeout(t); const x = args; args = null; if (x || fn.length === 0) fn(...(x || [])); };
  return d;
}
const loadbar = el("div", { id: "loadbar", role: "progressbar", "aria-label": "Loading" });
document.body.append(loadbar);
onBusy((b) => loadbar.classList.toggle("on", b));

async function route() {
  loadbar.classList.add("route");
  try { await route_(); } finally { loadbar.classList.remove("route"); }
}
async function route_() {
  cleanup.forEach((fn) => { try { fn(); } catch (e) { console.warn(e); } });
  cleanup = [];
  document.querySelectorAll(".modal-back").forEach((m) => m.remove());
  const h = location.hash.replace(/^#\/?/, "").split("/");
  try {
    if (h[0] === "trainer") return await viewTrainer();
    if (!session.get()) return await viewLogin();
    if (h[0] === "day") return await viewDay(Number(h[1]));
    if (h[0] === "lesson") return await viewLesson(Number(h[1]), h[2] === undefined ? undefined : Number(h[2]));
    if (h[0] === "sheet") return await viewSheet(Number(h[1]));
    if (h[0] === "sql") return await viewSql(Number(h[1]));
    if (h[0] === "act") return await viewActivity(Number(h[1]), h[2]);
    if (h[0] === "quiz") return await viewQuiz(Number(h[1]));
    if (h[0] === "files") return await viewFiles(Number(h[1]));
    if (h[0] === "login") return await viewLogin();
    if (h[0] === "progress") return await viewProgress();
    return await viewHome();
  } catch (e) {
    console.error(e);
    shell("Something went wrong", "#/", el("div", { class: "panel panel-pad" }, el("p", {}, e.message), el("a", { class: "btn", href: "#/" }, "Home")));
  }
}
window.addEventListener("hashchange", route);
window.addEventListener("beforeunload", () => cleanup.forEach((fn) => { try { fn(); } catch { /* ignore */ } }));
route();
queue.flush();
// ------------------------------------------------------------------ app updates, in real time
// An open app checks for a new version every minute (and when the tab comes back to the front). When one is ready, a window
// pauses the page: it says the work is saved and lists only what is new FOR THIS PERSON (student or trainer), then
// "Reload now" saves everything still open (a cell being typed, the SQL box, the answers) and reloads.
const SEEN_V = "dal3_seen_version";
const isTrainerView = () => location.hash.startsWith("#/trainer");
async function whatsNew() {
  try { return (await (await fetch(`whatsnew.json?t=${Date.now()}`, { cache: "no-store" })).json()).versions || []; }
  catch { return []; }
}
/** The notes since the last version this device saw (only the newest, if it never saw one), for students or for the trainer. */
function updateNotes(versions) {
  const seen = store.get(SEEN_V, null), side = isTrainerView() ? "trainer" : "student", fresh = [];
  for (const x of versions) { if (x.v === seen) break; fresh.push(x); }
  const list = (seen ? fresh : versions.slice(0, 1)).flatMap((x) => x[side] || []);
  return { list: [...new Set(list)], newest: versions[0] ? versions[0].v : null };
}
let updateOpen = false;
async function showUpdate() {
  if (updateOpen) return;
  updateOpen = true;
  const { list, newest } = updateNotes(await whatsNew());
  const trainerView = isTrainerView(), working = /^#\/(act|sheet|sql|quiz)\//.test(location.hash);
  const btn = el("button", { class: "btn primary", type: "button", name: "updreload" }, "Reload now");
  btn.addEventListener("click", () => {
    btn.disabled = true;
    btn.textContent = "Saving your work…";
    try { SheetUI.commitAll(); } catch (e) { console.warn(e); }
    cleanup.forEach((fn) => { try { fn(); } catch (e) { console.warn(e); } });   // every page saves what is still open
    cleanup = [];
    if (newest) store.set(SEEN_V, newest);
    setTimeout(() => location.reload(), 200);
  });
  const later = trainerView ? el("button", { class: "btn", type: "button", name: "updlater" }, "Finish first (remind me in 1 minute)") : "";
  const back = el("div", { class: "upd-back" });
  if (later) later.addEventListener("click", () => { back.remove(); updateOpen = false; setTimeout(showUpdate, 60000); });
  back.append(el("div", { class: "upd", role: "alertdialog", "aria-modal": "true", "aria-labelledby": "upd-h" },
    el("div", { class: "upd-top" }, icon("download"), el("h3", { id: "upd-h" }, "A new version of the lab is ready")),
    el("p", { class: "upd-safe" }, icon("check"), el("span", {}, trainerView
      ? "Your board stays the same. If you are typing a grade or feedback, finish and save it first."
      : working ? "Your work is saved on this phone. It stays exactly the same after the update: you continue where you stopped."
        : "Your work is saved and stays exactly the same after the update.")),
    list.length ? el("div", { class: "upd-new" }, el("div", { class: "k" }, "What's new for you"), el("ul", {}, ...list.map((t) => el("li", {}, t))))
      : el("p", { class: "small muted" }, "Small fixes and improvements."),
    el("div", { class: "actions" }, later, btn)));
  document.body.append(back);
  btn.focus();
}
if ("serviceWorker" in navigator && location.protocol !== "file:") {
  let controller = navigator.serviceWorker.controller;            // null on the very first visit
  navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).then((reg) => {
    const check = () => { if (navigator.onLine && !updateOpen) reg.update().catch(() => {}); };
    setInterval(check, 60000);
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") check(); });
    window.addEventListener("online", check);
  }).catch(() => {});
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (controller) showUpdate();                                   // a NEW version replaced the one this page runs on
    controller = navigator.serviceWorker.controller;               // (the first install taking control is not an update)
  });
  // first run on this device: remember the current version, so the next update only lists what is new after it
  if (!store.get(SEEN_V, null)) whatsNew().then((v) => { if (v[0] && !store.get(SEEN_V, null)) store.set(SEEN_V, v[0].v); });
}
