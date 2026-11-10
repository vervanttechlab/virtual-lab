// DAL3 Virtual Lab: router and views (student + trainer).
import { CONFIG } from "./config.js";
import { call, trainerCall, sendOrQueue, session, trainer, demoMode, queue, onQueue, onBusy, apiUrl, NetError } from "./api.js";
import { Workbook } from "./workbook.js";
import { SheetUI } from "./grid.js";
import { buildSubmissionFile, downloadBlob } from "./export.js";
import { VIDEOS, embedUrl, driveId } from "./videos.js";
import { el, esc, modal, toast, store, sha, timeAgo, icon } from "./ui.js";

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
    const v = { enabled: r.enabled, results: r.results || {}, links: r.links || {}, at: Date.now() };
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
  app.replaceChildren(head, el("main", { class: cls }, ...body));
  window.scrollTo(0, 0);
}
/** Tabs shown on every page of a day, so students can move between its parts in one tap. */
function dayTabs(n, active) {
  const tabs = [["day", "Activities", `#/day/${n}`, "list"], ["lesson", "Lesson", `#/lesson/${n}`, "book"],
    ["sheet", "Workbook", `#/sheet/${n}`, "grid"], ["quiz", "Quiz", `#/quiz/${n}`, "quiz"], ["files", "Files", `#/files/${n}`, "download"]];
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
    const sig = (c) => JSON.stringify([c.enabled, c.results, c.links]);
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
      el("div", { class: "label" }, "Training days"),
      el("div", { class: "list" }, ...rows),
      el("p", { class: "small muted", style: "margin-top:10px" }, "Days 4 to 15 are added as the course progresses."));
  });
}

async function viewDay(n) {
  const day = await getDay(n);
  return withConfig(async (cfg) => {
    const tag = day.tag, res = cfg.results[tag] || {}, local = store.get(K.res(tag), {});
    const statusOf = (open, r, l) => {
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
        el("div", { class: "status" }, ...statusOf(open, r, local[a.id])),
        el("div", { class: "chev" }, open ? icon("chev") : icon("padlock")),
        r && r.feedback ? el("div", { class: "feedback" }, "Trainer feedback: " + r.feedback) : "");
    });
    const qOpen = isOpen(cfg, tag, "quiz"), qr = res.quiz;
    const head = el("div", { class: "listhead actsrow acts" }, el("div", {}, "No."), el("div", {}, "Activity"), el("div", {}, "Mode"),
      el("div", {}, "Time"), el("div", {}, "Points"), el("div", {}, "Status"), el("div", {}));
    shell(`Day ${n}`, "#/", dayTabs(n, "day"),
      el("div", { class: "pagehead" }, el("div", { class: "kicker" }, `Day ${n} · ${day.unit}`), el("h2", {}, day.title)),
      VIDEOS[n] ? el("div", { class: "toolrow" }, el("a", { class: "btn", href: `#/lesson/${n}/0` }, icon("play"), "Watch the demo video")) : "",
      el("div", { class: "label" }, "Activities"), head, el("div", { class: "list" }, ...rows),
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
function videoBox(url, title, sub, chapterList) {
  const src = embedUrl(url);
  const stage = el("div", { class: "vstage" });
  const play = el("button", { class: "vplay", type: "button", "aria-label": `Play: ${title}` },
    el("span", { class: "vtri" }), el("span", {}, "Play video"));
  stage.append(play);
  play.addEventListener("click", () => {
    if (!navigator.onLine) return toast("You are offline. Connect to the internet to watch the video.", "bad");
    document.querySelectorAll(".vstage iframe").forEach((f) => {                 // stop any other player
      const s = f.parentElement; f.remove(); s.classList.remove("on"); s.querySelector(".vplay").hidden = false;
    });
    play.hidden = true;
    stage.classList.add("on");
    stage.append(el("iframe", { src, title, allow: "autoplay; fullscreen", allowfullscreen: true, loading: "lazy", referrerpolicy: "no-referrer" }));
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
    videoBox(vid.url, `Day ${n} demonstration: ${day.title}`, "The whole demonstration. The lesson parts that have a VIDEO tag also play their own chapter.",
      el("table", { class: "vchapters" }, el("tbody", {}, ...vid.chapters.slice().sort((x, y) => toSec(x.start) - toSec(y.start)).map((c) =>
        el("tr", {}, el("td", { class: "mono" }, `${c.start}${c.end ? "–" + c.end : ""}`), el("td", {}, c.title), el("td", { class: "muted" }, c.part)))))) });
  for (const sec of day.lesson) {
    const ch = chapterFor(sec.h);
    items.push({ title: sec.h, video: !!ch, short: sec.h.replace(/^Part (\d+) — /, "$1 · "), body: () => [el("div", { class: "prose", html: sec.html }),
      ch ? videoBox(vid.url, ch.title, `This part of the demo starts at ${ch.start}` + (ch.end ? ` and ends at ${ch.end}` : "") +
        `. Press play, then drag the slider to ${ch.start}.`) : ""] });
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

// ------------------------------------------------------------------ activity runner
async function viewActivity(n, actId) {
  const day = await getDay(n);
  let cfg = await loadConfig();
  const tag = day.tag, act = day.activities.find((a) => a.id === actId);
  if (!act) return viewDay(n);
  if (!isOpen(cfg, tag, actId)) cfg = await loadConfig(true);      // the cache may be older than the trainer's switch
  if (!isOpen(cfg, tag, actId)) { toast("This activity is not open yet."); location.hash = `#/day/${n}`; return; }
  const allAns = store.get(K.ans(tag), {});
  const ans = (allAns[actId] = allAns[actId] || {});
  const results = store.get(K.res(tag), {});
  const needsSheet = act.sheet || act.parts.some((p) => ["cells", "spec", "textcells"].includes(p.type));
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
  const actFiles = (day.files || []).filter((f) => f.group !== "pdf" || /Activities/i.test(f.name));
  const filesBox = actFiles.length ? el("details", { class: "actfiles" }, el("summary", {}, icon("download"), el("span", {}, "Files for this activity"),
    el("span", { class: "small muted" }, `${actFiles.length} file${actFiles.length > 1 ? "s" : ""}`)), el("div", { class: "flist" }, ...actFiles.map(fileRow))) : "";
  const steps = el("div", { class: "task-sec steps" }, el("div", { class: "label" }, "Instructions"), filesBox,
    el("div", { class: "prose", html: act.html }),
    el("details", {}, el("summary", { class: "small", style: "cursor:pointer;margin-top:8px;color:var(--info)" }, "Excel keys and their lab equivalents"),
      el("table", { class: "equiv" }, el("thead", {}, el("tr", {}, el("th", {}, "In Excel"), el("th", {}, "In the lab"))),
        el("tbody", {}, ...EQUIV.map(([x, y]) => el("tr", {}, el("td", { class: "mono" }, x), el("td", {}, y)))))));
  const partEls = {};
  const form = el("div", { class: "parts" });
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
  const left = el("div", { class: "pane pane-task" }, steps, el("div", { class: "task-sec" }, el("div", { class: "label" }, "Answers"), form));
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
    for (const p of act.parts) {
      const r = score.parts[p.id], { box, fb } = partEls[p.id];
      box.classList.remove("good", "bad", "half");
      if (r.manual) { fb.textContent = "Your trainer will grade this."; continue; }
      if (p.type === "cells") {
        const d = r.detail, right = Object.values(d).filter((x) => x === 1).length;
        const noF = Object.values(d).filter((x) => x === -1).length, empty = Object.values(d).filter((x) => x === -2).length;
        fb.textContent = `${right} of ${p.cells.length} cells correct.` + (noF ? ` ${noF} cell(s) have a typed number: use a formula (start with =).` : "") +
          (empty ? ` ${empty} still empty.` : "") + (right < p.cells.length ? " Wrong cells are red on the sheet." : " Great work!");
        for (const [ref, v] of Object.entries(d)) if (v !== -2) marks[ref] = v === 1 ? "ok" : "bad";
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
    toast(`Auto-checked: ${round(score.auto)} of ${round(score.autoMax)} pts` + (score.manualMax ? ` (+${score.manualMax} ${score.manualMax === 1 ? "pt" : "pts"} graded by the trainer)` : ""));
    saveDraft(true);
  });
  btnSubmit.addEventListener("click", async () => {
    const { sub, score } = await evaluate();
    const manualEmpty = act.parts.filter((p) => p.manual && p.type === "text" && !(ans[p.id] || "").trim()).length;
    const msg = `Submit Activity ${act.id}?\n\nAuto-checked part: ${round(score.auto)} / ${round(score.autoMax)} pts` +
      (score.manualMax ? `\nTrainer-graded part: ${score.manualMax} pts` : "") + (manualEmpty ? `\n\nNote: ${manualEmpty} written answer(s) are still empty.` : "") +
      "\n\nYou can submit again later; the latest submission counts.";
    if (!confirm(msg)) return;
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
  const exportBtn = el("button", { class: "btn small", type: "button" }, icon("download"), "Export CSV");
  const settings = el("button", { class: "btn small icon", type: "button", title: "Board settings", "aria-label": "Board settings" }, icon("gear"));
  const logout = el("button", { class: "btn small", type: "button", title: "Log out of the trainer board on this browser" }, icon("logout"), "Log out");
  logout.addEventListener("click", trainerLogout);
  shell("Trainer board", "#/", el("div", { class: "boardbar" }, daySel, live, el("div", { class: "grow" }), videoBtn, linksBtn, exportBtn, settings, logout),
    toggles, table, legend, { __cls: "page work" });
  let data = null;
  settings.addEventListener("click", trainerSettings);
  daySel.addEventListener("change", () => { store.set("dal3_board_day", daySel.value); render(); });
  exportBtn.addEventListener("click", () => data && exportCsv(data, Number(daySel.value)));
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
    const thead = el("tr", {}, el("th", {}, "Student"), ...items.map((it) => el("th", { title: it.title }, it.id === "quiz" ? "Quiz" : it.id)), el("th", {}, "Total"), el("th", {}, "Last active"));
    const rows = data.roster.map((st) => {
      let total = 0, last = null;
      const tds = items.map((it) => {
        const s = sub[`${st.id}|${it.id}`], l = lv[`${st.id}|${it.id}`];
        if (l && (!last || l.updated > last)) last = l.updated;
        if (s && (!last || s.ts > last)) last = s.ts;
        let txt = "—", cls = "none", title = "";
        if (s) {
          const graded = !s.manualMax || (s.manual !== "" && s.manual != null);
          const pts = Number(s.auto) + Number(s.manual || 0), max = Number(s.autoMax) + Number(s.manualMax || 0);
          total += pts;
          txt = `${round(pts)}/${round(max)}${graded ? "" : "*"}`;
          cls = pts / Math.max(1, max) >= 0.8 ? "good" : pts / Math.max(1, max) >= 0.5 ? "half" : "bad";
          if (!graded && !Number(s.autoMax)) { txt = "TO GRADE"; cls = "pend"; }    // all-written activity: not a fail, just waiting
          if (s.late) txt += " LATE";
          title = `Submitted ${timeAgo(s.ts)} · attempt ${s.attempts}` + (graded ? "" : " · needs grading");
        } else if (l) {
          const fresh = Date.now() - new Date(l.updated).getTime() < 3 * 60000;
          txt = `${l.pct}%`;
          cls = "draft" + (fresh ? " fresh" : "");
          title = `Working · last saved ${timeAgo(l.updated)}`;
        }
        const td = el("td", { class: cls, title }, txt);
        if (s || l) { td.tabIndex = 0; td.addEventListener("click", () => openDetail(st, day, it, s, l)); }
        return td;
      });
      return el("tr", {}, el("th", { class: "sname" }, st.name, el("span", { class: "sid" }, st.id)), ...tds,
        el("td", { class: "total" }, round(total)), el("td", { class: "when" }, last ? timeAgo(last) : "—"));
    });
    table.replaceChildren(el("table", { class: "boardtable" }, el("thead", {}, thead), el("tbody", {}, ...rows)));
    legend.replaceChildren(el("span", { class: "lg-good" }, "80% or more"), el("span", { class: "lg-half" }, "50–79%"), el("span", { class: "lg-bad" }, "Below 50%"),
      el("span", { class: "lg-draft" }, "Working now: % of auto-checked points correct"), el("span", { class: "lg-pend" }, "* or TO GRADE: written answers need your grade"),
      el("span", {}, "LATE: submitted after you closed it"), el("span", {}, "Select a cell to open the student's work"));
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

async function openDetail(st, day, item, s, l) {
  const body = el("div", { class: "detail" }, el("p", { class: "muted" }, "Loading…"));
  const m = modal(`${st.name} · ${item.id === "quiz" ? "Quiz" : "Activity " + item.id}`, body, { wide: true });
  let d;
  try { d = await trainerCall("submission", { sid: st.id, day: day.tag, act: item.id }); }
  catch (e) { body.replaceChildren(el("p", { class: "bad" }, e.message)); return; }
  const rec = d.final || d.live;
  if (!rec) { body.replaceChildren(el("p", {}, "No work yet.")); return; }
  const payload = rec.payload || {};
  const act = day.activities.find((a) => a.id === item.id);
  const kids = [el("p", { class: "small muted" },
    (d.final ? `Final submission ${new Date(d.final.ts).toLocaleString()} · attempt ${d.final.attempts}` : "Not submitted yet") +
    (d.live ? ` · live work saved ${timeAgo(d.live.updated)}` : ""))];
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
        row.append(el("pre", { class: "answer" }, Array.isArray(a) ? a.join(", ") : (a ?? "(no answer)")));
      } else row.append(el("p", { class: "small" }, r.frac === 1 ? "Done." : "Not found."));
      kids.push(row);
    }
    // grading form
    const manual = act.parts.filter((p) => p.manual);
    if (manual.length && d.final) {
      const g = el("form", { class: "form grade" });
      g.innerHTML = `<h3>Grade the written answers (${manual.reduce((s, p) => s + p.points, 0)} pts)</h3>
        <label>Points for the written part <input name="pts" type="number" step="0.5" min="0" max="${manual.reduce((s, p) => s + p.points, 0)}" value="${d.final.manual ?? ""}"></label>
        <label>Feedback to the student <textarea name="fb" rows="2">${esc(d.final.feedback || "")}</textarea></label>
        <button class="btn primary" type="submit">Save grade</button>`;
      g.addEventListener("submit", async (e) => {
        e.preventDefault();
        try { await trainerCall("grade", { sid: st.id, day: day.tag, act: item.id, manual: Number(g.pts.value), feedback: g.fb.value }); toast("Grade saved.", "good"); m.close(); }
        catch (err) { toast(err.message, "bad"); }
      });
      kids.push(g);
    }
    if (payload.wb) {
      const host = el("div", { class: "sheethost ro" });
      const btn = el("button", { class: "btn", type: "button" }, "Show the student's workbook");
      btn.addEventListener("click", async () => {
        btn.disabled = true;
        btn.textContent = "Loading the workbook…";
        try { await loadEngine(); } catch (e) { toast(e.message, "bad"); btn.disabled = false; btn.textContent = "Show the student's workbook"; return; }
        btn.remove();
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
  let t;
  const d = (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  d.flush = () => { clearTimeout(t); fn(); };
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
    if (h[0] === "act") return await viewActivity(Number(h[1]), h[2]);
    if (h[0] === "quiz") return await viewQuiz(Number(h[1]));
    if (h[0] === "files") return await viewFiles(Number(h[1]));
    if (h[0] === "login") return await viewLogin();
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
// Offline support + clean updates. When a new version takes over: reload on lists/lessons; inside an activity or a
// workbook, ask first so a student's typing is never interrupted.
if ("serviceWorker" in navigator && location.protocol !== "file:") {
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).catch(() => {});
  let reloading = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!hadController || reloading) return;
    const working = /^#\/(act|sheet|quiz)\//.test(location.hash);
    if (!working) { reloading = true; location.reload(); return; }
    const bar = el("div", { class: "updatebar", role: "status" }, "A new version of the lab is ready.",
      el("button", { class: "btn small primary", type: "button", onclick: () => { reloading = true; location.reload(); } }, "Reload"));
    document.body.append(bar);
  });
}
