// DAL3 Virtual Lab: router and views (student + trainer).
import { CONFIG } from "./config.js";
import { call, trainerCall, sendOrQueue, session, trainer, demoMode, queue, onQueue, apiUrl, NetError } from "./api.js";
import { Workbook } from "./workbook.js";
import { SheetUI } from "./grid.js";
import { el, esc, modal, toast, store, sha, timeAgo } from "./ui.js";

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
    const v = { enabled: r.enabled, results: r.results || {}, at: Date.now() };
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
  const s = session.get();
  const qn = queue.list().length;
  const pend = el("span", { class: "pending" + (qn ? " show" : ""), title: "Work waiting to be sent (offline)" }, `☁ ${qn} waiting`);
  onLeave(onQueue((n) => { pend.textContent = `☁ ${n} waiting`; pend.classList.toggle("show", n > 0); }));
  const head = el("header", { class: "appbar" },
    back ? el("a", { class: "back", href: back, "aria-label": "Back" }, "‹") : el("span", { class: "logo" }, "DAL3"),
    el("div", { class: "title" }, title),
    pend,
    demoMode() ? el("span", { class: "badge warn", title: "No server configured: work stays on this device" }, "DEMO") : "",
    s ? el("button", { class: "user", type: "button", onclick: () => accountMenu() }, (s.name || s.id).split(" ")[0]) : "");
  app.replaceChildren(head, el("main", { class: "page" }, ...body));
  window.scrollTo(0, 0);
}
function accountMenu() {
  const s = session.get();
  const box = el("div", { class: "stack" },
    el("p", {}, `Signed in as ${s.name || s.id} (${s.id})`),
    el("button", { class: "btn wide", type: "button", onclick: () => { m.close(); location.hash = "#/trainer"; } }, "Trainer board"),
    el("button", { class: "btn wide danger", type: "button", onclick: () => {
      if (queue.list().length && !confirm("Some work is still waiting to be sent. Sign out anyway? It will be sent after you sign in again on this phone.")) return;
      session.clear(); m.close(); location.hash = "#/login";
    } }, "Sign out"));
  const m = modal("Account", box);
}

// ------------------------------------------------------------------ views
async function viewLogin() {
  const f = el("form", { class: "card form login" });
  f.innerHTML = `
    <h2>Welcome to the DAL3 Virtual Lab</h2>
    <p class="muted">TESDA Data Analytics Level III · Tindahan ni Aling Nena</p>
    ${demoMode() ? `<p class="note">Demo mode: no class server is set up yet. Your work is saved on this device only.</p>
      <label>Your name <input name="name" required autocomplete="name"></label>` : `
      <label>Class code <input name="cls" required value="${esc(CONFIG.CLASS_CODE_HINT)}" autocapitalize="characters"></label>
      <label>Student ID <input name="sid" required autocapitalize="characters" autocomplete="username"></label>
      <label>PIN <input name="pin" required inputmode="numeric" type="password" autocomplete="current-password" maxlength="8"></label>`}
    <button class="btn primary wide" type="submit">Sign in</button>
    <p class="muted small">Trainer? <a href="#/trainer">Open the trainer board</a></p>`;
  shell("Sign in", null, f);
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
      queue.flush();
      location.hash = "#/";
    } catch (err) {
      toast(err.message, "bad");
      btn.disabled = false;
      btn.textContent = "Sign in";
    }
  });
}

async function viewHome() {
  const [idx, cfg] = await Promise.all([getIndex(), loadConfig(true)]);
  const list = el("div", { class: "cards" });
  for (const d of idx.days) {
    const open = d.activities.filter((a) => isOpen(cfg, d.tag, a.id)).length;
    const res = cfg.results[d.tag] || {};
    const done = Object.keys(res).filter((k) => k !== "quiz").length;
    list.append(el("a", { class: "card daycard", href: `#/day/${d.day}` },
      el("div", { class: "daynum" }, `Day ${d.day}`),
      el("div", {}, el("h3", {}, d.title), el("p", { class: "muted small" }, d.unit),
        el("p", { class: "small" }, cfg.enabled ? `${open} of ${d.activities.length} activities open · ${done} submitted` : `${d.activities.length} activities`))));
  }
  shell("DAL3 Virtual Lab", null,
    el("section", { class: "hello" }, el("h2", {}, `Hi, ${(session.get().name || "").split(" ")[0] || "there"}!`),
      el("p", { class: "muted" }, cfg.offline ? "You are offline. You can keep working; your answers are sent when you are back online." :
        "Choose a day. Activities your trainer has opened are ready to do.")),
    list,
    el("p", { class: "muted small center" }, "Days 4–15 will appear here as the course goes on."));
}

async function viewDay(n) {
  const [day, cfg] = await Promise.all([getDay(n), loadConfig(true)]);
  const tag = day.tag, res = cfg.results[tag] || {}, local = store.get(K.res(tag), {});
  const rows = day.activities.map((a) => {
    const open = isOpen(cfg, tag, a.id);
    const r = res[a.id], l = local[a.id];
    let status = "Not started";
    if (r) status = `Submitted · ${fmtScore(r)}`;
    else if (l && l.pending) status = "Submitted · waiting to send";
    else if (l && l.checked) status = `In progress · ${Math.round(100 * l.checked.auto / Math.max(1, l.checked.autoMax))}% correct so far`;
    return el(open ? "a" : "div", { class: "card actrow" + (open ? "" : " locked"), href: open ? `#/act/${n}/${a.id}` : null },
      el("div", { class: "actid" }, a.id),
      el("div", { class: "grow" }, el("h3", {}, a.title), el("p", { class: "muted small" }, `${a.mode} · ${a.minutes} min · ${a.points} pts`),
        el("p", { class: "small status" }, open ? status : "🔒 Not open yet: your trainer will open it"),
        r && r.feedback ? el("p", { class: "feedback" }, "Trainer: " + r.feedback) : ""),
      el("div", { class: "chev" }, open ? "›" : ""));
  });
  const qOpen = isOpen(cfg, tag, "quiz");
  const qr = res.quiz;
  shell(`Day ${n}`, "#/",
    el("section", { class: "dayhead" }, el("p", { class: "muted small" }, day.unit), el("h2", {}, day.title)),
    el("div", { class: "quick" },
      el("a", { class: "btn", href: `#/lesson/${n}` }, "📖 Lesson"),
      el("a", { class: "btn", href: `#/sheet/${n}` }, "🧪 Practice sheet"),
      CONFIG.VIDEOS[n] ? el("a", { class: "btn", href: CONFIG.VIDEOS[n], target: "_blank", rel: "noopener" }, "▶ Demo video") : ""),
    el("h3", { class: "sec" }, "Activities"), ...rows,
    el("h3", { class: "sec" }, "Self-check"),
    el(qOpen ? "a" : "div", { class: "card actrow" + (qOpen ? "" : " locked"), href: qOpen ? `#/quiz/${n}` : null },
      el("div", { class: "actid" }, "Q"),
      el("div", { class: "grow" }, el("h3", {}, day.quiz.title), el("p", { class: "small status" },
        !qOpen ? "🔒 Not open yet" : qr ? `Last score ${qr.auto}/${qr.autoMax}` : `${day.quiz.items.length} items · pass 80%`)),
      el("div", { class: "chev" }, qOpen ? "›" : "")));
}
const fmtScore = (r) => {
  const graded = r.manualMax ? (r.manual !== "" && r.manual != null) : true;
  const total = Number(r.auto || 0) + Number(r.manual || 0);
  return graded ? `${round(total)}/${round(Number(r.autoMax) + Number(r.manualMax || 0))}` :
    `${round(r.auto)}/${round(r.autoMax)} auto · ${r.manualMax} pts waiting for trainer`;
};
const round = (x) => Math.round(Number(x) * 10) / 10;

async function viewLesson(n) {
  const day = await getDay(n);
  const secs = day.lesson.map((s, i) => {
    const d = el("details", { class: "card lesson" }, el("summary", {}, s.h), el("div", { class: "prose", html: s.html }));
    if (i === 0) d.open = true;
    return d;
  });
  shell(`Day ${n} · Lesson`, `#/day/${n}`, el("h2", {}, day.title), ...secs);
}

async function viewSheet(n) {
  const day = await getDay(n);
  const host = el("div", { class: "sheethost full" });
  shell(`Day ${n} · Practice sheet`, `#/day/${n}`,
    el("p", { class: "muted small" }, "This is your Day " + n + " workbook. Activities use the same workbook, so your formulas here count too."),
    host);
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
  const wb = needsSheet ? new Workbook(day, store.get(K.wb(tag))) : null;
  let dirty = false, ui = null;
  const saveLocal = debounce(() => {
    if (wb) store.set(K.wb(tag), wb.serialize());
    store.set(K.ans(tag), allAns);
  }, 500);
  const touch = () => { dirty = true; saveLocal(); };

  // ---- left: steps + answers
  const steps = el("details", { class: "card steps", open: true }, el("summary", {}, "Steps from the activity sheet"),
    el("div", { class: "prose", html: act.html }),
    el("div", { class: "tips" }, el("b", {}, "In the lab app"), el("ul", {}, ...day.labTips.map((t) => el("li", {}, t)))));
  const partEls = {};
  const form = el("div", { class: "parts" });
  for (const p of act.parts) {
    const box = el("div", { class: "card part", "data-part": p.id });
    const head = el("div", { class: "phead" }, el("span", { class: "pprompt" }, p.prompt || p.label || ""),
      el("span", { class: "ppts" }, `${p.points} pt${p.points === 1 ? "" : "s"}${p.manual ? " · trainer" : ""}`));
    const fb = el("div", { class: "pfb" });
    box.append(head);
    if (p.type === "cells") {
      const refs = summarizeRefs(p.cells);
      box.append(el("p", { class: "small muted" }, `${p.cells.length} cells: ${refs}`));
      const go = el("button", { class: "btn small", type: "button" }, "Go to the sheet");
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
      box.append(el("p", { class: "small muted" }, { chart: "Use the 📊 Chart button.", toolpak: "Use the 🧮 Analysis button.", highlight: "Use the 🖍 Highlight button." }[p.check] || ""));
    }
    box.append(fb);
    partEls[p.id] = { box, fb };
    form.append(box);
  }
  const left = el("div", { class: "pane pane-steps" }, steps, el("h3", { class: "sec" }, "Your answers"), form);
  const right = needsSheet ? el("div", { class: "pane pane-sheet" }, el("div", { class: "sheethost" })) : null;

  // ---- bottom bar
  const status = el("span", { class: "small muted" }, results[actId]?.pending ? "Submitted · waiting to send" :
    results[actId]?.server ? `Submitted · ${fmtScore(results[actId].server)}` : "");
  const btnCheck = el("button", { class: "btn", type: "button" }, "✓ Check");
  const btnSubmit = el("button", { class: "btn primary", type: "button" }, "Submit");
  const bar = el("div", { class: "actbar" }, status, el("div", { class: "grow" }), btnCheck, btnSubmit);

  const seg = el("div", { class: "seg", role: "tablist" });
  const panes = { steps: left, sheet: right };
  const showPane = (k) => {
    if (!panes[k]) return;
    seg.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.k === k));
    Object.entries(panes).forEach(([kk, p]) => p && p.classList.toggle("show", kk === k));
  };
  if (right) {
    for (const [k, label] of [["steps", "Steps & answers"], ["sheet", "Sheet"]]) {
      const b = el("button", { type: "button", "data-k": k }, label);
      b.addEventListener("click", () => showPane(k));
      seg.append(b);
    }
  }
  shell(`Activity ${act.id}`, `#/day/${n}`,
    el("div", { class: "acthead" }, el("h2", {}, act.title), el("p", { class: "muted small" }, `${act.mode} · ${act.minutes} min · ${act.points} pts`)),
    right ? seg : "", el("div", { class: "actgrid" + (right ? "" : " single") }, left, right || ""), bar);
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
        fb.textContent = r.frac === 1 ? "✓ Correct" : p.type === "spec" ? "✗ Not found yet. Check the range and settings." : "✗ Not yet. Try again.";
      }
      box.classList.add(r.frac === 1 ? "good" : r.frac > 0 ? "half" : "bad");
    }
    ui && ui.setMarks(marks);
  }
  btnCheck.addEventListener("click", async () => {
    const { score } = await evaluate();
    showFeedback(score);
    results[actId] = { ...(results[actId] || {}), checked: { auto: score.auto, autoMax: score.autoMax } };
    store.set(K.res(tag), results);
    toast(`Auto-checked: ${round(score.auto)} of ${round(score.autoMax)} pts` + (score.manualMax ? ` (+${score.manualMax} pts graded by the trainer)` : ""));
    saveDraft(true);
  });
  btnSubmit.addEventListener("click", async () => {
    const { sub, score } = await evaluate();
    const manualEmpty = act.parts.filter((p) => p.manual && p.type === "text" && !(ans[p.id] || "").trim()).length;
    const msg = `Submit Activity ${act.id}?\n\nAuto-checked part: ${round(score.auto)} / ${round(score.autoMax)} pts` +
      (score.manualMax ? `\nTrainer-graded part: ${score.manualMax} pts` : "") + (manualEmpty ? `\n\n⚠ ${manualEmpty} written answer(s) are still empty.` : "") +
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
    try {
      const r = await sendOrQueue("submit", body, { onDoneKey: null });
      if (r.sent) {
        results[actId] = { ...(results[actId] || {}), server: r.data.result, pending: false };
        status.textContent = `Submitted · ${fmtScore(r.data.result)}`;
        toast("Submitted! Your trainer can see your work now.", "good");
      } else {
        results[actId] = { ...(results[actId] || {}), pending: true };
        status.textContent = "Submitted · waiting to send";
        toast("You are offline. Your work is saved and will be sent automatically.");
      }
      store.set(K.res(tag), results);
      dirty = false;
    } catch (e) {
      toast(e.message, "bad");
    } finally {
      btnSubmit.disabled = false;
    }
  });

  // ---- live autosave to the trainer board
  async function saveDraft(force) {
    if (demoMode() || (!dirty && !force)) return;
    dirty = false;
    const { sub, score } = await evaluate();
    sendOrQueue("save", { day: tag, act: act.id, final: false, sub, wb: wb ? wb.serialize() : null,
      client: { auto: score.auto, autoMax: score.autoMax } }, { replaceKey: `save|${tag}|${act.id}` }).catch(() => {});
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
    const box = el("fieldset", { class: "card qitem", "data-n": it.n }, el("legend", {}, `${it.n}. ${it.q}`));
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
  const result = el("div", { class: "card result", hidden: true });
  const submit = el("button", { class: "btn primary wide", type: "submit" }, "Submit quiz");
  form.append(submit);
  shell(`Day ${n} · Quiz`, `#/day/${n}`, el("h2", {}, day.quiz.title), el("p", { class: "muted small" }, `${day.quiz.items.length} items · Passing score 80%`), result, form);
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
    result.className = "card result " + (pass ? "good" : "bad");
    result.innerHTML = `<h3>Score: ${r.score} / ${r.max} ${pass ? "· Passed 🎉" : "· Not yet: review the lesson and try again"}</h3>`;
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
  const live = el("span", { class: "small muted" }, "");
  const toggles = el("div", { class: "toggles" });
  const table = el("div", { class: "board" });
  const exportBtn = el("button", { class: "btn small", type: "button" }, "⬇ CSV");
  const settings = el("button", { class: "btn small", type: "button" }, "⚙");
  shell("Trainer board", "#/", el("div", { class: "boardbar" }, daySel, live, el("div", { class: "grow" }), exportBtn, settings), toggles, table);
  let data = null;
  settings.addEventListener("click", trainerSettings);
  daySel.addEventListener("change", () => { store.set("dal3_board_day", daySel.value); render(); });
  exportBtn.addEventListener("click", () => data && exportCsv(data, Number(daySel.value)));

  async function refresh() {
    try {
      data = await trainerCall("progress");
      live.textContent = `● live · updated ${new Date().toLocaleTimeString()}`;
      live.className = "small live";
      render();
    } catch (e) {
      live.textContent = "⚠ " + e.message;
      live.className = "small bad";
      if (/trainer key/i.test(e.message)) { trainer.clear(); trainerLogin(); }
    }
  }
  async function render() {
    if (!data) return;
    const day = await getDay(Number(daySel.value)), tag = day.tag;
    const items = [...day.activities.map((a) => ({ id: a.id, title: a.title, points: a.points })), { id: "quiz", title: "Quiz", points: day.quiz.items.length }];
    const en = (data.config[tag] || {});
    toggles.replaceChildren(el("b", {}, "Open for students: "), ...items.map((it) => {
      const b = el("button", { class: "toggle" + (en[it.id] ? " on" : ""), type: "button", title: it.title }, `${it.id === "quiz" ? "Quiz" : it.id} ${en[it.id] ? "ON" : "OFF"}`);
      b.addEventListener("click", async () => {
        b.disabled = true;
        try { await trainerCall("setConfig", { day: tag, item: it.id, enabled: !en[it.id] }); await refresh(); }
        catch (e) { toast(e.message, "bad"); b.disabled = false; }
      });
      return b;
    }), el("button", { class: "btn small", type: "button", onclick: async () => { await trainerCall("setConfig", { day: tag, item: "*", enabled: true, items: items.map((i) => i.id) }); refresh(); } }, "Open all"),
    el("button", { class: "btn small", type: "button", onclick: async () => { await trainerCall("setConfig", { day: tag, item: "*", enabled: false, items: items.map((i) => i.id) }); refresh(); } }, "Close all"));

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
          txt = `${round(pts)}/${round(max)}${graded ? "" : " ⏳"}`;
          cls = pts / Math.max(1, max) >= 0.8 ? "good" : pts / Math.max(1, max) >= 0.5 ? "half" : "bad";
          if (!graded && !Number(s.autoMax)) { txt = "⏳ grade"; cls = "pend"; }       // all-written activity: not a fail, just waiting
          if (s.late) txt += " ⏰";
          title = `Submitted ${timeAgo(s.ts)} · attempt ${s.attempts}` + (graded ? "" : " · needs grading");
        } else if (l) {
          const fresh = Date.now() - new Date(l.updated).getTime() < 3 * 60000;
          txt = `✎ ${l.pct}%`;
          cls = "draft" + (fresh ? " fresh" : "");
          title = `Working · last saved ${timeAgo(l.updated)}`;
        }
        const td = el("td", { class: cls, title }, txt);
        if (s || l) { td.tabIndex = 0; td.addEventListener("click", () => openDetail(st, day, it, s, l)); }
        return td;
      });
      return el("tr", {}, el("th", { class: "sname" }, el("div", {}, st.name), el("div", { class: "small muted" }, st.id)), ...tds,
        el("td", { class: "total" }, round(total)), el("td", { class: "small muted" }, last ? timeAgo(last) : "—"));
    });
    table.replaceChildren(el("table", { class: "boardtable" }, el("thead", {}, thead), el("tbody", {}, ...rows)),
      el("p", { class: "small muted" }, "✎ = working now (live progress = % of auto-checked points correct) · ⏳ = written answers waiting for your grade · ⏰ = submitted while the activity was closed · tap a cell to see the student's work."));
  }
  refresh();
  const t = setInterval(refresh, CONFIG.BOARD_REFRESH_SECONDS * 1000);
  onLeave(() => clearInterval(t));
}
const indexBy = (arr, f) => Object.fromEntries(arr.map((x) => [f(x), x]));

function trainerLogin() {
  const f = el("form", { class: "card form login" });
  f.innerHTML = `<h2>Trainer board</h2>
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
    el("button", { class: "btn wide danger", type: "button", onclick: () => { trainer.clear(); m.close(); location.hash = "#/"; } }, "Lock the trainer board"));
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
      `${(payload.answers || {})[it.n] ?? "—"}  ${res[it.n] ? "✓" : "✗"}`))));
  } else {
    const det = (d.final && d.final.detail) || (d.live && d.live.detail) || {};
    const parts = (det.parts) || {};
    for (const p of act.parts) {
      const r = parts[p.id] || {};
      const a = (payload.sub && payload.sub.answers || {})[p.id];
      const row = el("div", { class: "card part " + (r.manual ? "" : r.frac === 1 ? "good" : r.frac > 0 ? "half" : "bad") },
        el("div", { class: "phead" }, el("span", { class: "pprompt" }, p.prompt || p.label), el("span", { class: "ppts" },
          r.manual ? `${p.points} pts · trainer` : `${r.pts ?? 0}/${p.points}`)));
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
      } else row.append(el("p", { class: "small" }, r.frac === 1 ? "✓ done" : "✗ not found"));
      kids.push(row);
    }
    // grading form
    const manual = act.parts.filter((p) => p.manual);
    if (manual.length && d.final) {
      const g = el("form", { class: "card form grade" });
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
      btn.addEventListener("click", () => {
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
async function route() {
  cleanup.forEach((fn) => { try { fn(); } catch (e) { console.warn(e); } });
  cleanup = [];
  document.querySelectorAll(".modal-back").forEach((m) => m.remove());
  const h = location.hash.replace(/^#\/?/, "").split("/");
  try {
    if (h[0] === "trainer") return await viewTrainer();
    if (!session.get()) return await viewLogin();
    if (h[0] === "day") return await viewDay(Number(h[1]));
    if (h[0] === "lesson") return await viewLesson(Number(h[1]));
    if (h[0] === "sheet") return await viewSheet(Number(h[1]));
    if (h[0] === "act") return await viewActivity(Number(h[1]), h[2]);
    if (h[0] === "quiz") return await viewQuiz(Number(h[1]));
    if (h[0] === "login") return await viewLogin();
    return await viewHome();
  } catch (e) {
    console.error(e);
    shell("Something went wrong", "#/", el("div", { class: "card" }, el("p", {}, e.message), el("a", { class: "btn", href: "#/" }, "Home")));
  }
}
window.addEventListener("hashchange", route);
window.addEventListener("beforeunload", () => cleanup.forEach((fn) => { try { fn(); } catch { /* ignore */ } }));
route();
queue.flush();
if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js").catch(() => {});
