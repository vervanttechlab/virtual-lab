// Backend client (Google Apps Script web app) with an offline queue.
// POST with a text/plain body avoids the CORS preflight that Apps Script cannot answer.
import { CONFIG } from "./config.js";
import { store } from "./ui.js";

export const apiUrl = () => store.get("dal3_api_override") || CONFIG.API_URL;
export const demoMode = () => !apiUrl();
export const session = {
  get: () => store.get("dal3_session"),
  set: (s) => store.set("dal3_session", s),
  clear: () => store.del("dal3_session"),
};
export const trainer = {
  get: () => store.get("dal3_trainer"),
  set: (k) => store.set("dal3_trainer", k),
  clear: () => store.del("dal3_trainer"),
};

export class NetError extends Error {}

// Busy indicator: counts foreground server calls so the UI can show a loading bar (background calls pass quiet: true).
let busy = 0;
const busySubs = new Set();
export const onBusy = (fn) => { busySubs.add(fn); fn(busy > 0); return () => busySubs.delete(fn); };
const setBusy = (d) => { busy = Math.max(0, busy + d); busySubs.forEach((fn) => fn(busy > 0)); };

export async function call(action, body = {}, { timeout = 25000, auth = true, quiet = false } = {}) {
  if (!quiet) setBusy(1);
  try { return await call_(action, body, { timeout, auth }); } finally { if (!quiet) setBusy(-1); }
}
const newNonce = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// A flaky DNS / Wi-Fi often loses only the REPLY: Google already ran the action. Each request therefore carries a one-time
// id (nonce); the server remembers the reply for that id, so retrying is safe (no double submissions or double switches).
async function call_(action, body, { timeout, auth }) {
  const url = apiUrl();
  if (!url) throw new NetError("No server configured (demo mode).");
  const s = session.get();
  const payload = { action, ...body };
  if (!payload.nonce) payload.nonce = newNonce();
  if (auth && s) Object.assign(payload, { id: s.id, token: s.token });
  for (let attempt = 0; ; attempt++) {
    try {
      return await send_(url, payload, timeout);
    } catch (e) {
      if (!(e instanceof NetError) || !navigator.onLine || attempt >= 2) throw e;
      await sleep(attempt ? 2500 : 900);                 // try again: 0.9 s, then 2.5 s
    }
  }
}
async function send_(url, payload, timeout) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), timeout);
  let res;
  try {
    res = await fetch(url, { method: "POST", body: JSON.stringify(payload), headers: { "Content-Type": "text/plain;charset=utf-8" },
      redirect: "follow", signal: ctl.signal });
  } catch (e) {
    throw new NetError(navigator.onLine ? "Cannot reach the server." : "You are offline.");
  } finally {
    clearTimeout(t);
  }
  let data;
  try { data = await res.json(); } catch { throw new NetError("The server sent an unexpected reply."); }
  if (!data.ok) throw new Error(data.error || "Server error");
  return data;
}
export const trainerCall = (action, body = {}, opt = {}) => call(action, { ...body, key: trainer.get() }, { auth: false, ...opt });

// ------------------------------------------------------------------ offline queue
const QK = "dal3_queue";
export const queue = {
  list: () => store.get(QK, []),
  add(item) {
    const q = queue.list().filter((x) => !(item.replaceKey && x.replaceKey === item.replaceKey));
    q.push({ ...item, queuedAt: new Date().toISOString() });
    store.set(QK, q);
    notify();
  },
  async flush() {
    if (demoMode() || queue.busy) return;
    queue.busy = true;
    try {
      let q = queue.list();
      while (q.length) {
        const it = q[0];
        try {
          const r = await call(it.action, it.body, { quiet: true });
          if (it.onDoneKey) store.set(it.onDoneKey, { ...(store.get(it.onDoneKey) || {}), server: r, pending: false });
        } catch (e) {
          if (e instanceof NetError) break;           // still offline: try later
          console.warn("queue item rejected", it, e);  // server refused: drop it so the queue does not jam
          store.set("dal3_queue_errors", [...store.get("dal3_queue_errors", []), { it, error: e.message }].slice(-20));
        }
        q = queue.list().slice(1);
        store.set(QK, q);
      }
    } finally {
      queue.busy = false;
      notify();
    }
  },
};
const subs = new Set();
export const onQueue = (fn) => { subs.add(fn); return () => subs.delete(fn); };
function notify() { subs.forEach((fn) => fn(queue.list().length)); }
window.addEventListener("online", () => queue.flush());
setInterval(() => queue.flush(), 30000);

/** Send now; if offline, queue it. Returns {sent:true, data} or {sent:false}. */
export async function sendOrQueue(action, body, opts = {}) {
  if (demoMode()) return { sent: false, demo: true };
  body = { ...body, nonce: body.nonce || newNonce() };   // same id if it is queued and sent again later
  try {
    const data = await call(action, body, { quiet: !!opts.quiet });
    return { sent: true, data };
  } catch (e) {
    if (e instanceof NetError) {
      queue.add({ action, body, replaceKey: opts.replaceKey, onDoneKey: opts.onDoneKey });
      return { sent: false, queued: true };
    }
    throw e;
  }
}
