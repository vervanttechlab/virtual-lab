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

export async function call(action, body = {}, { timeout = 25000, auth = true } = {}) {
  const url = apiUrl();
  if (!url) throw new NetError("No server configured (demo mode).");
  const s = session.get();
  const payload = { action, ...body };
  if (auth && s) Object.assign(payload, { id: s.id, token: s.token });
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
export const trainerCall = (action, body = {}) => call(action, { ...body, key: trainer.get() }, { auth: false });

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
          const r = await call(it.action, it.body);
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
  try {
    const data = await call(action, body);
    return { sent: true, data };
  } catch (e) {
    if (e instanceof NetError) {
      queue.add({ action, body, replaceKey: opts.replaceKey, onDoneKey: opts.onDoneKey });
      return { sent: false, queued: true };
    }
    throw e;
  }
}
