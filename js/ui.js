// Tiny DOM helpers: el(), esc(), modal(), toast(), storage.
export function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
export function el(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === false || v == null) continue;
    if (k === "class") n.className = v;
    else if (k === "html") n.innerHTML = v;
    else if (k.startsWith("on") && typeof v === "function") n.addEventListener(k.slice(2), v);
    else if (k in n && typeof v !== "string") n[k] = v;
    else n.setAttribute(k, v === true ? "" : v);
  }
  for (const k of kids.flat()) if (k !== null && k !== undefined && k !== "") n.append(k instanceof Node ? k : document.createTextNode(String(k)));
  return n;
}
// Line icons, 16 x 16, drawn with currentColor (no emoji anywhere in the UI).
const P = {
  db: '<ellipse cx="8" cy="4" rx="5" ry="1.8"/><path d="M3 4v8c0 1 2.2 1.8 5 1.8s5-.8 5-1.8V4M3 8c0 1 2.2 1.8 5 1.8s5-.8 5-1.8"/>',
  back: '<path d="M10 3 5 8l5 5"/>',
  close: '<path d="m4 4 8 8M12 4l-8 8"/>',
  check: '<path d="m3 8.5 3 3 7-7"/>',
  undo: '<path d="M6 4 3 7l3 3"/><path d="M3 7h6.5a3.5 3.5 0 0 1 0 7H8"/>',
  copy: '<rect x="5.5" y="5.5" width="8" height="8"/><path d="M3.5 10.5h-1v-8h8v1"/>',
  paste: '<rect x="3.5" y="3.5" width="9" height="10"/><path d="M6 3.5V2h4v1.5M6 7h4M6 10h4"/>',
  clear: '<path d="m3 13 7.5-9 3 2.5L7 13H3zM9 5.5l3 2.5"/>',
  filldown: '<rect x="3.5" y="2.5" width="9" height="3"/><path d="M8 7v6m-2.5-2.5L8 13l2.5-2.5"/>',
  fillright: '<rect x="2.5" y="3.5" width="3" height="9"/><path d="M7 8h6m-2.5-2.5L13 8l-2.5 2.5"/>',
  fillend: '<rect x="3.5" y="2.5" width="9" height="3"/><path d="M8 7v5.5M5 13.5h6m-5-3L8 12.5l2-2"/>',
  select: '<path d="M2.5 5V2.5H5M11 2.5h2.5V5M13.5 11v2.5H11M5 13.5H2.5V11"/><rect x="6" y="6" width="4" height="4"/>',
  format: '<path d="M3 13h10M5 10V3h3a2 2 0 0 1 0 4H5m3 0 2.5 3"/>',
  highlight: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M2.5 6.5h11M2.5 10h11M6.5 2.5v11"/><rect x="7" y="7" width="6" height="2.5" fill="currentColor" stroke="none"/>',
  sum: '<path d="M12 3H4l4.5 5L4 13h8"/>',
  fx: '<path d="M7 13c1 0 1.5-.7 1.7-1.8L10 3.8c.2-1 .8-1.3 1.6-1.3M6.5 6.5h5"/><path d="m10.5 9 3 4m0-4-3 4"/>',
  lock: '<rect x="3.5" y="7" width="9" height="6.5"/><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2"/>',
  analysis: '<path d="M2.5 13.5h11"/><path d="M3.5 11c2-6 3-6 4.5-1s2.5 2 4.5-5"/>',
  chart: '<path d="M2.5 13.5h11M4.5 13V8M7.5 13V4M10.5 13V6.5"/>',
  book: '<path d="M8 4.5C6.5 3 4.5 3 2.5 3.5v9c2-.5 4-.5 5.5 1 1.5-1.5 3.5-1.5 5.5-1v-9C11.5 3 9.5 3 8 4.5zm0 0v9"/>',
  grid: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M2.5 6.5h11M2.5 10h11M6 2.5v11M10 2.5v11"/>',
  play: '<rect x="2.5" y="3.5" width="11" height="9"/><path d="m7 6 3 2-3 2z"/>',
  padlock: '<rect x="4" y="7.5" width="8" height="6"/><path d="M5.5 7.5V5.5a2.5 2.5 0 0 1 5 0v2"/>',
  download: '<path d="M8 2.5v8m-3-3 3 3 3-3M3 13.5h10"/>',
  gear: '<circle cx="8" cy="8" r="2"/><path d="M8 1.5v2M8 12.5v2M1.5 8h2M12.5 8h2M3.4 3.4l1.4 1.4M11.2 11.2l1.4 1.4M3.4 12.6l1.4-1.4M11.2 4.8l1.4-1.4"/>',
  chev: '<path d="m6 3 5 5-5 5"/>',
  list: '<path d="M5.5 4h8M5.5 8h8M5.5 12h8M2.5 4h1M2.5 8h1M2.5 12h1"/>',
  logout: '<path d="M6.5 2.5h-4v11h4M10 5l3 3-3 3M13 8H6"/>',
  quiz: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="m5 6 1 1 2-2M5 10.5l1 1 2-2M9.5 6.5h2M9.5 11h2"/>',
};
export function icon(name, cls = "ic") {
  const s = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  s.setAttribute("viewBox", "0 0 16 16");
  s.setAttribute("class", cls);
  s.setAttribute("aria-hidden", "true");
  s.innerHTML = `<g fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="square" stroke-linejoin="miter">${P[name] || ""}</g>`;
  return s;
}

export function modal(title, content, { wide = false } = {}) {
  const back = el("div", { class: "modal-back" });
  const close = () => { back.remove(); document.removeEventListener("keydown", onKey); };
  const onKey = (e) => { if (e.key === "Escape") close(); };
  const x = el("button", { class: "btn icon", type: "button", "aria-label": "Close" }, icon("close"));
  x.addEventListener("click", close);
  const box = el("div", { class: "modal" + (wide ? " wide" : ""), role: "dialog", "aria-modal": "true", "aria-label": title },
    el("div", { class: "modal-head" }, el("h3", {}, title), x), el("div", { class: "modal-body" }, content));
  box.querySelectorAll("[data-cancel]").forEach((c) => c.addEventListener("click", close));
  back.append(box);
  back.addEventListener("pointerdown", (e) => { if (e.target === back) close(); });
  document.addEventListener("keydown", onKey);
  document.body.append(back);
  setTimeout(() => box.querySelector("input,select,textarea,button.btn.primary")?.focus(), 30);
  return { close, box };
}
let toastTimer;
export function toast(msg, kind = "") {
  let t = document.getElementById("toast");
  if (!t) { t = el("div", { id: "toast", role: "status", "aria-live": "polite" }); document.body.append(t); }
  t.className = "show " + kind;
  t.textContent = msg;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.className = ""), kind === "bad" ? 5200 : 3200);
}
export const store = {
  get(k, d = null) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem(k); } catch { /* ignore */ } },
};
export async function sha(s) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join("").slice(0, 20);
}
export const timeAgo = (t) => {
  if (!t) return "";
  const s = Math.max(0, (Date.now() - new Date(t).getTime()) / 1000);
  if (s < 60) return "just now";
  if (s < 3600) return Math.round(s / 60) + " min ago";
  if (s < 86400) return Math.round(s / 3600) + " h ago";
  return new Date(t).toLocaleDateString();
};
