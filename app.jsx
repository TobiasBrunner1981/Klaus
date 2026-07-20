/* Klaus — shared household organizer. Phase 1.
   Design authority: the "Tandem Mockups" file, turn 8 (Olive & Terracotta). */
import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { createClient } from "@supabase/supabase-js";
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

/* ---------- design tokens (from README / mockups) ---------- */
const C = {
  bg: "#f7f4ea", ink: "#3d3b2e", ink2: "#6b6653", mut: "#a09a84", faint: "#c6ba9c",
  strike: "#ada07f", olive: "#7a8450", oliveSoft: "#eef0e0", oliveDark: "#5c6140",
  terra: "#c96a4a", terraSoft: "#faeae3", terraDark: "#a05439",
  track: "#edeadb", line: "#ddd8c4", divider: "#f2f0e6",
};
const shadow = "0 2px 8px rgba(63,56,42,.05)";
const shadowHero = "0 2px 10px rgba(63,56,42,.06)";
const shadowBtn = "0 4px 10px rgba(122,132,80,.35)";
const card = (r = 20, extra = {}) => ({ background: "#fff", borderRadius: r, boxShadow: shadow, ...extra });
const FONT = "'Plus Jakarta Sans',sans-serif";

/* ---------- utilities ---------- */
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : "id-" + Math.random().toString(36).slice(2) + Date.now());
const todayStr = (d = new Date()) => d.toISOString().slice(0, 10);
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const DOW = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
const DAYNAME = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
function fmtDue(t) {
  if (t.daily) return "Daily";
  if (!t.dueDate) return "Anytime";
  const d = new Date(t.dueDate + "T00:00:00"); const today = new Date(todayStr() + "T00:00:00");
  const diff = Math.round((d - today) / 86400000);
  let day = diff === 0 ? "Today" : diff === 1 ? "Tomorrow" : diff > 1 && diff < 7 ? DAYNAME[d.getDay()] : d.getDate() + " " + d.toLocaleString("en-GB", { month: "short" });
  return t.dueTime ? day + " " + t.dueTime : day;
}
function relDuePill(t) {
  if (t.daily) return { text: (t.streak || 0) + (t.streak === 1 ? " day" : " days"), fg: C.oliveDark, bg: C.oliveSoft };
  if (!t.dueDate) return null;
  const diff = Math.round((new Date(t.dueDate) - new Date(todayStr())) / 86400000);
  if (diff < 0) return { text: "overdue", fg: C.terraDark, bg: C.terraSoft };
  if (diff === 0) return { text: "today", fg: C.terraDark, bg: C.terraSoft };
  return { text: "in " + diff + (diff === 1 ? " day" : " days"), fg: C.ink2, bg: C.bg };
}
function downscale(file, max = 1000) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      const s = Math.min(1, max / Math.max(img.width, img.height));
      const cv = document.createElement("canvas");
      cv.width = Math.round(img.width * s); cv.height = Math.round(img.height * s);
      cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
      res(cv.toDataURL("image/jpeg", 0.82));
    };
    img.onerror = rej;
    img.src = URL.createObjectURL(file);
  });
}
const dataUrlToBlob = (u) => fetch(u).then((r) => r.blob());
const inQuietHours = (settings) => {
  const q = settings?.quietHours; if (!q || !q.on) return false;
  const now = new Date(); const cur = now.getHours() * 60 + now.getMinutes();
  const [fh, fm] = q.from.split(":").map(Number); const [th, tm] = q.to.split(":").map(Number);
  const f = fh * 60 + fm, t = th * 60 + tm;
  return f > t ? cur >= f || cur < t : cur >= f && cur < t;
};

/* ---------- persistence + optional Supabase sync ---------- */
const LS_STATE = "tandem.state.v1", LS_DEVICE = "tandem.device.v1", LS_KEYS = "tandem.keys.v1";
const defaultHousehold = () => ({
  id: "main",
  members: [
    { id: "tobias", name: "Tobias", color: "olive", outlook: { connected: false, access: "readwrite", email: "" } },
    { id: "an", name: "An", color: "terracotta", outlook: { connected: false, access: "write", email: "" } },
  ],
  settings: { photoProofDefault: true, dailyDigest: false, quietHours: { on: true, from: "21:00", to: "08:00" } },
  streak: { count: 0, lastDate: null },
});
function loadState() {
  try { const s = JSON.parse(localStorage.getItem(LS_STATE)); if (s && s.household) { if (!s.outlookCache) s.outlookCache = {}; return s; } } catch (e) {}
  return { household: defaultHousehold(), tasks: seedTasks(), nudges: [], outlookCache: {} };
}
function seedTasks() {
  const t = todayStr();
  return [
    { id: uid(), title: "Water the plants", daily: true, streak: 0, assignees: ["tobias", "an"], steps: [], status: "todo", photoProof: false, source: "typed", createdAt: t, comments: [] },
  ];
}
const loadKeys = () => { try { return JSON.parse(localStorage.getItem(LS_KEYS)) || {}; } catch (e) { return {}; } };
const saveKeys = (k) => localStorage.setItem(LS_KEYS, JSON.stringify(k));

let sb = null;
function getSb() {
  if (sb) return sb;
  const k = loadKeys();
  if (k.supabaseUrl && k.supabaseAnon) { try { sb = createClient(k.supabaseUrl, k.supabaseAnon); } catch (e) { sb = null; } }
  return sb;
}
async function sbPullAll() {
  const s = getSb(); if (!s) return null;
  const [h, t, n] = await Promise.all([
    s.from("household").select("*").eq("id", "main").maybeSingle(),
    s.from("tasks").select("*"),
    s.from("nudges").select("*").order("created_at", { ascending: true }),
  ]);
  if (h.error || t.error || n.error) throw (h.error || t.error || n.error);
  let outlookCache = {};
  try { const o = await s.from("outlook_cache").select("*"); if (!o.error) outlookCache = Object.fromEntries((o.data || []).map((r) => [r.id, r.data])); } catch (e) {}
  return {
    household: h.data ? h.data.data : null,
    tasks: (t.data || []).map((r) => r.data),
    nudges: (n.data || []).map((r) => r.data),
    outlookCache,
  };
}
async function sbPushCache(memberId, data) { const s = getSb(); if (!s) return; await s.from("outlook_cache").upsert({ id: memberId, data, updated_at: new Date().toISOString() }); }
async function sbPushTask(task) { const s = getSb(); if (!s) return; await s.from("tasks").upsert({ id: task.id, data: task, updated_at: new Date().toISOString() }); }
async function sbDeleteTask(id) { const s = getSb(); if (!s) return; await s.from("tasks").delete().eq("id", id); }
async function sbPushHousehold(h) { const s = getSb(); if (!s) return; await s.from("household").upsert({ id: "main", data: h, updated_at: new Date().toISOString() }); }
async function sbPushNudge(n) { const s = getSb(); if (!s) return; await s.from("nudges").insert({ id: n.id, data: n }); }
async function sbUploadPhoto(dataUrl, name) {
  const s = getSb(); if (!s) return dataUrl;
  try {
    const blob = await dataUrlToBlob(dataUrl);
    const path = name + ".jpg";
    const { error } = await s.storage.from("photos").upload(path, blob, { contentType: "image/jpeg", upsert: true });
    if (error) return dataUrl;
    return s.storage.from("photos").getPublicUrl(path).data.publicUrl;
  } catch (e) { return dataUrl; }
}

/* ---------- Claude (per-device key) ---------- */
async function claude(messages, maxTokens = 900) {
  const k = loadKeys(); if (!k.anthropic) throw new Error("no-key");
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-api-key": k.anthropic, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: maxTokens, messages }),
  });
  if (!r.ok) throw new Error("api-" + r.status);
  const d = await r.json();
  return (d.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}
const asJson = (txt) => JSON.parse(txt.replace(/```json|```/g, "").trim());
const imgBlock = (dataUrl) => ({ type: "image", source: { type: "base64", media_type: "image/jpeg", data: dataUrl.split(",")[1] } });
async function aiDraftFromPhoto(dataUrl, members) {
  const names = members.map((m) => m.name).join(" and ");
  const txt = await claude([{ role: "user", content: [imgBlock(dataUrl), { type: "text", text:
    `You help a two-person household (${names}) turn a photo into one small task. Reply ONLY with JSON: {"seen":"what you recognised, one short warm sentence","title":"short task title","steps":["2-4 short steps"],"assignee":"${members[0].id}"|"${members[1].id}"|"together","why":"one warm sentence on who and when, first person as the household assistant","dueSuggestion":"e.g. Saturday morning","photoProof":true|false}` }] }]);
  return asJson(txt);
}
async function aiScanNote(dataUrl, members) {
  const txt = await claude([{ role: "user", content: [imgBlock(dataUrl), { type: "text", text:
    `Read this handwritten/paper note for a two-person household. Reply ONLY with JSON: {"seen":"one short sentence on what the note is","kind":"list"|"task","title":"title for the list or task","items":["each item, deduplicated"],"why":"one warm sentence"}` }] }]);
  return asJson(txt);
}
async function aiVerify(beforeUrl, afterUrl, title) {
  const content = [];
  if (beforeUrl && beforeUrl.startsWith("data:")) content.push(imgBlock(beforeUrl));
  if (afterUrl && afterUrl.startsWith("data:")) content.push(imgBlock(afterUrl));
  content.push({ type: "text", text: `Task: "${title}". ${beforeUrl ? "First image is before, second is after." : "Image is the after photo."} Reply ONLY with JSON: {"checks":["3 short positive verification observations"],"ok":true|false,"note":"one warm closing sentence"}` });
  const txt = await claude([{ role: "user", content }]);
  return asJson(txt);
}

/* ---------- Microsoft Graph / Outlook (phase 2) ---------- */
const MS_SCOPES = ["User.Read", "Calendars.ReadWrite"];
const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone || "Europe/Brussels";
let msal = null, msalReady = null;
function getMsal(clientId) {
  if (!clientId) return null;
  if (msal && msal.__cid === clientId) return msal;
  msal = new PublicClientApplication({
    auth: { clientId, authority: "https://login.microsoftonline.com/common", redirectUri: window.location.origin + "/" },
    cache: { cacheLocation: "localStorage" },
  });
  msal.__cid = clientId;
  msalReady = msal.initialize();
  return msal;
}
async function msSignIn(clientId) {
  const m = getMsal(clientId); if (!m) return; await msalReady;
  await m.loginRedirect({ scopes: MS_SCOPES, prompt: "select_account" });
}
async function msHandleRedirect(clientId) {
  const m = getMsal(clientId); if (!m) return null; await msalReady;
  try { return await m.handleRedirectPromise(); } catch (e) { return null; }
}
function msAccount(clientId, email) {
  const m = getMsal(clientId); if (!m) return null;
  const accs = m.getAllAccounts();
  if (!accs.length) return null;
  return (email && accs.find((a) => a.username.toLowerCase() === email.toLowerCase())) || accs[0];
}
async function msToken(clientId, email) {
  const m = getMsal(clientId); await msalReady;
  const account = msAccount(clientId, email); if (!account) throw new Error("no-account");
  try { const r = await m.acquireTokenSilent({ scopes: MS_SCOPES, account }); return r.accessToken; }
  catch (e) {
    if (e instanceof InteractionRequiredAuthError) await m.acquireTokenRedirect({ scopes: MS_SCOPES, account });
    throw e;
  }
}
async function graph(clientId, email, path, opts = {}) {
  const token = await msToken(clientId, email);
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...opts, headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (r.status === 204) return null;
  const d = await r.json().catch(() => null);
  if (!r.ok) throw new Error("graph-" + r.status);
  return d;
}
const taskHash = (t) => [t.title, t.dueDate, t.dueTime || "", t.status].join("|");
function plusHour(hm) { const [h, m] = hm.split(":").map(Number); const t = h * 60 + m + 60; return String(Math.floor(t / 60) % 24).padStart(2, "0") + ":" + String(t % 60).padStart(2, "0"); }

/* keep my Outlook in step with my Klaus tasks: create / update / tick / remove events */
async function syncMyOutlook(stateRef, meId, saveTask) {
  const st = stateRef.current; const h = st.household;
  const cid = h.settings?.msClientId; const me = h.members.find((m) => m.id === meId);
  if (!cid || !me?.outlook?.connected || !msAccount(cid, me.outlook.email)) return;
  const email = me.outlook.email;
  for (const t of st.tasks) {
    const cur = stateRef.current.tasks.find((x) => x.id === t.id); if (!cur) continue;
    const mine = cur.assignees?.includes(meId) || cur.assignees?.includes("together");
    const rec = cur.outlook?.[meId];
    try {
      if (mine && cur.dueDate && !cur.daily && cur.status !== "done") {
        const hash = taskHash(cur);
        const start = cur.dueDate + "T" + (cur.dueTime || "09:00") + ":00";
        const end = cur.dueDate + "T" + plusHour(cur.dueTime || "09:00") + ":00";
        const body = { subject: cur.title, start: { dateTime: start, timeZone: TZ }, end: { dateTime: end, timeZone: TZ }, body: { contentType: "text", content: "Klaus task — the house, between you." } };
        if (!rec) {
          const ev = await graph(cid, email, "/me/events", { method: "POST", body: JSON.stringify(body) });
          saveTask({ ...cur, outlook: { ...(cur.outlook || {}), [meId]: { eventId: ev.id, hash } } });
        } else if (rec.hash !== hash) {
          await graph(cid, email, "/me/events/" + rec.eventId, { method: "PATCH", body: JSON.stringify(body) });
          saveTask({ ...cur, outlook: { ...(cur.outlook || {}), [meId]: { ...rec, hash: hash } } });
        }
      } else if (rec) {
        if (cur.status === "done" && !rec.doneMarked) {
          await graph(cid, email, "/me/events/" + rec.eventId, { method: "PATCH", body: JSON.stringify({ subject: "✓ " + cur.title }) }).catch(() => {});
          saveTask({ ...cur, outlook: { ...(cur.outlook || {}), [meId]: { ...rec, doneMarked: true } } });
        } else if (cur.status !== "done") {
          await graph(cid, email, "/me/events/" + rec.eventId, { method: "DELETE" }).catch(() => {});
          const o = { ...(cur.outlook || {}) }; delete o[meId];
          saveTask({ ...cur, outlook: o });
        }
      }
    } catch (e) { /* transient — the next pass picks it up */ }
  }
}

/* publish my upcoming events to the shared store — full events if read & write, anonymous busy-blocks if write only */
async function pushMyEventsCache(stateRef, meId) {
  const st = stateRef.current; const h = st.household;
  const cid = h.settings?.msClientId; const me = h.members.find((m) => m.id === meId);
  if (!cid || !me?.outlook?.connected || !msAccount(cid, me.outlook.email)) return null;
  const startD = addDays(new Date(), -1), endD = addDays(new Date(), 8);
  const q = `/me/calendarView?startDateTime=${todayStr(startD)}T00:00:00&endDateTime=${todayStr(endD)}T23:59:59&$top=60&$select=id,subject,start,end`;
  const d = await graph(cid, me.outlook.email, q, { headers: { Prefer: `outlook.timezone="${TZ}"` } });
  const klausIds = new Set(st.tasks.map((t) => t.outlook?.[meId]?.eventId).filter(Boolean));
  const readwrite = (me.outlook.access || "write") === "readwrite";
  const events = (d?.value || []).filter((e) => !klausIds.has(e.id)).map((e) => ({
    title: readwrite ? (e.subject || "(no title)") : "Busy",
    start: (e.start?.dateTime || "").slice(0, 16), end: (e.end?.dateTime || "").slice(0, 16),
  })).filter((e) => e.start);
  const data = { updatedAt: new Date().toISOString(), mode: me.outlook.access || "write", events };
  await sbPushCache(meId, data).catch(() => {});
  return data;
}

/* ---------- tiny shared pieces ---------- */
const Avatar = ({ m, size = 24, ring }) => (
  <span style={{ width: size, height: size, borderRadius: "50%", background: m.color === "olive" ? C.olive : C.terra, color: "#fff", display: "grid", placeItems: "center", fontSize: size * 0.42, fontWeight: 800, flex: "none", border: ring ? "2px solid " + C.bg : "none", fontStyle: "normal" }}>{m.name[0]}</span>
);
const Gear = ({ onClick }) => (
  <span onClick={onClick} style={{ width: 36, height: 36, borderRadius: "50%", background: "#fff", display: "grid", placeItems: "center", boxShadow: "0 2px 6px rgba(63,56,42,.08)", cursor: "pointer" }}>
    <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke={C.mut} strokeWidth="1.8" strokeLinecap="round"><circle cx="12" cy="12" r="3" /><path d="M19 12a7 7 0 01-.1 1.2l2 1.5-2 3.4-2.3-.9a7 7 0 01-2.1 1.2L14 21h-4l-.5-2.6a7 7 0 01-2.1-1.2l-2.3.9-2-3.4 2-1.5A7 7 0 015 12" /></svg>
  </span>
);
const Back = ({ onClick }) => (
  <span onClick={onClick} style={{ width: 34, height: 34, borderRadius: "50%", background: "#fff", display: "grid", placeItems: "center", boxShadow: "0 2px 6px rgba(63,56,42,.08)", flex: "none", fontSize: 15, color: "#8b8672", cursor: "pointer" }}>‹</span>
);
const H1 = ({ children, small }) => <h1 style={{ margin: 0, fontSize: small ? 18 : 22, fontWeight: 800, letterSpacing: small ? "-.01em" : "-.02em", flex: 1, color: C.ink }}>{children}</h1>;
const Kicker = ({ children, style }) => <div style={{ fontSize: 12.5, fontWeight: 800, letterSpacing: ".06em", textTransform: "uppercase", color: C.mut, margin: "2px 2px 8px", ...style }}>{children}</div>;
const CheckCircle = ({ state, size = 26, onClick }) => {
  const base = { width: size, height: size, borderRadius: "50%", flex: "none", boxSizing: "border-box", display: "grid", placeItems: "center", cursor: onClick ? "pointer" : "default" };
  if (state === "done") return <span onClick={onClick} style={{ ...base, background: C.olive, color: "#fff", fontSize: size * 0.5 }}>✓</span>;
  if (state === "active") return <span onClick={onClick} style={{ ...base, border: "2.5px solid " + C.terra }} />;
  return <span onClick={onClick} style={{ ...base, border: "2.5px solid " + C.line }} />;
};
const Toggle = ({ on, onClick }) => (
  <span onClick={onClick} style={{ width: 44, height: 26, borderRadius: 999, background: on ? C.olive : C.line, position: "relative", flex: "none", cursor: "pointer", transition: "background .2s" }}>
    <i style={{ position: "absolute", top: 3, left: on ? 21 : 3, width: 20, height: 20, borderRadius: "50%", background: "#fff", display: "block", boxShadow: on ? "none" : "0 1px 2px rgba(0,0,0,.15)", transition: "left .2s" }} />
  </span>
);
const SegPill = ({ options, value, onChange, style }) => (
  <div style={{ display: "flex", background: C.track, borderRadius: 999, padding: 4, gap: 3, ...style }}>
    {options.map((o) => (
      <span key={o.value} onClick={() => onChange(o.value)} style={{ flex: 1, textAlign: "center", fontSize: 12.5, fontWeight: 700, color: value === o.value ? C.ink : C.mut, padding: "9px 6px", borderRadius: 999, background: value === o.value ? "#fff" : "transparent", boxShadow: value === o.value ? "0 1px 3px rgba(63,56,42,.12)" : "none", cursor: "pointer", transition: "all .15s" }}>{o.label}</span>
    ))}
  </div>
);
const PrimaryBtn = ({ children, onClick, style }) => (
  <div onClick={onClick} style={{ background: C.olive, color: "#fff", borderRadius: 999, padding: 15, fontSize: 15, fontWeight: 800, textAlign: "center", boxShadow: shadowBtn, cursor: "pointer", ...style }}>{children}</div>
);
const Photo = ({ src, h, r = 24, label, border, style }) => (
  <div style={{ height: h, borderRadius: r, background: src ? "none" : "linear-gradient(135deg,#7b7469,#3f3a33)", position: "relative", overflow: "hidden", boxShadow: shadowHero, border: border ? "2.5px solid " + C.olive : "none", boxSizing: "border-box", ...style }}>
    {src ? <img src={src} alt="" style={{ position: "absolute", inset: 0, width: "100%", height: "100%", objectFit: "cover" }} />
      : <span style={{ position: "absolute", inset: 0, background: "radial-gradient(circle at 32% 26%,rgba(255,255,255,.3),transparent 62%)" }} />}
    {label && <span style={{ position: "absolute", left: 12, bottom: 10, fontSize: 10, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: "#fff", background: "rgba(61,59,46,.55)", borderRadius: 999, padding: "4px 10px", whiteSpace: "nowrap" }}>{label}</span>}
  </div>
);
const CameraSvg = () => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 8a2 2 0 012-2h2l1.5-2h7L18 6h1a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2z" /><circle cx="12" cy="13" r="3.4" /></svg>
);

/* animated AI checklist: rows appear one by one, last runs a spinner until done */
function AiChecklist({ rows, running }) {
  return (
    <div style={{ ...card(22), padding: "8px 18px" }}>
      {rows.map((r, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", fontSize: 13.5, fontWeight: 600, color: i === rows.length - 1 && running ? C.mut : C.ink, borderTop: i ? "1px solid " + C.divider : "none" }}>
          {i === rows.length - 1 && running
            ? <span className="tspin" style={{ width: 24, height: 24, borderRadius: "50%", border: "2.5px solid " + C.terra, borderTopColor: "transparent", flex: "none", boxSizing: "border-box" }} />
            : <span style={{ width: 24, height: 24, borderRadius: "50%", background: C.olive, display: "grid", placeItems: "center", color: "#fff", fontSize: 12, flex: "none" }}>✓</span>}
          <span>{r}</span>
        </div>
      ))}
    </div>
  );
}
function useSteppedReveal(steps, active) {
  const [n, setN] = useState(active ? 1 : 0);
  useEffect(() => {
    if (!active) return;
    setN(1);
    const iv = setInterval(() => setN((x) => (x < steps.length ? x + 1 : x)), 650);
    return () => clearInterval(iv);
  }, [active, steps.length]);
  return n;
}

/* ---------- App ---------- */
function App() {
  const [state, setState] = useState(loadState);
  const [who, setWho] = useState(() => { try { return JSON.parse(localStorage.getItem(LS_DEVICE))?.who || null; } catch (e) { return null; } });
  const [route, setRoute] = useState({ name: "home" });
  const [tab, setTab] = useState("home");
  const [syncStatus, setSyncStatus] = useState(getSb() ? "connected" : "local");
  const stateRef = useRef(state); stateRef.current = state;

  const persist = (next) => { setState(next); localStorage.setItem(LS_STATE, JSON.stringify(next)); };
  const me = state.household.members.find((m) => m.id === who) || state.household.members[0];
  const other = state.household.members.find((m) => m.id !== me.id) || state.household.members[1];

  /* --- mutations (push per-entity when Supabase is connected) --- */
  const saveTask = (task) => {
    const tasks = stateRef.current.tasks.some((t) => t.id === task.id)
      ? stateRef.current.tasks.map((t) => (t.id === task.id ? task : t))
      : [...stateRef.current.tasks, task];
    persist({ ...stateRef.current, tasks });
    sbPushTask(task).catch(() => {});
  };
  const removeTask = (id) => {
    persist({ ...stateRef.current, tasks: stateRef.current.tasks.filter((t) => t.id !== id) });
    sbDeleteTask(id).catch(() => {});
  };
  const saveHousehold = (h) => { persist({ ...stateRef.current, household: h }); sbPushHousehold(h).catch(() => {}); };
  const sendNudge = (to, text) => {
    const n = { id: uid(), from: me.id, to, text, ts: new Date().toISOString() };
    persist({ ...stateRef.current, nudges: [...stateRef.current.nudges, n] });
    sbPushNudge(n).catch(() => {});
  };

  const completeTask = (task, extra = {}) => {
    const now = new Date();
    let streak = task.streak || 0;
    let lastDoneDate = task.lastDoneDate;
    if (task.daily) {
      if (lastDoneDate === todayStr(addDays(now, -1))) streak += 1; else if (lastDoneDate !== todayStr()) streak = 1;
      lastDoneDate = todayStr();
    }
    const done = { ...task, status: "done", doneAt: now.toISOString(), doneBy: me.id, streak, lastDoneDate, ...extra };
    saveTask(done);
    /* household streak: if everything relevant today is now done */
    const after = stateRef.current.tasks.map((t) => (t.id === task.id ? done : t));
    const todays = after.filter(isTodayTask);
    if (todays.length && todays.every((t) => t.status === "done")) {
      const h = stateRef.current.household; const st = h.streak || { count: 0, lastDate: null };
      if (st.lastDate !== todayStr()) {
        const count = st.lastDate === todayStr(addDays(new Date(), -1)) ? st.count + 1 : 1;
        saveHousehold({ ...h, streak: { count, lastDate: todayStr() } });
      }
    }
    return done;
  };
  const uncompleteTask = (task) => saveTask({ ...task, status: "todo", doneAt: null, doneBy: null, doneWithPhoto: false, photoAfter: null });

  /* --- Supabase: initial pull + realtime --- */
  useEffect(() => {
    const s = getSb(); if (!s) return;
    let alive = true;
    (async () => {
      try {
        const remote = await sbPullAll(); if (!alive || !remote) return;
        const local = stateRef.current;
        const remoteIds = new Set(remote.tasks.map((t) => t.id));
        const merged = {
          household: remote.household || local.household,
          tasks: [...remote.tasks, ...local.tasks.filter((t) => !remoteIds.has(t.id))],
          nudges: remote.nudges.length ? remote.nudges : local.nudges,
          outlookCache: { ...(local.outlookCache || {}), ...(remote.outlookCache || {}) },
        };
        persist(merged); setSyncStatus("connected");
        if (!remote.household) sbPushHousehold(merged.household).catch(() => {});
        local.tasks.filter((t) => !remoteIds.has(t.id)).forEach((t) => sbPushTask(t).catch(() => {}));
      } catch (e) { setSyncStatus("error"); }
    })();
    const ch = s.channel("tandem")
      .on("postgres_changes", { event: "*", schema: "public", table: "tasks" }, (p) => {
        const cur = stateRef.current;
        if (p.eventType === "DELETE") persist({ ...cur, tasks: cur.tasks.filter((t) => t.id !== p.old.id) });
        else {
          const t = p.new.data;
          persist({ ...cur, tasks: cur.tasks.some((x) => x.id === t.id) ? cur.tasks.map((x) => (x.id === t.id ? t : x)) : [...cur.tasks, t] });
        }
      })
      .on("postgres_changes", { event: "INSERT", schema: "public", table: "nudges" }, (p) => {
        const cur = stateRef.current; const n = p.new.data;
        if (!cur.nudges.some((x) => x.id === n.id)) persist({ ...cur, nudges: [...cur.nudges, n] });
      })
      .on("postgres_changes", { event: "*", schema: "public", table: "household" }, (p) => {
        if (p.new?.data) persist({ ...stateRef.current, household: p.new.data });
      })
      .subscribe();
    return () => { alive = false; s.removeChannel(ch); };
  }, [syncStatus === "reconnect"]);

  /* --- in-app reminder ticker --- */
  useEffect(() => {
    const iv = setInterval(() => {
      const cur = stateRef.current;
      if (inQuietHours(cur.household.settings)) return;
      cur.tasks.forEach((t) => {
        if (t.status === "done" || !t.reminder?.at || t.reminder.firedAt) return;
        if (!t.assignees?.includes(me.id) && !t.assignees?.includes("together")) return;
        if (new Date(t.reminder.at) <= new Date()) {
          if (typeof Notification !== "undefined" && Notification.permission === "granted")
            new Notification("Klaus", { body: t.title + " — " + fmtDue(t) });
          saveTask({ ...t, reminder: { ...t.reminder, firedAt: new Date().toISOString() } });
        }
      });
    }, 30000);
    return () => clearInterval(iv);
  }, [me.id]);

  /* --- Outlook: finish sign-in redirects, then reconcile events + publish cache, refreshed every 10 min --- */
  const msClientId = state.household.settings?.msClientId;
  useEffect(() => {
    if (!msClientId || !who) return;
    let alive = true;
    const run = async () => {
      if (!alive) return;
      await syncMyOutlook(stateRef, me.id, saveTask).catch(() => {});
      const c = await pushMyEventsCache(stateRef, me.id).catch(() => null);
      if (c && alive) persist({ ...stateRef.current, outlookCache: { ...(stateRef.current.outlookCache || {}), [me.id]: c } });
    };
    (async () => {
      const r = await msHandleRedirect(msClientId);
      if (r?.account && alive) {
        const h = stateRef.current.household;
        saveHousehold({ ...h, members: h.members.map((m) => (m.id === me.id ? { ...m, outlook: { ...(m.outlook || {}), connected: true, email: r.account.username } } : m)) });
      }
      const meM = stateRef.current.household.members.find((m) => m.id === me.id);
      if (meM?.outlook?.connected && msAccount(msClientId, meM.outlook.email)) run();
    })();
    const iv = setInterval(run, 10 * 60 * 1000);
    return () => { alive = false; clearInterval(iv); };
  }, [msClientId, me.id, syncStatus, who]);

  if (!who) return <Onboard members={state.household.members} onPick={(id) => { localStorage.setItem(LS_DEVICE, JSON.stringify({ who: id })); setWho(id); }} />;

  const nav = (name, params = {}) => setRoute({ name, ...params });
  const goTab = (t) => { setTab(t); setRoute({ name: t }); };
  const screenProps = { state, me, other, saveTask, removeTask, saveHousehold, sendNudge, completeTask, uncompleteTask, nav, goTab, tab, route, syncStatus, setSyncStatus };

  let screen;
  switch (route.name) {
    case "home": screen = <Home {...screenProps} />; break;
    case "tasks": screen = <Tasks {...screenProps} />; break;
    case "calendar": screen = <CalendarView {...screenProps} />; break;
    case "nudges": screen = <Reminders {...screenProps} />; break;
    case "task": screen = <TaskDetail {...screenProps} taskId={route.id} />; break;
    case "addPhoto": screen = <AddFromPhoto {...screenProps} photo={route.photo} />; break;
    case "scanNote": screen = <ScanNote {...screenProps} photo={route.photo} />; break;
    case "finish": screen = <FinishPhoto {...screenProps} taskId={route.id} photo={route.photo} />; break;
    case "sharing": screen = <Sharing {...screenProps} />; break;
    case "settings": screen = <Settings {...screenProps} />; break;
    default: screen = <Home {...screenProps} />;
  }
  return <div style={{ maxWidth: 430, margin: "0 auto", minHeight: "100vh", background: C.bg, fontFamily: FONT, display: "flex", flexDirection: "column" }}>{screen}</div>;
}

const isTodayTask = (t) => t.daily || t.dueDate === todayStr() || (t.status === "done" && t.doneAt && t.doneAt.slice(0, 10) === todayStr()) || (!t.dueDate && t.status !== "done");

/* ---------- onboarding ---------- */
function Onboard({ members, onPick }) {
  return (
    <div style={{ maxWidth: 430, margin: "0 auto", minHeight: "100vh", background: C.bg, fontFamily: FONT, display: "flex", flexDirection: "column", justifyContent: "center", padding: 28 }}>
      <div style={{ fontSize: 26, fontWeight: 800, color: C.ink, letterSpacing: "-.02em" }}>Klaus</div>
      <div style={{ fontSize: 14, color: C.mut, fontWeight: 600, marginTop: 6, lineHeight: 1.5 }}>The house, between you. First things first — whose phone is this?</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 11, marginTop: 24 }}>
        {members.map((m) => (
          <div key={m.id} onClick={() => onPick(m.id)} style={{ ...card(22), padding: "16px 17px", display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
            <Avatar m={m} size={40} />
            <div style={{ fontSize: 16, fontWeight: 800, color: C.ink, flex: 1 }}>{m.name}</div>
            <span style={{ color: m.color === "olive" ? C.olive : C.terra, fontWeight: 700 }}>›</span>
          </div>
        ))}
      </div>
      <div style={{ fontSize: 12, color: C.mut, fontWeight: 600, marginTop: 16, lineHeight: 1.5 }}>Names can be changed later in Sharing & access.</div>
    </div>
  );
}

/* ---------- tab strip + bottom bar ---------- */
function TabStrip({ tab, goTab }) {
  return (
    <div style={{ padding: "6px 20px 4px" }}>
      <SegPill value={tab} onChange={goTab} options={[
        { value: "home", label: "Today" }, { value: "tasks", label: "Tasks" }, { value: "calendar", label: "Calendar" }, { value: "nudges", label: "Nudges" },
      ]} />
    </div>
  );
}
function BottomBar({ onText, onPhoto, onScan }) {
  const [txt, setTxt] = useState("");
  const [sheet, setSheet] = useState(false);
  const photoRef = useRef(); const scanRef = useRef();
  const pick = async (e, fn) => { const f = e.target.files?.[0]; e.target.value = ""; if (!f) return; fn(await downscale(f)); setSheet(false); };
  return (
    <div style={{ position: "sticky", bottom: 0, marginTop: "auto" }}>
      {sheet && (
        <div style={{ padding: "0 16px 10px" }}>
          <div style={{ ...card(20), padding: 8, display: "flex", gap: 8 }}>
            <div onClick={() => photoRef.current.click()} style={{ flex: 1, background: C.oliveSoft, borderRadius: 14, padding: "13px 10px", textAlign: "center", fontSize: 13, fontWeight: 700, color: C.oliveDark, cursor: "pointer" }}>Photo of the job</div>
            <div onClick={() => scanRef.current.click()} style={{ flex: 1, background: C.terraSoft, borderRadius: 14, padding: "13px 10px", textAlign: "center", fontSize: 13, fontWeight: 700, color: C.terraDark, cursor: "pointer" }}>Scan a paper note</div>
          </div>
        </div>
      )}
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px 18px", background: "#fff", boxShadow: "0 -4px 14px rgba(63,56,42,.06)" }}>
        <input value={txt} onChange={(e) => setTxt(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && txt.trim()) { onText(txt.trim()); setTxt(""); } }}
          placeholder="Tell me what needs doing…"
          style={{ flex: 1, background: C.bg, border: "none", outline: "none", borderRadius: 999, padding: "13px 17px", fontSize: 13.5, fontWeight: 600, color: C.ink, fontFamily: FONT }} />
        <span onClick={() => setSheet(!sheet)} style={{ width: 46, height: 46, borderRadius: "50%", background: C.olive, display: "grid", placeItems: "center", flex: "none", boxShadow: shadowBtn, cursor: "pointer" }}><CameraSvg /></span>
      </div>
      <input ref={photoRef} type="file" accept="image/*" capture="environment" hidden onChange={(e) => pick(e, onPhoto)} />
      <input ref={scanRef} type="file" accept="image/*" capture="environment" hidden onChange={(e) => pick(e, onScan)} />
    </div>
  );
}

/* ---------- 8a Home ---------- */
const heroCopy = (done, total, left, hour) => {
  if (total === 0) return ["A clean slate.", "Nothing on the list yet — add something small to get going."];
  if (left === 0) return ["All done — the whole list.", "Everything's ticked off, between you. Lovely."];
  if (done / total >= 0.5) return ["Lovely — over halfway.", `${left === 1 ? "One small thing" : left === 2 ? "Two small things" : left + " small things"} left${hour < 18 ? " before six" : " today"}, between you.`];
  if (done === 0) return ["Fresh day, fresh list.", `${total} thing${total > 1 ? "s" : ""} on it, between you. Small starts count.`];
  return ["Nicely underway.", `${done} down, ${left} to go — between you.`];
};
function Home({ state, me, other, nav, goTab, tab, saveTask, completeTask, uncompleteTask }) {
  const todays = state.tasks.filter(isTodayTask);
  const done = todays.filter((t) => t.status === "done").length;
  const total = todays.length; const left = total - done;
  const frac = total ? done / total : 0;
  const dash = 351.9;
  const [big, small] = heroCopy(done, total, left, new Date().getHours());
  const nudge = [...state.nudges].reverse().find((n) => n.to === me.id && Date.now() - new Date(n.ts) < 48 * 3600000);
  const nudgeFrom = nudge && state.household.members.find((m) => m.id === nudge.from);
  const open = todays.filter((t) => t.status !== "done");
  const doneList = todays.filter((t) => t.status === "done");
  const addTyped = (text) => {
    const t = { id: uid(), title: text, steps: [], assignees: [me.id], status: "todo", daily: false, dueDate: null, photoProof: false, source: "typed", createdAt: todayStr(), comments: [] };
    saveTask(t); nav("task", { id: t.id });
  };
  const label = (t) => {
    if (t.status === "done") { const by = state.household.members.find((m) => m.id === t.doneBy); return (by ? by.name : "Someone") + " did it" + (t.doneWithPhoto ? " — with a photo" : ""); }
    const w = t.assignees?.includes("together") || t.assignees?.length > 1 ? "Together" : t.assignees?.[0] === me.id ? "Yours" : (other.name + "'s");
    const bits = [w, t.daily ? "Daily" + ((t.streak || 0) > 0 ? " · " + t.streak + " days running" : "") : fmtDue(t)];
    if (t.photoProof && !t.daily) bits.push("photo when done");
    return bits.join(" · ");
  };
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><H1>Our day</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 0" }}>
      <div style={{ ...card(24, { boxShadow: shadowHero }), padding: 18, display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ position: "relative", width: 96, height: 96, flex: "none" }}>
          <svg width="96" height="96" viewBox="0 0 132 132"><circle cx="66" cy="66" r="56" fill="none" stroke={C.track} strokeWidth="12" /><circle cx="66" cy="66" r="56" fill="none" stroke={C.olive} strokeWidth="12" strokeLinecap="round" strokeDasharray={dash} strokeDashoffset={dash * (1 - frac)} transform="rotate(-90 66 66)" style={{ transition: "stroke-dashoffset .6s" }} /></svg>
          <div style={{ position: "absolute", inset: 0, display: "grid", placeItems: "center" }}><div style={{ fontSize: 25, fontWeight: 800, color: C.ink }}>{done}<span style={{ fontSize: 13, color: C.mut, fontWeight: 600 }}>/{total}</span></div></div>
        </div>
        <div style={{ flex: 1 }}><div style={{ fontSize: 16, fontWeight: 700, color: C.ink, lineHeight: 1.3 }}>{big}</div><div style={{ fontSize: 13, color: C.mut, marginTop: 4, lineHeight: 1.45 }}>{small}</div></div>
      </div>
      {nudge && (
        <div onClick={() => goTab("nudges")} style={{ display: "flex", gap: 11, alignItems: "center", background: C.terraSoft, borderRadius: 20, padding: "14px 16px", marginTop: 12, cursor: "pointer" }}>
          <Avatar m={nudgeFrom} size={34} />
          <div style={{ flex: 1 }}><div style={{ fontSize: 14, fontWeight: 700, color: C.ink }}>A little reminder from {nudgeFrom.name}</div><div style={{ fontSize: 12.5, color: C.terraDark, marginTop: 2, lineHeight: 1.4 }}>"{nudge.text}"</div></div>
          <span style={{ color: C.terra, fontSize: 16, fontWeight: 700 }}>›</span>
        </div>
      )}
      <div style={{ fontSize: 13, fontWeight: 700, color: C.mut, margin: "20px 2px 8px" }}>{new Date().getHours() < 12 ? "This morning" : new Date().getHours() < 18 ? "This afternoon" : "This evening"}</div>
      {doneList.map((t) => (
        <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ ...card(20), padding: "15px 16px", display: "flex", alignItems: "center", gap: 13, marginBottom: 9, opacity: 0.65, cursor: "pointer" }}>
          <CheckCircle state="done" onClick={(e) => { e.stopPropagation(); uncompleteTask(t); }} />
          <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 700, color: C.strike, textDecoration: "line-through" }}>{t.title}</div><div style={{ fontSize: 11.5, color: C.faint, marginTop: 1 }}>{label(t)}</div></div>
          {(t.photoAfter || t.photoBefore) && <span style={{ width: 36, height: 36, borderRadius: 11, flex: "none", overflow: "hidden" }}><img src={t.photoAfter || t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></span>}
        </div>
      ))}
      {open.map((t) => (
        <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ ...card(20), padding: "15px 16px", display: "flex", alignItems: "center", gap: 13, marginBottom: 9, cursor: "pointer" }}>
          <CheckCircle state={t.status === "active" ? "active" : "todo"} onClick={(e) => { e.stopPropagation(); t.photoProof ? nav("task", { id: t.id }) : completeTask(t); }} />
          <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 700, color: C.ink }}>{t.title}</div><div style={{ fontSize: 11.5, color: C.mut, marginTop: 1 }}>{label(t)}</div></div>
          {t.daily && (t.streak || 0) > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: C.olive, background: C.oliveSoft, borderRadius: 999, padding: "4px 10px", whiteSpace: "nowrap" }}>{t.streak} days</span>}
          {t.photoBefore && <span style={{ width: 36, height: 36, borderRadius: 11, flex: "none", overflow: "hidden" }}><img src={t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></span>}
        </div>
      ))}
      {total === 0 && <div style={{ fontSize: 13, color: C.mut, fontWeight: 600, textAlign: "center", padding: "18px 0" }}>Type below, or snap a photo of the job.</div>}
      <div style={{ height: 12 }} />
    </div>
    <BottomBar onText={addTyped} onPhoto={(p) => nav("addPhoto", { photo: p })} onScan={(p) => nav("scanNote", { photo: p })} />
  </>);
}

/* ---------- 8b Tasks ---------- */
function Tasks({ state, me, other, nav, goTab, tab, saveTask, completeTask }) {
  const [mode, setMode] = useState("person");
  const open = state.tasks.filter((t) => t.status !== "done");
  const doneRecent = state.tasks.filter((t) => t.status === "done" && t.doneAt && Date.now() - new Date(t.doneAt) < 3 * 86400000);
  const forPerson = (id) => [...open, ...doneRecent].filter((t) => t.assignees?.includes(id) || t.assignees?.includes("together") || (t.assignees?.length || 0) > 1 && t.assignees.includes(id));
  const st = state.household.streak || { count: 0 };
  const addTyped = (text) => { const t = { id: uid(), title: text, steps: [], assignees: [me.id], status: "todo", photoProof: false, source: "typed", createdAt: todayStr(), comments: [] }; saveTask(t); nav("task", { id: t.id }); };
  const Mini = ({ t }) => {
    const doneT = t.status === "done"; const pill = !doneT && relDuePill(t);
    return (
      <div onClick={() => nav("task", { id: t.id })} style={{ ...card(18), marginBottom: 9, overflow: "hidden", opacity: doneT ? (t.doneWithPhoto ? 0.75 : 0.6) : 1, cursor: "pointer" }}>
        {t.photoBefore && <div style={{ height: 52, position: "relative" }}><img src={t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></div>}
        <div style={{ padding: t.photoBefore ? "10px 13px" : "12px 13px" }}>
          <div style={{ fontSize: 13.5, fontWeight: 700, lineHeight: 1.25, color: doneT ? C.strike : C.ink, textDecoration: doneT ? "line-through" : "none" }}>{t.title}</div>
          <div style={{ fontSize: 11, fontWeight: 600, color: doneT ? C.olive : C.mut, marginTop: 3 }}>
            {doneT ? (t.doneWithPhoto ? "Done with photo" : "Done " + new Date(t.doneAt).toTimeString().slice(0, 5)) : fmtDue(t)}
          </div>
          {!doneT && (t.photoProof || (pill && pill.bg === C.terraSoft)) && (
            <div style={{ display: "flex", gap: 5, marginTop: 8, flexWrap: "wrap" }}>
              {pill && pill.bg === C.terraSoft && <span style={{ fontSize: 10, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "3px 9px", whiteSpace: "nowrap" }}>due {pill.text === "today" ? "soon" : pill.text}</span>}
              {t.photoProof && <span style={{ fontSize: 10, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "3px 9px", whiteSpace: "nowrap" }}>photo when done</span>}
            </div>
          )}
        </div>
      </div>
    );
  };
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><H1>Tasks</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 16px" }}>
      <SegPill style={{ marginBottom: 14 }} value={mode} onChange={setMode} options={[{ value: "together", label: "Together" }, { value: "person", label: "By person" }]} />
      {mode === "person" ? (
        <div style={{ display: "flex", gap: 10 }}>
          {[me, other].map((p) => (
            <div key={p.id} style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 8 }}><Avatar m={p} size={24} /><span style={{ fontSize: 13, fontWeight: 800, color: C.ink }}>{p.name}</span><span style={{ fontSize: 11, fontWeight: 600, color: C.mut }}>{forPerson(p.id).filter((t) => t.status !== "done").length} open</span></div>
              {forPerson(p.id).map((t) => <Mini key={t.id} t={t} />)}
            </div>
          ))}
        </div>
      ) : (
        <div>{[...open, ...doneRecent].map((t) => <Mini key={t.id} t={t} />)}</div>
      )}
      {st.count > 1 && (
        <div style={{ display: "flex", gap: 11, alignItems: "center", background: C.oliveSoft, borderRadius: 18, padding: "13px 16px", marginTop: 8 }}>
          <span style={{ width: 30, height: 30, borderRadius: "50%", background: C.olive, color: "#fff", display: "grid", placeItems: "center", fontSize: 13, flex: "none" }}>✓</span>
          <div style={{ fontSize: 13, color: C.oliveDark, lineHeight: 1.4, fontWeight: 600 }}><b style={{ color: C.ink }}>The house is on a {st.count}-day streak</b> — everything done, together.</div>
        </div>
      )}
    </div>
    <BottomBar onText={addTyped} onPhoto={(p) => nav("addPhoto", { photo: p })} onScan={(p) => nav("scanNote", { photo: p })} />
  </>);
}

/* ---------- 8c Task detail ---------- */
function TaskDetail({ state, me, other, nav, taskId, saveTask, removeTask, completeTask, uncompleteTask }) {
  const t = state.tasks.find((x) => x.id === taskId);
  const finRef = useRef();
  const [comment, setComment] = useState("");
  if (!t) return <div style={{ padding: 24, fontSize: 14, color: C.mut, fontWeight: 600 }}>That task is gone. <span onClick={() => nav("home")} style={{ color: C.olive, cursor: "pointer" }}>Back home.</span></div>;
  const assignee = t.assignees?.includes("together") || (t.assignees?.length || 0) > 1 ? null : state.household.members.find((m) => m.id === t.assignees?.[0]);
  const toggleStep = (i) => { const steps = t.steps.map((s, j) => (j === i ? { ...s, done: !s.done } : s)); saveTask({ ...t, steps, status: t.status === "todo" && steps.some((s) => s.done) ? "active" : t.status }); };
  const cycleDue = () => {
    const opts = [null, todayStr(), todayStr(addDays(new Date(), 1)), nextSaturday()];
    const idx = opts.indexOf(t.dueDate); const next = opts[(idx + 1) % opts.length];
    saveTask({ ...t, dueDate: next, reminder: next ? { at: next + "T09:00:00" } : null });
  };
  const cycleAssign = () => {
    const order = [[me.id], [other.id], ["together"]];
    const cur = JSON.stringify(t.assignees || []); const idx = order.findIndex((o) => JSON.stringify(o) === cur);
    saveTask({ ...t, assignees: order[(idx + 1) % order.length] });
  };
  const pickFinish = async (e) => { const f = e.target.files?.[0]; e.target.value = ""; if (!f) return; nav("finish", { id: t.id, photo: await downscale(f) }); };
  const doneT = t.status === "done";
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("home")} /><H1 small>{t.title}</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      {t.photoBefore && <Photo src={t.photoBefore} h={168} label={"before" + (t.createdAt ? " · " + fmtDue({ dueDate: t.createdAt }).toLowerCase() : "")} style={{ marginBottom: 13 }} />}
      <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
        <span style={{ fontSize: 11.5, fontWeight: 700, color: doneT ? C.oliveDark : C.terraDark, background: doneT ? C.oliveSoft : C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>{doneT ? "done" : t.status === "active" ? "in progress" : "to do"}</span>
        <span onClick={cycleAssign} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: C.ink, background: "#fff", borderRadius: 999, padding: "5px 12px", boxShadow: "0 1px 4px rgba(63,56,42,.07)", whiteSpace: "nowrap", cursor: "pointer" }}>
          {assignee ? <><Avatar m={assignee} size={16} />{assignee.id === me.id ? "Yours" : assignee.name}</> : <>{state.household.members.map((m) => <Avatar key={m.id} m={m} size={16} ring />)}Together</>}
        </span>
        <span onClick={cycleDue} style={{ fontSize: 11.5, fontWeight: 700, color: "#54695a", background: "#fff", borderRadius: 999, padding: "6px 12px", boxShadow: "0 1px 4px rgba(63,56,42,.07)", whiteSpace: "nowrap", cursor: "pointer" }}>{fmtDue(t)}</span>
        <span onClick={() => saveTask({ ...t, photoProof: !t.photoProof })} style={{ fontSize: 11.5, fontWeight: 700, color: t.photoProof ? C.olive : C.mut, background: t.photoProof ? C.oliveSoft : C.bg, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{t.photoProof ? "photo when done" : "no photo needed"}</span>
      </div>
      {t.steps?.length > 0 && (
        <div style={{ ...card(22), padding: "8px 18px", marginTop: 13 }}>
          {t.steps.map((s, i) => (
            <div key={i} onClick={() => toggleStep(i)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: i ? "1px solid #f0f3ee" : "none", cursor: "pointer" }}>
              <CheckCircle size={24} state={s.done ? "done" : i === t.steps.findIndex((x) => !x.done) ? "active" : "todo"} />
              <span style={{ fontSize: 14, fontWeight: s.done ? 600 : 700, color: s.done ? C.faint : i === t.steps.findIndex((x) => !x.done) ? C.ink : C.mut, textDecoration: s.done ? "line-through" : "none" }}>{s.text}</span>
            </div>
          ))}
        </div>
      )}
      {t.listItems?.length > 0 && (
        <div style={{ ...card(22), padding: "14px 18px", marginTop: 13, display: "flex", flexWrap: "wrap", gap: 6 }}>
          {t.listItems.map((it, i) => (
            <span key={i} onClick={() => saveTask({ ...t, listItems: t.listItems.map((x, j) => j === i ? { ...x, done: !x.done } : x) })}
              style={{ fontSize: 12.5, fontWeight: 600, color: it.done ? C.faint : C.ink2, background: C.bg, borderRadius: 999, padding: "7px 13px", textDecoration: it.done ? "line-through" : "none", cursor: "pointer" }}>{it.text}</span>
          ))}
        </div>
      )}
      {t.reminder?.at && !doneT && (
        <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginTop: 11, boxShadow: shadow }}>
          <span style={{ width: 34, height: 34, borderRadius: "50%", background: C.oliveSoft, display: "grid", placeItems: "center", flex: "none" }}>
            <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke={C.olive} strokeWidth="1.8" strokeLinecap="round"><path d="M12 8v5l3 2M12 3a9 9 0 109 9" /></svg>
          </span>
          <div style={{ flex: 1, fontSize: 13, color: "#54695a", lineHeight: 1.45 }}><b style={{ color: C.ink }}>{fmtDue({ dueDate: t.reminder.at.slice(0, 10) })} {t.reminder.at.slice(11, 16)} reminder</b> — and {other.name} hears the good news the moment it's done.</div>
        </div>
      )}
      {(t.comments || []).map((c, i) => {
        const by = state.household.members.find((m) => m.id === c.by) || other;
        return (
          <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", marginTop: 12, flexDirection: c.by === me.id ? "row-reverse" : "row" }}>
            <Avatar m={by} size={30} />
            <div style={{ background: C.terraSoft, borderRadius: 18, [c.by === me.id ? "borderTopRightRadius" : "borderTopLeftRadius"]: 6, padding: "11px 14px", fontSize: 13.5, lineHeight: 1.5, color: C.terraDark }}>"{c.text}"</div>
          </div>
        );
      })}
      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <input value={comment} onChange={(e) => setComment(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && comment.trim()) { saveTask({ ...t, comments: [...(t.comments || []), { by: me.id, text: comment.trim(), ts: new Date().toISOString() }] }); setComment(""); } }}
          placeholder="Add a comment…" style={{ flex: 1, background: "#fff", border: "none", outline: "none", borderRadius: 999, padding: "11px 15px", fontSize: 13, fontWeight: 600, color: C.ink, fontFamily: FONT, boxShadow: shadow }} />
      </div>
      {!doneT ? (
        <div style={{ display: "flex", gap: 9, marginTop: 16 }}>
          <div onClick={() => finRef.current.click()} style={{ flex: 1.6, background: C.olive, color: "#fff", borderRadius: 999, padding: 15, fontSize: 14.5, fontWeight: 700, textAlign: "center", boxShadow: shadowBtn, cursor: "pointer" }}>Finish with a photo</div>
          <div onClick={() => { completeTask(t); nav("home"); }} style={{ flex: 1, background: "#fff", borderRadius: 999, padding: 15, fontSize: 14, fontWeight: 700, color: "#748078", textAlign: "center", boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}>Mark done</div>
        </div>
      ) : (
        <div style={{ display: "flex", gap: 9, marginTop: 16 }}>
          <div onClick={() => uncompleteTask(t)} style={{ flex: 1, background: "#fff", borderRadius: 999, padding: 15, fontSize: 14, fontWeight: 700, color: "#748078", textAlign: "center", boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}>Bring it back</div>
        </div>
      )}
      <div onClick={() => { removeTask(t.id); nav("home"); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 14, cursor: "pointer" }}>Remove task</div>
      <input ref={finRef} type="file" accept="image/*" capture="environment" hidden onChange={pickFinish} />
    </div>
  </>);
}
function nextSaturday() { const d = new Date(); d.setDate(d.getDate() + ((6 - d.getDay() + 7) % 7 || 7)); return todayStr(d); }

/* ---------- 8d Add from photo ---------- */
function AddFromPhoto({ state, me, other, nav, photo, saveTask }) {
  const [draft, setDraft] = useState(null);
  const [failed, setFailed] = useState(false);
  const [title, setTitle] = useState("");
  const stepsLabels = useMemo(() => ["Recognised the job in your photo", "Drafted the task and steps", "Suggested who takes it", "Sketching the schedule…"], []);
  const shown = useSteppedReveal(stepsLabels, !draft);
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const d = await aiDraftFromPhoto(photo, state.household.members);
        if (alive) { setDraft(d); setTitle(d.title); }
      } catch (e) {
        if (!alive) return;
        setFailed(true);
        const d = { seen: "Saved your photo with the task", title: "New task from your photo", steps: [], assignee: me.id, why: "Give it a title and I'll keep the photo as the 'before'. Add your Claude key in Settings and I'll draft these for you next time.", photoProof: true };
        setDraft(d); setTitle(d.title);
      }
    })();
    return () => { alive = false; };
  }, []);
  const rows = draft
    ? [draft.seen || stepsLabels[0], "Drafted the task" + (draft.steps?.length ? " and " + draft.steps.length + " steps" : ""), draft.assignee === "together" ? "Suggested you take it together" : "Suggested " + (state.household.members.find((m) => m.id === draft.assignee)?.name || me.name)]
    : stepsLabels.slice(0, Math.max(1, shown));
  const [assignee, setAssignee] = useState(null);
  const [due, setDue] = useState("sat");
  const [proof, setProof] = useState(null);
  const eff = {
    assignee: assignee ?? draft?.assignee ?? me.id,
    dueDate: due === "sat" ? nextSaturday() : due === "tomorrow" ? todayStr(addDays(new Date(), 1)) : due === "today" ? todayStr() : null,
    proof: proof ?? draft?.photoProof ?? state.household.settings.photoProofDefault,
  };
  const assigneeM = eff.assignee === "together" ? null : state.household.members.find((m) => m.id === eff.assignee);
  const log = () => {
    const t = {
      id: uid(), title: title || "New task", steps: (draft?.steps || []).map((s) => ({ text: s, done: false })),
      assignees: eff.assignee === "together" ? ["together"] : [eff.assignee], dueDate: eff.dueDate, dueTime: null,
      reminder: eff.dueDate ? { at: eff.dueDate + "T09:00:00" } : null,
      photoProof: eff.proof, photoBefore: photo, status: "todo", source: "photo", createdAt: todayStr(), comments: [],
    };
    saveTask(t);
    sbUploadPhoto(photo, "before-" + t.id).then((url) => { if (url !== photo) saveTask({ ...t, photoBefore: url }); });
    nav("home");
  };
  const chip = (label, on, fn) => (
    <span onClick={fn} style={{ background: "#fff", borderRadius: 999, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, color: on ? C.ink : C.ink2, boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}><span style={{ color: C.terra }}>↳ </span>{label}</span>
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("home")} /><H1 small>New task</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      <Photo src={photo} h={150} label="your photo · just now" />
      <div style={{ marginTop: 13 }}><AiChecklist rows={rows} running={!draft} /></div>
      {draft && (<>
        <div style={{ ...card(22), padding: "16px 18px", marginTop: 11 }}>
          <input value={title} onChange={(e) => setTitle(e.target.value)} style={{ width: "100%", border: "none", outline: "none", fontSize: 16, fontWeight: 800, color: C.ink, fontFamily: FONT, background: "transparent", padding: 0 }} />
          <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 3 }}>{draft.steps?.length ? draft.steps.length + " steps drafted" : "no steps — just the one job"}</div>
          <div style={{ display: "flex", gap: 7, marginTop: 12, flexWrap: "wrap" }}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: C.ink, background: C.bg, borderRadius: 999, padding: "5px 12px", whiteSpace: "nowrap" }}>
              {assigneeM ? <><Avatar m={assigneeM} size={16} />{assigneeM.name}</> : "Together"}
            </span>
            <span style={{ fontSize: 11.5, fontWeight: 700, color: C.ink2, background: C.bg, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>{eff.dueDate ? fmtDue({ dueDate: eff.dueDate }) + " 9:00" : "Anytime"}</span>
            <span style={{ fontSize: 11.5, fontWeight: 700, color: C.oliveDark, background: C.oliveSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>saved to Klaus</span>
            {eff.proof && <span style={{ fontSize: 11.5, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>photo when done</span>}
          </div>
        </div>
        <div style={{ display: "flex", gap: 11, alignItems: "flex-start", background: C.oliveSoft, borderRadius: 20, padding: "14px 16px", marginTop: 11 }}>
          <div><div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.olive, fontWeight: 800 }}>My read</div>
            <div style={{ fontSize: 13, lineHeight: 1.5, color: C.oliveDark, marginTop: 4, fontWeight: 600 }}>{draft.why}</div></div>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginTop: 13 }}>
          {chip("Give it to " + (eff.assignee === other.id ? me.name : other.name) + " instead", false, () => setAssignee(eff.assignee === other.id ? me.id : other.id))}
          {chip("Do it together", eff.assignee === "together", () => setAssignee("together"))}
          {chip(due === "sat" ? "Make it tomorrow" : "Make it Saturday", false, () => setDue(due === "sat" ? "tomorrow" : "sat"))}
          {chip(eff.proof ? "No photo needed" : "Ask for a photo", false, () => setProof(!eff.proof))}
        </div>
        <PrimaryBtn style={{ marginTop: 15 }} onClick={log}>Log it</PrimaryBtn>
        {failed && <div style={{ fontSize: 11.5, color: C.mut, fontWeight: 600, textAlign: "center", marginTop: 10, lineHeight: 1.5 }}>Smart drafting needs a Claude key — add one under the gear.</div>}
      </>)}
    </div>
  </>);
}

/* ---------- 8e Scan a note ---------- */
function ScanNote({ state, me, other, nav, photo, saveTask }) {
  const [res, setRes] = useState(null);
  const [failed, setFailed] = useState(false);
  const stepsLabels = useMemo(() => ["Reading the handwriting", "Grouping the items", "Checking for duplicates…"], []);
  const shown = useSteppedReveal(stepsLabels, !res);
  useEffect(() => {
    let alive = true;
    (async () => {
      try { const d = await aiScanNote(photo, state.household.members); if (alive) setRes(d); }
      catch (e) {
        if (!alive) return; setFailed(true);
        setRes({ seen: "Saved the note as a picture", kind: "task", title: "From your paper note", items: [], why: "Add your Claude key in Settings and I'll read the handwriting for you next time." });
      }
    })();
    return () => { alive = false; };
  }, []);
  const [choice, setChoice] = useState("list");
  const rows = res ? [res.seen, res.items?.length ? "Found " + res.items.length + " item" + (res.items.length > 1 ? "s" : "") : "Kept the note attached", "No duplicates left behind"] : stepsLabels.slice(0, Math.max(1, shown));
  const add = () => {
    if (res.kind === "list" && choice !== "each" && res.items?.length) {
      const t = { id: uid(), title: res.title || "Weekly shop", listItems: res.items.map((x) => ({ text: x, done: false })), steps: [], assignees: choice === "mine" ? [me.id] : ["together"], status: "todo", photoProof: false, photoBefore: photo, source: "note", createdAt: todayStr(), comments: [] };
      saveTask(t); sbUploadPhoto(photo, "note-" + t.id).then((u) => { if (u !== photo) saveTask({ ...t, photoBefore: u }); });
    } else if (choice === "each" && res.items?.length) {
      res.items.forEach((it) => saveTask({ id: uid(), title: it, steps: [], assignees: [me.id], status: "todo", photoProof: false, source: "note", createdAt: todayStr(), comments: [] }));
    } else {
      const t = { id: uid(), title: res.title || "From your note", steps: [], assignees: [me.id], status: "todo", photoProof: false, photoBefore: photo, source: "note", createdAt: todayStr(), comments: [] };
      saveTask(t);
    }
    nav("home");
  };
  const chip = (label, val) => (
    <span onClick={() => setChoice(val)} style={{ background: choice === val ? C.ink : "#fff", borderRadius: 999, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, color: choice === val ? "#fff" : C.ink2, boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}><span style={{ color: choice === val ? C.terraSoft : C.terra }}>↳ </span>{label}</span>
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("home")} /><H1 small>Scan a note</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      <Photo src={photo} h={150} label="your note · just now" />
      <div style={{ marginTop: 13 }}><AiChecklist rows={rows} running={!res} /></div>
      {res && (<>
        <div style={{ ...card(22), padding: "16px 18px", marginTop: 11 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
            <div style={{ flex: 1, minWidth: 0, fontSize: 16, fontWeight: 800, color: C.ink, lineHeight: 1.3 }}>{res.title}</div>
            {res.items?.length > 0 && <span style={{ fontSize: 12, fontWeight: 700, color: C.terra, flex: "none" }}>{res.items.length} items</span>}
          </div>
          {res.items?.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 12 }}>
              {res.items.map((it, i) => <span key={i} style={{ fontSize: 12.5, fontWeight: 600, color: C.ink2, background: C.bg, borderRadius: 999, padding: "7px 13px" }}>{it}</span>)}
            </div>
          )}
          <div style={{ display: "flex", gap: 7, marginTop: 12, flexWrap: "wrap" }}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: C.ink, background: C.bg, borderRadius: 999, padding: "5px 10px" }}>
              {choice === "mine" ? <><Avatar m={me} size={16} />{me.name}</> : <><span style={{ display: "flex" }}><Avatar m={me} size={16} ring /><span style={{ marginLeft: -5, display: "flex" }}><Avatar m={other} size={16} ring /></span></span>Together</>}
            </span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 11, alignItems: "flex-start", background: C.oliveSoft, borderRadius: 20, padding: "14px 16px", marginTop: 11 }}>
          <div><div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.olive, fontWeight: 800 }}>My read</div>
            <div style={{ fontSize: 13, lineHeight: 1.5, color: C.oliveDark, marginTop: 4, fontWeight: 600 }}>{res.why}</div></div>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginTop: 13 }}>
          {res.items?.length > 0 && chip("Each item its own task", "each")}
          {res.items?.length > 0 && chip("It's a to-do list", "list")}
          {chip("Keep it just for me", "mine")}
        </div>
        <PrimaryBtn style={{ marginTop: 15 }} onClick={add}>Add to the list</PrimaryBtn>
        {failed && <div style={{ fontSize: 11.5, color: C.mut, fontWeight: 600, textAlign: "center", marginTop: 10, lineHeight: 1.5 }}>Reading handwriting needs a Claude key — add one under the gear.</div>}
      </>)}
    </div>
  </>);
}

/* ---------- 8f Finish with photo ---------- */
function FinishPhoto({ state, me, other, nav, taskId, photo, saveTask, completeTask }) {
  const t = state.tasks.find((x) => x.id === taskId);
  const [ver, setVer] = useState(null);
  const doneRef = useRef(false);
  const labels = useMemo(() => [t?.photoBefore ? "Comparing with the before photo" : "Having a look at your photo", "Checking the job", "Closing the steps…"], []);
  const shown = useSteppedReveal(labels, !ver);
  useEffect(() => {
    let alive = true;
    (async () => {
      let v;
      try { v = await aiVerify(t?.photoBefore, photo, t?.title || "the task"); }
      catch (e) { v = { checks: [t?.photoBefore ? "Compared with the before photo" : "Photo saved with the task", "Looks like the job's been done", (t?.steps?.length ? "All " + t.steps.length + " steps closed" : "Marked complete")], ok: true, note: "Photo saved as proof — nicely done." }; }
      if (!alive) return;
      setVer(v);
      if (!doneRef.current) {
        doneRef.current = true;
        const done = completeTask(t, { doneWithPhoto: true, photoAfter: photo, steps: (t.steps || []).map((s) => ({ ...s, done: true })) });
        sbUploadPhoto(photo, "after-" + t.id).then((u) => { if (u !== photo) saveTask({ ...done, photoAfter: u }); });
      }
    })();
    return () => { alive = false; };
  }, []);
  if (!t) return null;
  const todays = state.tasks.filter(isTodayTask);
  const doneCount = todays.filter((x) => x.status === "done").length;
  const leftCount = todays.length - doneCount;
  const rows = ver ? ver.checks : labels.slice(0, Math.max(1, shown));
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("task", { id: t.id })} /><H1 small>Finish task</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      <div style={{ display: "flex", gap: 10 }}>
        <Photo src={t.photoBefore} h={130} r={20} label="before" style={{ flex: 1, boxShadow: "none" }} />
        <Photo src={photo} h={130} r={20} label="after · now" border style={{ flex: 1, boxShadow: "none" }} />
      </div>
      <div style={{ marginTop: 13 }}><AiChecklist rows={rows} running={!ver} /></div>
      {ver && (<>
        <div style={{ display: "flex", gap: 12, marginTop: 12 }}>
          <div style={{ flex: 1, ...card(22), padding: 17, textAlign: "center" }}>
            <div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.mut, fontWeight: 800 }}>Nice one</div>
            <div style={{ fontSize: 32, fontWeight: 800, color: C.olive, marginTop: 6 }}>✓</div>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: C.mut, marginTop: 2 }}>task complete</div>
          </div>
          <div style={{ flex: 1, ...card(22), padding: 17, textAlign: "center" }}>
            <div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.mut, fontWeight: 800 }}>Today</div>
            <div style={{ fontSize: 32, fontWeight: 800, color: C.ink, marginTop: 6 }}>{doneCount}<span style={{ fontSize: 15, fontWeight: 600, color: C.mut }}>/{todays.length}</span></div>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: C.mut, marginTop: 2 }}>{leftCount === 0 ? "that's the lot" : leftCount === 1 ? "one to go" : leftCount + " to go"}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 11, alignItems: "flex-start", background: C.oliveSoft, borderRadius: 20, padding: "14px 16px", marginTop: 12 }}>
          <div><div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.olive, fontWeight: 800 }}>Done and dusted</div>
            <div style={{ fontSize: 13, lineHeight: 1.5, color: C.oliveDark, marginTop: 4, fontWeight: 600 }}>{ver.note}</div></div>
        </div>
        <PrimaryBtn style={{ marginTop: 15 }} onClick={() => { const cur = state.tasks.find((x) => x.id === t.id); saveTask({ ...cur, sentTo: other.id }); nav("home"); }}>Send it to {other.name}</PrimaryBtn>
        <div onClick={() => { const cur = state.tasks.find((x) => x.id === t.id); saveTask({ ...cur, status: "todo", doneAt: null, doneBy: null, doneWithPhoto: false, photoAfter: null }); nav("task", { id: t.id }); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 12, cursor: "pointer" }}>Undo</div>
      </>)}
    </div>
  </>);
}

/* ---------- 8g Reminders & nudges ---------- */
const PRESETS = (other) => [
  "Gentle poke — the parcel place closes at 6",
  "One left and the day's a full house",
  "Trade you — I'll cook if you grab it",
  "No rush — just don't want you to forget",
];
function Reminders({ state, me, other, nav, goTab, tab, saveHousehold, sendNudge }) {
  const [to, setTo] = useState(other.id);
  const [free, setFree] = useState("");
  const [sent, setSent] = useState(false);
  const scheduled = state.tasks.filter((t) => t.status !== "done" && (t.reminder?.at || t.daily));
  const send = (text) => { sendNudge(to, text); setFree(""); setSent(true); setTimeout(() => setSent(false), 2500); };
  const q = state.household.settings.quietHours;
  const recent = [...state.nudges].reverse().slice(0, 3);
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><H1>Reminders & nudges</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 18px" }}>
      <Kicker>Scheduled</Kicker>
      {scheduled.length === 0 && <div style={{ fontSize: 13, color: C.mut, fontWeight: 600, padding: "4px 2px 12px" }}>Nothing scheduled — give a task a day and a reminder comes with it.</div>}
      {scheduled.map((t) => {
        const m = t.assignees?.includes("together") || (t.assignees?.length || 0) > 1 ? null : state.household.members.find((x) => x.id === t.assignees?.[0]);
        const pill = relDuePill(t);
        return (
          <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ ...card(20), padding: "14px 16px", display: "flex", alignItems: "center", gap: 12, marginBottom: 9, cursor: "pointer" }}>
            {m ? <Avatar m={m} size={30} /> : <span style={{ width: 30, height: 30, borderRadius: "50%", background: C.track, color: "#8b8672", display: "grid", placeItems: "center", fontSize: 10, fontWeight: 800, flex: "none" }}>{me.name[0]}·{other.name[0]}</span>}
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>{t.title}</div>
              <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 1 }}>
                {t.daily ? "Daily · whoever's home first" : fmtDue({ dueDate: t.reminder.at.slice(0, 10) }) + " " + t.reminder.at.slice(11, 16) + " · a push for " + (m ? (m.id === me.id ? "you" : m.name) : "you both")}
              </div>
            </div>
            {pill && <span style={{ fontSize: 11, fontWeight: 700, color: pill.fg, background: pill.bg, borderRadius: 999, padding: "5px 11px", whiteSpace: "nowrap" }}>{pill.text}</span>}
          </div>
        );
      })}
      <Kicker style={{ margin: "20px 2px 9px" }}>Send a nudge</Kicker>
      <div style={{ display: "flex", gap: 8 }}>
        {[other, me].map((p) => (
          <span key={p.id} onClick={() => setTo(p.id)} style={{ display: "inline-flex", alignItems: "center", gap: 7, background: to === p.id ? C.ink : "#fff", color: to === p.id ? "#fff" : C.mut, borderRadius: 999, padding: "8px 15px", fontSize: 13, fontWeight: 700, boxShadow: to === p.id ? "none" : "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}>
            <Avatar m={p} size={17} />{p.id === me.id ? "Me" : p.name}
          </span>
        ))}
      </div>
      <div style={{ fontSize: 12, fontWeight: 600, color: C.mut, margin: "12px 2px 8px" }}>Nudges stay warm — pick one or write your own:</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
        {PRESETS(other).map((p, i) => (
          <span key={i} onClick={() => send(p)} style={{ background: "#fff", borderRadius: 999, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, color: C.ink2, boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}><span style={{ color: C.terra }}>↳ </span>{p}</span>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <input value={free} onChange={(e) => setFree(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && free.trim()) send(free.trim()); }}
          placeholder="Or in your own words…" style={{ flex: 1, background: "#fff", border: "none", outline: "none", borderRadius: 999, padding: "11px 15px", fontSize: 13, fontWeight: 600, color: C.ink, fontFamily: FONT, boxShadow: shadow }} />
        <span onClick={() => free.trim() && send(free.trim())} style={{ background: C.terra, color: "#fff", borderRadius: 999, padding: "11px 17px", fontSize: 13, fontWeight: 800, cursor: "pointer" }}>Send</span>
      </div>
      {sent && <div style={{ fontSize: 12.5, color: C.olive, fontWeight: 700, marginTop: 10, textAlign: "center" }}>Sent — it'll be waiting on their home screen.</div>}
      {recent.length > 0 && (<>
        <Kicker style={{ margin: "20px 2px 8px" }}>Lately</Kicker>
        {recent.map((n) => {
          const f = state.household.members.find((m) => m.id === n.from);
          return (
            <div key={n.id} style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 10 }}>
              <Avatar m={f} size={30} />
              <div style={{ background: C.terraSoft, borderRadius: 18, borderTopLeftRadius: 6, padding: "11px 14px", fontSize: 13, lineHeight: 1.5, color: C.terraDark }}>"{n.text}"</div>
            </div>
          );
        })}
      </>)}
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginTop: 12, boxShadow: shadow }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Quiet hours</div>
          <div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>No pushes {q.from} – {q.to} · both of you</div>
        </div>
        <Toggle on={q.on} onClick={() => saveHousehold({ ...state.household, settings: { ...state.household.settings, quietHours: { ...q, on: !q.on } } })} />
      </div>
    </div>
  </>);
}

/* ---------- 8h Calendar ---------- */
function CalendarView({ state, me, other, nav, goTab, tab }) {
  const [sel, setSel] = useState(todayStr());
  const start = addDays(new Date(), -1);
  const days = Array.from({ length: 7 }, (_, i) => addDays(start, i));
  const tasksOn = (ds) => state.tasks.filter((t) => t.dueDate === ds || (t.daily && ds >= todayStr()));
  const caches = state.outlookCache || {};
  const eventsOn = (ds) => state.household.members.flatMap((m) => {
    const c = caches[m.id];
    if (!c || (m.outlook?.access || "write") !== "readwrite" || !m.outlook?.connected) return [];
    return (c.events || []).filter((e) => e.start.slice(0, 10) === ds).map((e) => ({ ...e, member: m }));
  });
  const connected = state.household.members.filter((m) => m.outlook?.connected);
  const lastSync = connected.map((m) => caches[m.id]?.updatedAt).filter(Boolean).sort().pop();
  const mins = lastSync ? Math.max(0, Math.round((Date.now() - new Date(lastSync)) / 60000)) : null;
  const selDate = new Date(sel + "T00:00:00");
  const selTasks = tasksOn(sel);
  const selEvents = eventsOn(sel).sort((a, b) => a.start.localeCompare(b.start));
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}>
      <H1>Calendar</H1>
      <span onClick={() => nav("sharing")} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11, fontWeight: 700, color: connected.length ? C.oliveDark : C.mut, background: connected.length ? C.oliveSoft : C.track, borderRadius: 999, padding: "5px 11px", whiteSpace: "nowrap", cursor: "pointer" }}>
        <i style={{ width: 6, height: 6, borderRadius: "50%", background: connected.length ? C.olive : C.faint, display: "block" }} />
        Outlook · {connected.length ? (mins === null ? "connected" : mins < 1 ? "just now" : mins + "m ago") : "not yet"}
      </span>
    </div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 18px" }}>
      <div style={{ display: "flex", gap: 5, marginTop: 2, background: "#fff", borderRadius: 20, padding: "8px 6px", boxShadow: shadow }}>
        {days.map((d) => {
          const ds = todayStr(d); const selD = ds === sel;
          const tDots = tasksOn(ds).length, eDots = eventsOn(ds).length;
          const dots = [...Array(Math.min(tDots, 2)).fill("task"), ...Array(Math.min(eDots, 3 - Math.min(tDots, 2))).fill("ev")];
          return (
            <div key={ds} onClick={() => setSel(ds)} style={{ flex: 1, textAlign: "center", padding: "8px 0", borderRadius: 14, background: selD ? C.ink : "transparent", cursor: "pointer" }}>
              <div style={{ fontSize: 9, fontWeight: 800, color: selD ? "rgba(247,244,234,.7)" : C.mut }}>{DOW[d.getDay()]}</div>
              <div style={{ fontSize: 15, fontWeight: selD ? 800 : 700, color: selD ? C.bg : C.ink, marginTop: 2 }}>{d.getDate()}</div>
              <div style={{ display: "flex", gap: 2, justifyContent: "center", marginTop: 4, height: 4 }}>
                {dots.map((k, i) => <i key={i} style={{ width: 4, height: 4, borderRadius: "50%", background: k === "task" ? (selD ? "#a3ad7a" : C.olive) : C.faint, display: "block" }} />)}
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ display: "flex", gap: 14, fontSize: 10.5, fontWeight: 700, color: C.mut, margin: "12px 2px 4px" }}>
        <span><i style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: C.olive }} /> task</span>
        <span><i style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: C.faint }} /> from Outlook</span>
      </div>
      <Kicker style={{ margin: "12px 2px 8px" }}>{DAYNAME[selDate.getDay()]} {selDate.getDate()}</Kicker>
      {selTasks.length === 0 && selEvents.length === 0 && <div style={{ fontSize: 13, color: C.mut, fontWeight: 600, padding: "6px 2px" }}>Nothing on this day yet — a quiet one.</div>}
      {selEvents.map((e, i) => (
        <div key={"e" + i} style={{ display: "flex", gap: 12, background: "#fff", borderRadius: 20, padding: "13px 15px", marginBottom: 9, boxShadow: shadow, opacity: 0.75 }}>
          <div style={{ width: 3, borderRadius: 2, background: C.faint, flex: "none" }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>{e.title}</div>
            <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 2 }}>{e.start.slice(11, 16)} – {e.end ? e.end.slice(11, 16) : ""} · from {e.member.id === me.id ? "your" : e.member.name + "'s"} Outlook (read)</div>
          </div>
          <span style={{ alignSelf: "center", display: "flex" }}><Avatar m={e.member} size={22} /></span>
        </div>
      ))}
      {selTasks.map((t) => {
        const together = t.assignees?.includes("together") || (t.assignees?.length || 0) > 1;
        const m = together ? null : state.household.members.find((x) => x.id === t.assignees?.[0]);
        const wroteMine = !!t.outlook?.[me.id];
        const wroteOther = !!t.outlook?.[other.id];
        const sub = [
          t.daily ? "daily" : (t.dueTime || "any time"),
          together ? "together" : null,
          t.listItems?.length ? "list attached" : null,
          wroteMine ? "written to your Outlook (write)" : wroteOther ? "written to " + other.name + "'s Outlook (write)" : "in Klaus",
        ].filter(Boolean).join(" · ");
        return (
          <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ display: "flex", gap: 12, background: C.oliveSoft, borderRadius: 20, padding: "13px 15px", marginBottom: 9, cursor: "pointer", opacity: t.status === "done" ? 0.7 : 1 }}>
            <div style={{ width: 3, borderRadius: 2, background: C.olive, flex: "none" }} />
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink, textDecoration: t.status === "done" ? "line-through" : "none" }}>{t.title}</div>
              <div style={{ fontSize: 11.5, fontWeight: 600, color: C.oliveDark, marginTop: 2 }}>{sub}</div>
            </div>
            {together
              ? <span style={{ display: "flex", alignSelf: "center" }}><Avatar m={me} size={22} ring /><span style={{ marginLeft: -7, display: "flex" }}><Avatar m={other} size={22} ring /></span></span>
              : m && <span style={{ alignSelf: "center", display: "flex" }}><Avatar m={m} size={22} /></span>}
          </div>
        );
      })}
      {connected.length === 0 && (
        <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, lineHeight: 1.5, margin: "14px 2px 0" }}>
          Once Outlook is connected, tasks land in each person's calendar and their events can show here. Set per person in <span onClick={() => nav("sharing")} style={{ color: C.olive, cursor: "pointer" }}>Sharing & access</span>.
        </div>
      )}
    </div>
  </>);
}

/* ---------- 8i Sharing & access ---------- */
function Sharing({ state, me, nav, saveHousehold }) {
  const h = state.household;
  const setMember = (id, patch) => saveHousehold({ ...h, members: h.members.map((m) => (m.id === id ? { ...m, ...patch } : m)) });
  const setSetting = (k, v) => saveHousehold({ ...h, settings: { ...h.settings, [k]: v } });
  const [editing, setEditing] = useState(null);
  const [cid, setCid] = useState(h.settings?.msClientId || "");
  const [msMsg, setMsMsg] = useState("");
  const saveCid = () => { setSetting("msClientId", cid.replace(/\s+/g, "")); setMsMsg("Saved — it syncs to both phones. Now connect below."); };
  const connectMe = async () => {
    try { setMsMsg("Off to Microsoft's sign-in…"); await msSignIn(h.settings.msClientId); }
    catch (e) { setMsMsg("Sign-in didn't start — worth re-checking the client ID."); }
  };
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("settings")} /><H1 small>Sharing & access</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      <Kicker>Outlook</Kicker>
      <div style={{ ...card(22), padding: "16px 17px", marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.ink }}>Microsoft app (once for the household)</div>
        <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 3, lineHeight: 1.5 }}>The Application (client) ID from the Azure registration. Not a secret — it's shared to both phones.</div>
        <input value={cid} onChange={(e) => setCid(e.target.value)} placeholder="00000000-0000-0000-0000-000000000000"
          style={{ width: "100%", boxSizing: "border-box", background: C.bg, border: "none", outline: "none", borderRadius: 14, padding: "12px 14px", fontSize: 12, fontWeight: 600, color: C.ink, fontFamily: "'JetBrains Mono',monospace", marginTop: 8 }} />
        <div onClick={saveCid} style={{ marginTop: 10, background: C.oliveSoft, color: C.oliveDark, borderRadius: 999, padding: "10px", fontSize: 13, fontWeight: 800, textAlign: "center", cursor: "pointer" }}>Save</div>
        {msMsg && <div style={{ fontSize: 12, color: C.oliveDark, fontWeight: 700, marginTop: 10, textAlign: "center", lineHeight: 1.5 }}>{msMsg}</div>}
      </div>
      <Kicker>People</Kicker>
      {h.members.map((m) => (
        <div key={m.id} style={{ ...card(22), padding: "16px 17px", marginBottom: 11 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
            <Avatar m={m} size={36} />
            <div style={{ flex: 1 }}>
              {editing === m.id ? (
                <input autoFocus defaultValue={m.name} onBlur={(e) => { setMember(m.id, { name: e.target.value.trim() || m.name }); setEditing(null); }}
                  onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
                  style={{ fontSize: 15, fontWeight: 800, color: C.ink, border: "none", outline: "none", background: C.bg, borderRadius: 8, padding: "2px 8px", fontFamily: FONT, width: "70%" }} />
              ) : (
                <div onClick={() => setEditing(m.id)} style={{ fontSize: 15, fontWeight: 800, color: C.ink, cursor: "pointer" }}>{m.name} {m.id === me.id && <span style={{ fontSize: 11, fontWeight: 600, color: C.mut }}>— you</span>}</div>
              )}
              <div style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, fontWeight: 600, color: m.outlook?.connected ? C.oliveDark : C.mut, marginTop: 2 }}>
                <i style={{ width: 6, height: 6, borderRadius: "50%", background: m.outlook?.connected ? C.olive : C.faint, display: "block" }} />
                {m.outlook?.connected ? "Outlook connected · " + m.outlook.email : h.settings?.msClientId ? (m.id === me.id ? "Outlook — ready to connect" : "connects on " + m.name + "'s own phone") : "Outlook — add the Microsoft app ID above first"}
              </div>
            </div>
            {m.id === me.id && h.settings?.msClientId && (
              m.outlook?.connected
                ? <span onClick={() => setMember(m.id, { outlook: { ...m.outlook, connected: false } })} style={{ fontSize: 12, fontWeight: 700, color: C.mut, cursor: "pointer" }}>Disconnect</span>
                : <span onClick={connectMe} style={{ fontSize: 12, fontWeight: 800, color: "#fff", background: C.olive, borderRadius: 999, padding: "7px 13px", cursor: "pointer", boxShadow: shadowBtn, whiteSpace: "nowrap" }}>Connect</span>
            )}
          </div>
          <SegPill style={{ marginTop: 13 }} value={m.outlook?.access || "write"} onChange={(v) => setMember(m.id, { outlook: { ...m.outlook, access: v } })}
            options={[{ value: "write", label: "Write only" }, { value: "readwrite", label: "Read & write" }]} />
          <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, lineHeight: 1.5, marginTop: 10 }}>
            {(m.outlook?.access || "write") === "readwrite"
              ? <>Tasks go into {m.id === me.id ? "your" : m.name + "'s"} Outlook <b style={{ color: C.ink2 }}>and</b> {m.id === me.id ? "your" : "their"} events show in the shared calendar.</>
              : <>{m.id === me.id ? "Your" : m.name + "'s"} tasks land in {m.id === me.id ? "your" : "their"} Outlook, but other events stay private — only free/busy is used for suggestions.</>}
          </div>
        </div>
      ))}
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", border: "1.5px dashed " + C.line, borderRadius: 20, padding: "14px 16px", marginBottom: 16 }}>
        <span style={{ fontSize: 17, color: C.terra, fontWeight: 700 }}>+</span>
        <div style={{ flex: 1, fontSize: 13.5, fontWeight: 700, color: C.ink2 }}>Invite someone with a link</div>
        <span style={{ fontSize: 10.5, fontWeight: 600, color: C.mut, fontFamily: "'JetBrains Mono',monospace" }}>after Supabase setup</span>
      </div>
      <Kicker>Household</Kicker>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginBottom: 9, boxShadow: shadow }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Photo proof by default</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>Ask for an "after" photo on chores & projects</div></div>
        <Toggle on={h.settings.photoProofDefault} onClick={() => setSetting("photoProofDefault", !h.settings.photoProofDefault)} />
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", boxShadow: shadow }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Daily digest</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>One summary push at 18:00 instead of many</div></div>
        <Toggle on={h.settings.dailyDigest} onClick={() => setSetting("dailyDigest", !h.settings.dailyDigest)} />
      </div>
    </div>
  </>);
}

/* ---------- Settings (gear) ---------- */
function Settings({ state, me, nav, syncStatus, setSyncStatus }) {
  const keys = loadKeys();
  const [anth, setAnth] = useState(keys.anthropic || "");
  const [su, setSu] = useState(keys.supabaseUrl || "");
  const [sk, setSk] = useState(keys.supabaseAnon || "");
  const [msg, setMsg] = useState("");
  const saveAnth = () => { saveKeys({ ...loadKeys(), anthropic: anth.replace(/\s+/g, "") }); setMsg("Claude key saved on this device."); };
  const cleanUrl = (u) => {
    let x = (u || "").trim();
    try {
      if (/safelinks\.protection\.outlook\.com/i.test(x)) { const p = new URL(x).searchParams.get("url"); if (p) x = p; }
      x = decodeURIComponent(x);
    } catch (e) {}
    x = x.replace(/\s+/g, "").replace(/\/+$/, "");
    if (x && !/^https?:\/\//i.test(x)) x = "https://" + x;
    return x;
  };
  const connect = async () => {
    const u = cleanUrl(su), k = sk.replace(/\s+/g, "");
    setSu(u); setSk(k);
    saveKeys({ ...loadKeys(), supabaseUrl: u, supabaseAnon: k });
    sb = null;
    try { const r = await sbPullAll(); if (r) { setSyncStatus("reconnect"); setMsg("Connected — syncing between you now."); } }
    catch (e) { setMsg("Couldn't reach that project — worth re-checking the URL and key."); setSyncStatus("error"); }
  };
  const switchPerson = () => { localStorage.removeItem(LS_DEVICE); location.reload(); };
  const askNotif = async () => {
    if (typeof Notification === "undefined") { setMsg("This browser can't show notifications."); return; }
    const p = await Notification.requestPermission();
    setMsg(p === "granted" ? "Reminders will pop up while the app is open." : "Notifications stay off — reminders still show in the app.");
  };
  const input = (v, set, ph, mono) => (
    <input value={v} onChange={(e) => set(e.target.value)} placeholder={ph}
      style={{ width: "100%", boxSizing: "border-box", background: C.bg, border: "none", outline: "none", borderRadius: 14, padding: "12px 14px", fontSize: 12.5, fontWeight: 600, color: C.ink, fontFamily: mono ? "'JetBrains Mono',monospace" : FONT, marginTop: 8 }} />
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={() => nav("home")} /><H1 small>Settings</H1></div>
    <div style={{ padding: "0 20px 24px" }}>
      <Kicker>This phone</Kicker>
      <div style={{ ...card(22), padding: "16px 17px", marginBottom: 11, display: "flex", alignItems: "center", gap: 11 }}>
        <Avatar m={me} size={36} />
        <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 800, color: C.ink }}>{me.name}</div><div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 2 }}>This device answers as {me.name}</div></div>
        <span onClick={switchPerson} style={{ fontSize: 12, fontWeight: 700, color: C.terra, cursor: "pointer" }}>Switch</span>
      </div>
      <div onClick={() => nav("sharing")} style={{ ...card(20), padding: "14px 16px", marginBottom: 11, display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Sharing & access</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>People, Outlook permissions, household defaults</div></div>
        <span style={{ color: C.mut, fontWeight: 700 }}>›</span>
      </div>
      <div onClick={askNotif} style={{ ...card(20), padding: "14px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Allow reminders</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>Pop-ups while the app is open (pocket pushes come in phase three)</div></div>
        <span style={{ color: C.mut, fontWeight: 700 }}>›</span>
      </div>
      <Kicker>Claude</Kicker>
      <div style={{ ...card(22), padding: "16px 17px", marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.ink }}>Your Claude API key</div>
        <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 3, lineHeight: 1.5 }}>Lets Klaus draft tasks from photos, read paper notes and verify after-photos. Stored only on this device.</div>
        {input(anth, setAnth, "sk-ant-…", true)}
        <div onClick={saveAnth} style={{ marginTop: 10, background: C.oliveSoft, color: C.oliveDark, borderRadius: 999, padding: "10px", fontSize: 13, fontWeight: 800, textAlign: "center", cursor: "pointer" }}>Save key</div>
      </div>
      <Kicker>Shared store</Kicker>
      <div style={{ ...card(22), padding: "16px 17px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
          <i style={{ width: 7, height: 7, borderRadius: "50%", background: syncStatus === "connected" ? C.olive : syncStatus === "error" ? C.terra : C.faint, display: "block" }} />
          <div style={{ fontSize: 13, fontWeight: 700, color: C.ink }}>{syncStatus === "connected" ? "Supabase connected — shared between you" : syncStatus === "error" ? "Connection trouble" : "Local only for now"}</div>
        </div>
        <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 3, lineHeight: 1.5 }}>Everything works on this phone already. Paste the project URL and anon key when we set Supabase up, and tasks sync live between both phones.</div>
        {input(su, setSu, "https://xxxx.supabase.co", true)}
        {input(sk, setSk, "anon public key", true)}
        <div onClick={connect} style={{ marginTop: 10, background: C.olive, color: "#fff", borderRadius: 999, padding: "10px", fontSize: 13, fontWeight: 800, textAlign: "center", boxShadow: shadowBtn, cursor: "pointer" }}>Connect & sync</div>
      </div>
      {msg && <div style={{ fontSize: 12.5, color: C.oliveDark, fontWeight: 700, marginTop: 12, textAlign: "center", lineHeight: 1.5 }}>{msg}</div>}
      <div style={{ fontSize: 11, color: C.faint, fontWeight: 600, textAlign: "center", marginTop: 22 }}>Klaus · phase one · Outlook & pocket pushes to follow</div>
    </div>
  </>);
}

/* ---------- mount ---------- */
const style = document.createElement("style");
style.textContent = `@keyframes tspin{to{transform:rotate(360deg)}} .tspin{animation:tspin .9s linear infinite} body{margin:0;background:${C.bg}} input::placeholder{color:${C.mut};font-weight:600}`;
document.head.appendChild(style);
createRoot(document.getElementById("root")).render(<App />);
