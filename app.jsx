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
function shortDay(ds) {
  const d = new Date(ds + "T00:00:00"); const today = new Date(todayStr() + "T00:00:00");
  const diff = Math.round((d - today) / 86400000);
  if (diff === 0) return "today"; if (diff === 1) return "tomorrow";
  if (diff > 1 && diff < 7) return DAYNAME[d.getDay()];
  return d.getDate() + " " + d.toLocaleString("en-GB", { month: "short" });
}
function relDuePill(t) {
  if (t.daily) return { text: (t.streak || 0) + (t.streak === 1 ? " day" : " days"), fg: C.oliveDark, bg: C.oliveSoft };
  const ref = t.deadline || t.dueDate;
  if (!ref) return null;
  const diff = Math.round((new Date(ref) - new Date(todayStr())) / 86400000);
  if (t.deadline) {
    const dl = new Date(t.deadline + "T" + (t.deadlineTime || "23:59") + ":00");
    if (dl < new Date()) return { text: "overdue", fg: C.terraDark, bg: C.terraSoft };
    if (diff === 0) return { text: "due " + (t.deadlineTime ? "by " + t.deadlineTime : "today"), fg: C.terraDark, bg: C.terraSoft };
    if (diff <= 3) return { text: "due in " + diff + (diff === 1 ? " day" : " days"), fg: C.terraDark, bg: C.terraSoft };
    return { text: "by " + shortDay(t.deadline), fg: C.ink2, bg: C.bg };
  }
  if (diff < 0) return { text: "overdue", fg: C.terraDark, bg: C.terraSoft };
  if (diff === 0) return { text: "today", fg: C.terraDark, bg: C.terraSoft };
  return { text: "in " + diff + (diff === 1 ? " day" : " days"), fg: C.ink2, bg: C.bg };
}

function deadlineLabel(t) {
  if (!t.deadline) return null;
  if (t.deadline === t.dueDate && !t.deadlineTime) return null;
  return "due by " + shortDay(t.deadline) + (t.deadlineTime ? " " + t.deadlineTime : "");
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
    { id: "seed-water-plants", seed: true, title: "Water the plants", daily: true, streak: 0, assignees: ["tobias", "an"], steps: [], status: "todo", photoProof: false, source: "typed", createdAt: t, comments: [] },
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
async function aiVerify(beforeUrl, afterUrl, title) {
  const content = [];
  if (beforeUrl && beforeUrl.startsWith("data:")) content.push(imgBlock(beforeUrl));
  if (afterUrl && afterUrl.startsWith("data:")) content.push(imgBlock(afterUrl));
  content.push({ type: "text", text: `Task: "${title}". ${beforeUrl ? "First image is before, second is after." : "Image is the after photo."} Reply ONLY with JSON: {"checks":["3 short positive verification observations"],"ok":true|false,"note":"one warm closing sentence"}` });
  const txt = await claude([{ role: "user", content }]);
  return asJson(txt);
}

/* ---------- natural language timing & shopping ---------- */
function dayFromWord(w) {
  w = w.toLowerCase();
  if (w === "today" || w === "tonight") return todayStr();
  if (w === "tomorrow") return todayStr(addDays(new Date(), 1));
  const i = DAYNAME.findIndex((d) => d.toLowerCase() === w);
  if (i >= 0) { const d = new Date(); const diff = (i - d.getDay() + 7) % 7 || 7; return todayStr(addDays(d, diff)); }
  return null;
}
function localParse(text) {
  const raw = text.trim(); const low = " " + raw.toLowerCase() + " ";
  const out = { title: raw, on: null, time: null, deadline: null, deadlineTime: null, shopping: false, items: [] };
  const dayRe = /\b(today|tonight|tomorrow|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i;
  const dm = low.match(dayRe); const dstr = dm ? dayFromWord(dm[1]) : null;
  let hm = null;
  const tm = low.match(/\b(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/) || low.match(/\b(?:at\s+)(\d{1,2})(?::(\d{2}))?\b/) || low.match(/\b(\d{1,2}):(\d{2})\b/);
  if (tm) {
    let h = +tm[1], mi = +(tm[2] || 0);
    const ap = tm[3];
    if (ap === "pm" && h < 12) h += 12; if (ap === "am" && h === 12) h = 0;
    if (h >= 0 && h < 24 && mi < 60) hm = String(h).padStart(2, "0") + ":" + String(mi).padStart(2, "0");
  }
  if (/\bby\b/.test(low)) { out.deadline = dstr || todayStr(); out.deadlineTime = hm; }
  else if (dstr) { out.on = dstr; out.time = hm; }
  else if (hm) { out.on = todayStr(); out.time = hm; }
  if (/\b(shopping|groceries|grocery)\b/.test(low) || /^\s*buy\b/i.test(raw)) {
    out.shopping = true;
    const c = raw.split(":");
    if (c.length > 1) out.items = c.slice(1).join(":").split(/,|\band\b/i).map((x) => x.trim()).filter(Boolean);
    else if (/^\s*buy\s+/i.test(raw)) out.items = raw.replace(/^\s*buy\s+/i, "").split(/,|\band\b/i).map((x) => x.trim()).filter(Boolean);
  }
  let title = raw.replace(/\bby\b[^,;]*$/i, "").replace(dayRe, "").replace(/\bat\s+\d{1,2}(:\d{2})?\s*(am|pm)?\b/gi, "").replace(/\b\d{1,2}(:\d{2})?\s*(am|pm)\b/gi, "").replace(/\s{2,}/g, " ").replace(/[\s,:;-]+$/, "").trim();
  if (out.shopping) title = "Shopping";
  out.title = title || raw;
  return out;
}
async function aiParseTyped(text) {
  const txt = await claude([{ role: "user", content: [{ type: "text", text:
    `Parse this quick household task entry. Today is ${todayStr()} (${DAYNAME[new Date().getDay()]}). Reply ONLY JSON: {"title":"clean task title without timing words","on":"YYYY-MM-DD or null - the day to DO it","time":"HH:MM 24h or null - time of day to do it","deadline":"YYYY-MM-DD or null - a due-by date","deadlineTime":"HH:MM or null","shopping":true|false - true if this is shopping/groceries OR a list of food items,"items":["shopping items with quantities, e.g. 2 apples"]}. Rules: "by <time or day>" means deadline, not the day of action. A day named without "by" is the day to do it. If shopping, title is "Shopping". Input: ${JSON.stringify(text)}` }] }], 300);
  return asJson(txt);
}
async function parseTyped(text) {
  const lp = localParse(text);
  if (!loadKeys().anthropic) return lp;
  try {
    const ai = await Promise.race([aiParseTyped(text), new Promise((_, rj) => setTimeout(() => rj(new Error("timeout")), 7000))]);
    return { title: ai.title || lp.title, on: ai.on ?? null, time: ai.time ?? null, deadline: ai.deadline ?? null, deadlineTime: ai.deadlineTime ?? null, shopping: !!ai.shopping, items: Array.isArray(ai.items) ? ai.items : [] };
  } catch (e) { return lp; }
}
async function aiSmartPhoto(dataUrl, members) {
  const names = members.map((m) => m.name).join(" and ");
  const txt = await claude([{ role: "user", content: [imgBlock(dataUrl), { type: "text", text:
    `You help a two-person household (${names}). Look at this photo and work out what they're doing. Reply ONLY JSON: {"kind":"job"|"note"|"info","seen":"one short warm sentence about what you recognised","title":"short title (if shopping, use Shopping)","steps":["2-4 short steps, only for a job, else []"],"items":["lines or items read from the photo, in order, deduplicated, else []"],"shopping":true|false,"assignee":"${members[0].id}"|"${members[1].id}"|"together","why":"one warm first-person sentence — for info, say why it's worth keeping","photoProof":true|false}. kind: "job" = a thing to do/fix/clean; "note" = an actionable list (shopping, errands, to-dos); "info" = knowledge with NO action needed — meal plans, weekly menus, schedules, timetables, wifi codes, instructions, recipes, opening hours, posters. For info, put each line of the content into items (e.g. "Monday — pasta").` }] }]);
  return asJson(txt);
}
const NUDGE_LINES = [
  (t) => `Gentle poke — "${t}" is waiting`,
  (t) => `Little reminder: "${t}"`,
  (t) => `No rush — just don't forget "${t}"`,
  (t) => `"${t}" would love some attention`,
];

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
  try {
    const m = getMsal(clientId); if (!m) return null;
    const accs = m.getAllAccounts();
    if (!accs.length) return null;
    return (email && accs.find((a) => a.username.toLowerCase() === email.toLowerCase())) || accs[0];
  } catch (e) { return null; }
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
  if (!cid || !me?.outlook?.connected) return;
  getMsal(cid); try { await msalReady; } catch (e) {}
  if (!msAccount(cid, me.outlook.email)) return;
  const email = me.outlook.email;
  const access = me.outlook.access || "write";
  const writeMode = access === "write" || access === "readwrite";
  for (const t of st.tasks) {
    const cur = stateRef.current.tasks.find((x) => x.id === t.id); if (!cur) continue;
    if (cur.kind === "info") continue;
    const mine = cur.assignees?.includes(meId) || cur.assignees?.includes("together");
    const rec = cur.outlook?.[meId];
    try {
      if (writeMode && mine && cur.dueDate && !cur.daily && cur.status !== "done") {
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
        if (!writeMode) {
          await graph(cid, email, "/me/events/" + rec.eventId, { method: "DELETE" }).catch(() => {});
          const o = { ...(cur.outlook || {}) }; delete o[meId];
          saveTask({ ...cur, outlook: o });
        } else if (cur.status === "done" && !rec.doneMarked) {
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
  if (!cid || !me?.outlook?.connected) return null;
  getMsal(cid); try { await msalReady; } catch (e) {}
  if (!msAccount(cid, me.outlook.email)) return null;
  const startD = addDays(new Date(), -1), endD = addDays(new Date(), 8);
  const q = `/me/calendarView?startDateTime=${todayStr(startD)}T00:00:00&endDateTime=${todayStr(endD)}T23:59:59&$top=60&$select=id,subject,start,end`;
  const d = await graph(cid, me.outlook.email, q, { headers: { Prefer: `outlook.timezone="${TZ}"` } });
  const klausIds = new Set(st.tasks.map((t) => t.outlook?.[meId]?.eventId).filter(Boolean));
  const access = me.outlook.access || "write";
  const publish = access === "readwrite" || access === "read";
  const events = (d?.value || []).filter((e) => !klausIds.has(e.id)).map((e) => ({
    title: publish ? (e.subject || "(no title)") : "Busy",
    start: (e.start?.dateTime || "").slice(0, 16), end: (e.end?.dateTime || "").slice(0, 16),
  })).filter((e) => e.start);
  const data = { updatedAt: new Date().toISOString(), mode: access, events };
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
const KlausMark = () => (
  <img src="icon-180.png" alt="" style={{ width: 34, height: 34, borderRadius: 11, marginRight: 10, flex: "none", boxShadow: "0 2px 6px rgba(63,56,42,.12)" }} />
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
  const patchTask = (id, patch) => { const cur = stateRef.current.tasks.find((x) => x.id === id); if (cur) saveTask({ ...cur, ...patch }); };
  const outlookKick = useRef(null);
  const scheduleOutlookSync = () => {
    if (outlookKick.current) clearTimeout(outlookKick.current);
    outlookKick.current = setTimeout(() => { syncMyOutlook(stateRef, me.id, saveTask).catch(() => {}); }, 1500);
  };
  const saveTask = (task) => {
    const tasks = stateRef.current.tasks.some((t) => t.id === task.id)
      ? stateRef.current.tasks.map((t) => (t.id === task.id ? task : t))
      : [...stateRef.current.tasks, task];
    persist({ ...stateRef.current, tasks });
    sbPushTask(task).catch(() => {});
    scheduleOutlookSync();
  };
  const removeTask = (id) => {
    persist({ ...stateRef.current, tasks: stateRef.current.tasks.filter((t) => t.id !== id) });
    sbDeleteTask(id).catch(() => {});
  };
  const saveHousehold = (h) => { persist({ ...stateRef.current, household: h }); sbPushHousehold(h).catch(() => {}); };
  const sendNudge = (to, text, taskId) => {
    const n = { id: uid(), from: me.id, to, text, taskId: taskId || null, ts: new Date().toISOString() };
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
        const localKeep = remote.tasks.length ? local.tasks.filter((t) => !t.seed) : local.tasks;
        const merged = {
          household: remote.household || local.household,
          tasks: [...remote.tasks, ...localKeep.filter((t) => !remoteIds.has(t.id))],
          nudges: remote.nudges.length ? remote.nudges : local.nudges,
          outlookCache: { ...(local.outlookCache || {}), ...(remote.outlookCache || {}) },
        };
        persist(merged); setSyncStatus("connected");
        if (!remote.household) sbPushHousehold(merged.household).catch(() => {});
        localKeep.filter((t) => !remoteIds.has(t.id)).forEach((t) => sbPushTask(t).catch(() => {}));
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

  const addTyped = async (text) => {
    const p = await parseTyped(text);
    if (p.shopping) {
      const ex = stateRef.current.tasks.find((x) => x.kind !== "info" && x.status !== "done" && (x.shopping || /shopping|groceries/i.test(x.title)));
      if (ex) {
        const items = [...(ex.listItems || []), ...p.items.map((i) => ({ text: i, done: false }))];
        const on = p.on || ex.dueDate;
        const auto = ex.deadlineAuto !== false;
        saveTask({ ...ex, shopping: true, listItems: items, dueDate: on, dueTime: p.time || ex.dueTime, deadline: p.deadline || (auto ? on : ex.deadline), deadlineAuto: p.deadline ? false : auto, reminder: on ? { at: on + "T" + (p.time || "09:00") + ":00" } : ex.reminder });
        setRoute({ name: "task", id: ex.id });
        return;
      }
    }
    const deadline = p.deadline || p.on || null;
    const t = {
      id: uid(), title: p.title, steps: [], listItems: p.shopping ? p.items.map((i) => ({ text: i, done: false })) : [], shopping: p.shopping,
      assignees: [me.id], status: "todo", dueDate: p.on, dueTime: p.time, deadline, deadlineTime: p.deadlineTime, deadlineAuto: !p.deadline,
      reminder: p.on ? { at: p.on + "T" + (p.time || "09:00") + ":00" } : null,
      photoProof: false, source: "typed", createdAt: todayStr(), comments: [],
    };
    saveTask(t);
    stackRef.current.push(route);
    setRoute({ name: "task", id: t.id, edit: true, create: true });
  };

  if (!who) return <Onboard members={state.household.members} onPick={(id) => { localStorage.setItem(LS_DEVICE, JSON.stringify({ who: id })); setWho(id); }} />;

  const stackRef = useRef([]);
  const TABS = ["home", "tasks", "calendar", "info"];
  const nav = (name, params = {}) => { stackRef.current.push(route); if (stackRef.current.length > 20) stackRef.current.shift(); if (TABS.includes(name)) setTab(name); setRoute({ name, ...params }); };
  const goBack = () => { const prev = stackRef.current.pop() || { name: tab }; if (TABS.includes(prev.name)) setTab(prev.name); setRoute(prev); };
  const goTab = (t) => { stackRef.current = []; setTab(t); setRoute({ name: t }); };
  const touchRef = useRef(null);
  const onTouchStart = (e) => { const t = e.touches[0]; touchRef.current = { x: t.clientX, y: t.clientY, el: e.target }; };
  const onTouchEnd = (e) => {
    const st = touchRef.current; touchRef.current = null; if (!st) return;
    const tag = (st.el?.tagName || "").toLowerCase(); if (tag === "input" || tag === "textarea") return;
    const t = e.changedTouches[0]; const dx = t.clientX - st.x, dy = t.clientY - st.y;
    if (Math.abs(dx) < 70 || Math.abs(dy) > 60) return;
    if (TABS.includes(route.name)) {
      const idx = TABS.indexOf(route.name); const ni = dx < 0 ? idx + 1 : idx - 1;
      if (ni >= 0 && ni < TABS.length) goTab(TABS[ni]);
    } else if (dx > 0) goBack();
  };
  const screenProps = { state, me, other, saveTask, removeTask, saveHousehold, sendNudge, completeTask, uncompleteTask, addTyped, patchTask, nav, goBack, goTab, tab, route, syncStatus, setSyncStatus };

  let screen;
  switch (route.name) {
    case "home": screen = <Home {...screenProps} />; break;
    case "tasks": screen = <Tasks {...screenProps} />; break;
    case "calendar": screen = <CalendarView {...screenProps} />; break;
    case "info": screen = <InfoList {...screenProps} />; break;
    case "infoItem": screen = <InfoDetail {...screenProps} infoId={route.id} />; break;
    case "task": screen = <TaskDetail {...screenProps} taskId={route.id} />; break;
    case "addPhoto": screen = <AddFromPhoto {...screenProps} photo={route.photo} />; break;
    case "finish": screen = <FinishPhoto {...screenProps} taskId={route.id} photo={route.photo} />; break;
    case "sharing": screen = <Sharing {...screenProps} />; break;
    case "settings": screen = <Settings {...screenProps} />; break;
    default: screen = <Home {...screenProps} />;
  }
  return <div onTouchStart={onTouchStart} onTouchEnd={onTouchEnd} style={{ maxWidth: 430, margin: "0 auto", minHeight: "100vh", background: C.bg, fontFamily: FONT, display: "flex", flexDirection: "column" }}>{screen}</div>;
}

const isTodayTask = (t) => t.kind !== "info" && (t.daily || t.dueDate === todayStr() || (t.dueDate && t.dueDate < todayStr() && t.status !== "done") || (t.deadline && t.deadline <= todayStr() && t.status !== "done") || (t.status === "done" && t.doneAt && t.doneAt.slice(0, 10) === todayStr()) || (!t.dueDate && t.status !== "done"));

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
        { value: "home", label: "Today" }, { value: "tasks", label: "Tasks" }, { value: "calendar", label: "Calendar" }, { value: "info", label: "Info" },
      ]} />
    </div>
  );
}
function BottomBar({ onText, onPhoto }) {
  const [txt, setTxt] = useState("");
  const [busy, setBusy] = useState(false);
  const photoRef = useRef();
  const submit = async () => {
    if (!txt.trim() || busy) return;
    setBusy(true);
    try { await onText(txt.trim()); setTxt(""); } finally { setBusy(false); }
  };
  const pick = async (e) => { const f = e.target.files?.[0]; e.target.value = ""; if (!f) return; onPhoto(await downscale(f)); };
  return (
    <div style={{ position: "sticky", bottom: 0, marginTop: "auto" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px 18px", background: "#fff", boxShadow: "0 -4px 14px rgba(63,56,42,.06)" }}>
        <input value={txt} onChange={(e) => setTxt(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
          placeholder="Tell me what needs doing…"
          style={{ flex: 1, background: C.bg, border: "none", outline: "none", borderRadius: 999, padding: "13px 17px", fontSize: 13.5, fontWeight: 600, color: C.ink, fontFamily: FONT, minWidth: 0 }} />
        {(txt.trim() || busy) && (busy
          ? <span className="tspin" style={{ width: 40, height: 40, borderRadius: "50%", border: "3px solid " + C.olive, borderTopColor: "transparent", flex: "none", boxSizing: "border-box" }} />
          : <span onClick={submit} style={{ width: 40, height: 40, borderRadius: "50%", background: C.oliveSoft, display: "grid", placeItems: "center", flex: "none", cursor: "pointer", color: C.olive, fontWeight: 800, fontSize: 17 }}>✓</span>)}
        <span onClick={() => photoRef.current.click()} style={{ width: 46, height: 46, borderRadius: "50%", background: C.olive, display: "grid", placeItems: "center", flex: "none", boxShadow: shadowBtn, cursor: "pointer" }}><CameraSvg /></span>
      </div>
      <input ref={photoRef} type="file" accept="image/*" capture="environment" hidden onChange={pick} />
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
function NudgeCard({ nudge, from, onDismiss, onOpen }) {
  const [dx, setDx] = useState(0);
  const start = useRef(null);
  const moved = useRef(false);
  return (
    <div
      onTouchStart={(e) => { e.stopPropagation(); start.current = e.touches[0].clientX; moved.current = false; }}
      onTouchMove={(e) => { if (start.current != null) { const d = e.touches[0].clientX - start.current; if (Math.abs(d) > 8) moved.current = true; setDx(d); } }}
      onTouchEnd={(e) => { e.stopPropagation(); const d = dx; start.current = null; if (Math.abs(d) > 80) { setDx(d > 0 ? 420 : -420); setTimeout(onDismiss, 160); } else setDx(0); }}
      onClick={() => { if (!moved.current) onOpen(); }}
      style={{ display: "flex", gap: 11, alignItems: "center", background: C.terraSoft, borderRadius: 20, padding: "14px 16px", marginTop: 12, cursor: "pointer", transform: `translateX(${dx}px)`, opacity: Math.max(0, 1 - Math.abs(dx) / 320), transition: start.current == null ? "transform .18s, opacity .18s" : "none" }}>
      <Avatar m={from} size={34} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.ink }}>A little reminder from {from.name}</div>
        <div style={{ fontSize: 12.5, color: C.terraDark, marginTop: 2, lineHeight: 1.4 }}>"{nudge.text}"</div>
        <div style={{ fontSize: 10.5, color: C.mut, marginTop: 3, fontWeight: 600 }}>swipe to mark as read</div>
      </div>
      <span style={{ color: C.terra, fontSize: 16, fontWeight: 700 }}>›</span>
    </div>
  );
}
function Home({ state, me, other, nav, goTab, tab, addTyped, completeTask, uncompleteTask }) {
  const [showDone, setShowDone] = useState(false);
  const [readIds, setReadIds] = useState(() => { try { return JSON.parse(localStorage.getItem("klaus.readNudges")) || []; } catch (e) { return []; } });
  const markRead = (id) => { const r = [...readIds, id].slice(-100); setReadIds(r); localStorage.setItem("klaus.readNudges", JSON.stringify(r)); };
  const todays = state.tasks.filter(isTodayTask);
  const done = todays.filter((t) => t.status === "done").length;
  const total = todays.length; const left = total - done;
  const frac = total ? done / total : 0;
  const dash = 351.9;
  const [big, small] = heroCopy(done, total, left, new Date().getHours());
  const nudge = [...state.nudges].reverse().find((n) => n.to === me.id && !readIds.includes(n.id) && Date.now() - new Date(n.ts) < 48 * 3600000);
  const nudgeFrom = nudge && state.household.members.find((m) => m.id === nudge.from);
  const open = todays.filter((t) => t.status !== "done");
  const doneList = todays.filter((t) => t.status === "done");
  const label = (t) => {
    if (t.status === "done") { const by = state.household.members.find((m) => m.id === t.doneBy); return (by ? by.name : "Someone") + " did it" + (t.doneWithPhoto ? " — with a photo" : ""); }
    const w = t.assignees?.includes("together") || t.assignees?.length > 1 ? "Together" : t.assignees?.[0] === me.id ? "Yours" : (other.name + "'s");
    const bits = [w, t.daily ? "Daily" + ((t.streak || 0) > 0 ? " · " + t.streak + " days running" : "") : fmtDue(t)];
    const dl = deadlineLabel(t);
    if (dl) bits.push(dl);
    if (t.photoProof && !t.daily) bits.push("photo when done");
    return bits.join(" · ");
  };
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><KlausMark /><H1>Our day</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 0" }}>
      <div style={{ ...card(24, { boxShadow: shadowHero }), padding: 18, display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ position: "relative", width: 96, height: 96, flex: "none" }}>
          <svg width="96" height="96" viewBox="0 0 132 132"><circle cx="66" cy="66" r="56" fill="none" stroke={C.track} strokeWidth="12" /><circle cx="66" cy="66" r="56" fill="none" stroke={C.olive} strokeWidth="12" strokeLinecap="round" strokeDasharray={dash} strokeDashoffset={dash * (1 - frac)} transform="rotate(-90 66 66)" style={{ transition: "stroke-dashoffset .6s" }} /></svg>
          <div style={{ position: "absolute", inset: 0, display: "grid", placeItems: "center" }}><div style={{ fontSize: 25, fontWeight: 800, color: C.ink }}>{done}<span style={{ fontSize: 13, color: C.mut, fontWeight: 600 }}>/{total}</span></div></div>
        </div>
        <div style={{ flex: 1 }}><div style={{ fontSize: 16, fontWeight: 700, color: C.ink, lineHeight: 1.3 }}>{big}</div><div style={{ fontSize: 13, color: C.mut, marginTop: 4, lineHeight: 1.45 }}>{small}</div></div>
      </div>
      {nudge && <NudgeCard nudge={nudge} from={nudgeFrom} onDismiss={() => markRead(nudge.id)} onOpen={() => (nudge.taskId && state.tasks.some((t) => t.id === nudge.taskId)) ? nav("task", { id: nudge.taskId }) : markRead(nudge.id)} />}
      <div style={{ fontSize: 13, fontWeight: 700, color: C.mut, margin: "20px 2px 8px" }}>{new Date().getHours() < 12 ? "This morning" : new Date().getHours() < 18 ? "This afternoon" : "This evening"}</div>
      {open.map((t) => (
        <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ ...card(20), padding: "15px 16px", display: "flex", alignItems: "center", gap: 13, marginBottom: 9, cursor: "pointer" }}>
          <CheckCircle state={t.status === "active" ? "active" : "todo"} onClick={(e) => { e.stopPropagation(); t.photoProof ? nav("task", { id: t.id }) : completeTask(t); }} />
          <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 700, color: C.ink }}>{t.title}</div><div style={{ fontSize: 11.5, color: C.mut, marginTop: 1 }}>{label(t)}</div></div>
          {t.daily && (t.streak || 0) > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: C.olive, background: C.oliveSoft, borderRadius: 999, padding: "4px 10px", whiteSpace: "nowrap" }}>{t.streak} days</span>}
          {t.photoBefore && <span style={{ width: 36, height: 36, borderRadius: 11, flex: "none", overflow: "hidden" }}><img src={t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></span>}
        </div>
      ))}
      {doneList.length > 0 && (
        <div onClick={() => setShowDone(!showDone)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 2px 10px", cursor: "pointer" }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: C.mut }}>Done today · {doneList.length}</span>
          <span style={{ fontSize: 11, color: C.faint, transform: showDone ? "rotate(180deg)" : "none", transition: "transform .15s" }}>▾</span>
        </div>
      )}
      {showDone && doneList.map((t) => (
        <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ ...card(20), padding: "15px 16px", display: "flex", alignItems: "center", gap: 13, marginBottom: 9, opacity: 0.65, cursor: "pointer" }}>
          <CheckCircle state="done" onClick={(e) => { e.stopPropagation(); uncompleteTask(t); }} />
          <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 700, color: C.strike, textDecoration: "line-through" }}>{t.title}</div><div style={{ fontSize: 11.5, color: C.faint, marginTop: 1 }}>{label(t)}</div></div>
          {(t.photoAfter || t.photoBefore) && <span style={{ width: 36, height: 36, borderRadius: 11, flex: "none", overflow: "hidden" }}><img src={t.photoAfter || t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></span>}
        </div>
      ))}
      {total === 0 && <div style={{ fontSize: 13, color: C.mut, fontWeight: 600, textAlign: "center", padding: "18px 0" }}>Type below, or snap a photo of the job.</div>}
      <div style={{ height: 12 }} />
    </div>
    <BottomBar onText={addTyped} onPhoto={(p) => nav("addPhoto", { photo: p })} />
  </>);
}

/* ---------- 8b Tasks ---------- */
function Tasks({ state, me, other, nav, goTab, tab, addTyped }) {
  const [mode, setMode] = useState("person");
  const [showDone, setShowDone] = useState(false);
  const open = state.tasks.filter((t) => t.status !== "done" && t.kind !== "info");
  const doneRecent = state.tasks.filter((t) => t.kind !== "info" && t.status === "done" && t.doneAt && Date.now() - new Date(t.doneAt) < 3 * 86400000);
  const forPerson = (id) => open.filter((t) => t.assignees?.includes(id) || t.assignees?.includes("together"));
  const st = state.household.streak || { count: 0 };
  const Mini = ({ t }) => {
    const doneT = t.status === "done"; const pill = !doneT && relDuePill(t);
    return (
      <div onClick={() => nav("task", { id: t.id })} style={{ ...card(18), marginBottom: 9, overflow: "hidden", opacity: doneT ? (t.doneWithPhoto ? 0.75 : 0.6) : 1, cursor: "pointer" }}>
        {t.photoBefore && <div style={{ height: 52, position: "relative" }}><img src={t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></div>}
        <div style={{ padding: t.photoBefore ? "10px 13px" : "12px 13px" }}>
          <div style={{ fontSize: 13.5, fontWeight: 700, lineHeight: 1.25, color: doneT ? C.strike : C.ink, textDecoration: doneT ? "line-through" : "none" }}>{t.title}</div>
          <div style={{ fontSize: 11, fontWeight: 600, color: doneT ? C.olive : C.mut, marginTop: 3 }}>
            {doneT ? (t.doneWithPhoto ? "Done with photo" : "Done " + new Date(t.doneAt).toTimeString().slice(0, 5)) : fmtDue(t) + (deadlineLabel(t) ? " · " + deadlineLabel(t) : "")}
          </div>
          {!doneT && (t.photoProof || (pill && pill.bg === C.terraSoft)) && (
            <div style={{ display: "flex", gap: 5, marginTop: 8, flexWrap: "wrap" }}>
              {pill && pill.bg === C.terraSoft && <span style={{ fontSize: 10, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "3px 9px", whiteSpace: "nowrap" }}>{pill.text}</span>}
              {t.photoProof && <span style={{ fontSize: 10, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "3px 9px", whiteSpace: "nowrap" }}>photo when done</span>}
            </div>
          )}
        </div>
      </div>
    );
  };
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><KlausMark /><H1>Tasks</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 16px" }}>
      <SegPill style={{ marginBottom: 14 }} value={mode} onChange={setMode} options={[{ value: "together", label: "Together" }, { value: "person", label: "By person" }]} />
      {mode === "person" ? (
        <div style={{ display: "flex", gap: 10 }}>
          {[me, other].map((p) => (
            <div key={p.id} style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 8 }}><Avatar m={p} size={24} /><span style={{ fontSize: 13, fontWeight: 800, color: C.ink }}>{p.name}</span><span style={{ fontSize: 11, fontWeight: 600, color: C.mut }}>{forPerson(p.id).length} open</span></div>
              {forPerson(p.id).map((t) => <Mini key={t.id} t={t} />)}
            </div>
          ))}
        </div>
      ) : (
        <div>{open.map((t) => <Mini key={t.id} t={t} />)}</div>
      )}
      {doneRecent.length > 0 && (
        <div onClick={() => setShowDone(!showDone)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 2px 10px", cursor: "pointer" }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: C.mut }}>Done lately · {doneRecent.length}</span>
          <span style={{ fontSize: 11, color: C.faint, transform: showDone ? "rotate(180deg)" : "none", transition: "transform .15s" }}>▾</span>
        </div>
      )}
      {showDone && doneRecent.map((t) => <Mini key={t.id} t={t} />)}
      {st.count > 1 && (
        <div style={{ display: "flex", gap: 11, alignItems: "center", background: C.oliveSoft, borderRadius: 18, padding: "13px 16px", marginTop: 8 }}>
          <span style={{ width: 30, height: 30, borderRadius: "50%", background: C.olive, color: "#fff", display: "grid", placeItems: "center", fontSize: 13, flex: "none" }}>✓</span>
          <div style={{ fontSize: 13, color: C.oliveDark, lineHeight: 1.4, fontWeight: 600 }}><b style={{ color: C.ink }}>The house is on a {st.count}-day streak</b> — everything done, together.</div>
        </div>
      )}
    </div>
    <BottomBar onText={addTyped} onPhoto={(p) => nav("addPhoto", { photo: p })} />
  </>);
}

/* ---------- 8c Task detail ---------- */
function TaskDetail({ state, me, other, nav, goBack, route, taskId, saveTask, removeTask, completeTask, uncompleteTask, sendNudge }) {
  const t = state.tasks.find((x) => x.id === taskId);
  const finRef = useRef();
  const [comment, setComment] = useState("");
  const [editing, setEditing] = useState(!!route.edit);
  const [editTitle, setEditTitle] = useState(!!route.edit && !!route.create);
  const [newStep, setNewStep] = useState("");
  const [nudged, setNudged] = useState("");
  if (!t) return <div style={{ padding: 24, fontSize: 14, color: C.mut, fontWeight: 600 }}>That task is gone. <span onClick={() => nav("home")} style={{ color: C.olive, cursor: "pointer" }}>Back home.</span></div>;
  const assignee = t.assignees?.includes("together") || (t.assignees?.length || 0) > 1 ? null : state.household.members.find((m) => m.id === t.assignees?.[0]);
  const toggleStep = (i) => { const steps = t.steps.map((s, j) => (j === i ? { ...s, done: !s.done } : s)); saveTask({ ...t, steps, status: t.status === "todo" && steps.some((s) => s.done) ? "active" : t.status }); };
  const remTime = t.reminder?.at ? t.reminder.at.slice(11, 16) : "09:00";
  const setDue = (dueDate) => {
    const auto = t.deadlineAuto !== false;
    saveTask({ ...t, daily: false, dueDate, deadline: auto ? dueDate : t.deadline, deadlineAuto: auto, reminder: dueDate ? { at: dueDate + "T" + remTime + ":00" } : null });
  };
  const pickFinish = async (e) => { const f = e.target.files?.[0]; e.target.value = ""; if (!f) return; nav("finish", { id: t.id, photo: await downscale(f) }); };
  const doneT = t.status === "done";
  const dlText = deadlineLabel(t);
  const nudgeTargets = t.assignees?.includes("together") ? [other.id] : (t.assignees || []).filter((id) => id !== me.id);
  const doNudge = () => {
    if (!nudgeTargets.length) return;
    const msg = NUDGE_LINES[Math.floor(Math.random() * NUDGE_LINES.length)](t.title);
    nudgeTargets.forEach((id) => sendNudge(id, msg, t.id));
    const name = state.household.members.find((m) => m.id === nudgeTargets[0])?.name || "them";
    setNudged("Nudged " + name + " — a little heart is waiting on their home screen.");
    setTimeout(() => setNudged(""), 2600);
  };
  const dateChip = (label, val) => (
    <span onClick={() => setDue(val)} style={{ fontSize: 12, fontWeight: 700, color: t.dueDate === val && !t.daily ? "#fff" : C.ink2, background: t.dueDate === val && !t.daily ? C.ink : C.bg, borderRadius: 999, padding: "8px 13px", cursor: "pointer", whiteSpace: "nowrap" }}>{label}</span>
  );
  const editRow = (label, control) => (
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0", borderTop: "1px solid " + C.divider }}>
      <div style={{ fontSize: 12.5, fontWeight: 700, color: C.mut, width: 74, flex: "none" }}>{label}</div>
      <div style={{ flex: 1, display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>{control}</div>
    </div>
  );
  const timeInput = (val, fn) => (
    <input type="time" value={val || ""} onChange={(e) => fn(e.target.value)}
      style={{ background: C.bg, border: "none", outline: "none", borderRadius: 10, padding: "7px 10px", fontSize: 12.5, fontWeight: 700, color: C.ink, fontFamily: FONT }} />
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}>
      <Back onClick={goBack} />
      {editTitle ? (
        <input autoFocus defaultValue={t.title} onBlur={(e) => { saveTask({ ...t, title: e.target.value.trim() || t.title }); setEditTitle(false); }}
          onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
          style={{ flex: 1, fontSize: 18, fontWeight: 800, letterSpacing: "-.01em", color: C.ink, border: "none", outline: "none", background: "#fff", borderRadius: 12, padding: "6px 10px", fontFamily: FONT, boxShadow: shadow, minWidth: 0 }} />
      ) : (
        <span onClick={() => setEditTitle(true)} style={{ flex: 1, cursor: "pointer", minWidth: 0 }}><H1 small>{t.title}</H1></span>
      )}
    </div>
    <div style={{ padding: "0 20px 18px" }}>
      {t.photoBefore && <Photo src={t.photoBefore} h={168} label={"before" + (t.createdAt ? " · " + fmtDue({ dueDate: t.createdAt }).toLowerCase() : "")} style={{ marginBottom: 13 }} />}
      <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
        <span style={{ fontSize: 11.5, fontWeight: 700, color: doneT ? C.oliveDark : C.terraDark, background: doneT ? C.oliveSoft : C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>{doneT ? "done" : t.status === "active" ? "in progress" : "to do"}</span>
        <span onClick={() => setEditing(!editing)} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: C.ink, background: "#fff", borderRadius: 999, padding: "5px 12px", boxShadow: "0 1px 4px rgba(63,56,42,.07)", whiteSpace: "nowrap", cursor: "pointer" }}>
          {assignee ? <><Avatar m={assignee} size={16} />{assignee.id === me.id ? "Yours" : assignee.name}</> : <>{state.household.members.map((m) => <Avatar key={m.id} m={m} size={16} ring />)}Together</>}
        </span>
        <span onClick={() => setEditing(!editing)} style={{ fontSize: 11.5, fontWeight: 700, color: "#54695a", background: "#fff", borderRadius: 999, padding: "6px 12px", boxShadow: "0 1px 4px rgba(63,56,42,.07)", whiteSpace: "nowrap", cursor: "pointer" }}>{fmtDue(t)}</span>
        {dlText && <span onClick={() => setEditing(!editing)} style={{ fontSize: 11.5, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>{dlText}</span>}
        {t.photoProof && !doneT && <span onClick={() => setEditing(!editing)} style={{ fontSize: 11.5, fontWeight: 700, color: C.olive, background: C.oliveSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap", cursor: "pointer" }}>photo when done</span>}
      </div>
      {editing && (
        <div style={{ ...card(22), padding: "6px 18px 12px", marginTop: 13 }}>
          {editRow("Who", state.household.members.map((m) => (
            <span key={m.id} onClick={() => saveTask({ ...t, assignees: [m.id] })} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 700, color: t.assignees?.[0] === m.id && t.assignees.length === 1 ? "#fff" : C.ink2, background: t.assignees?.[0] === m.id && t.assignees.length === 1 ? C.ink : C.bg, borderRadius: 999, padding: "7px 12px", cursor: "pointer" }}><Avatar m={m} size={16} />{m.name}</span>
          )).concat(
            <span key="tog" onClick={() => saveTask({ ...t, assignees: ["together"] })} style={{ fontSize: 12, fontWeight: 700, color: t.assignees?.includes("together") ? "#fff" : C.ink2, background: t.assignees?.includes("together") ? C.ink : C.bg, borderRadius: 999, padding: "8px 13px", cursor: "pointer" }}>Together</span>
          ))}
          {editRow("On", <>
            {dateChip("Anytime", null)}{dateChip("Today", todayStr())}{dateChip("Tomorrow", todayStr(addDays(new Date(), 1)))}{dateChip("Saturday", nextSaturday())}
            <input type="date" value={t.dueDate || ""} onChange={(e) => setDue(e.target.value || null)}
              style={{ background: C.bg, border: "none", outline: "none", borderRadius: 10, padding: "7px 10px", fontSize: 12.5, fontWeight: 700, color: C.ink, fontFamily: FONT }} />
          </>)}
          {t.dueDate && !t.daily && editRow("At", <>
            {timeInput(t.dueTime, (v) => saveTask({ ...t, dueTime: v || null }))}
            <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>{t.dueTime ? "— in the calendar at " + t.dueTime : "no time — calendar entry at 9:00"}</span>
          </>)}
          {editRow("Due by", <>
            <input type="date" value={t.deadline || ""} onChange={(e) => saveTask({ ...t, deadline: e.target.value || null, deadlineAuto: e.target.value ? false : true, deadlineTime: e.target.value ? t.deadlineTime : null })}
              style={{ background: C.bg, border: "none", outline: "none", borderRadius: 10, padding: "7px 10px", fontSize: 12.5, fontWeight: 700, color: C.ink, fontFamily: FONT }} />
            {t.deadline && timeInput(t.deadlineTime, (v) => saveTask({ ...t, deadlineTime: v || null, deadlineAuto: false }))}
            {t.deadline && <span onClick={() => saveTask({ ...t, deadline: null, deadlineTime: null, deadlineAuto: true })} style={{ fontSize: 14, color: C.faint, cursor: "pointer", padding: "0 4px" }}>×</span>}
            <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>{t.deadlineAuto !== false && t.deadline ? "follows the day it's planned on" : "deadline — plan it any day before"}</span>
          </>)}
          {editRow("Remind", t.dueDate && !t.daily ? <>
            {timeInput(remTime, (v) => saveTask({ ...t, reminder: { at: t.dueDate + "T" + v + ":00" } }))}
            <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>on the day</span>
          </> : <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>set a day first</span>)}
          {editRow("Daily", <>
            <Toggle on={!!t.daily} onClick={() => saveTask({ ...t, daily: !t.daily, dueDate: null, dueTime: null, deadline: null, deadlineTime: null, deadlineAuto: true, reminder: null })} />
            <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>repeats every day, builds a streak</span>
          </>)}
          {editRow("Photo", <>
            <Toggle on={!!t.photoProof} onClick={() => saveTask({ ...t, photoProof: !t.photoProof })} />
            <span style={{ fontSize: 11.5, fontWeight: 600, color: C.mut }}>finish with an "after" photo</span>
          </>)}
        </div>
      )}
      {!doneT && (
        <div style={{ ...card(22), padding: "8px 18px", marginTop: 13 }}>
          {(t.steps || []).map((s, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: i ? "1px solid #f0f3ee" : "none" }}>
              <CheckCircle size={24} state={s.done ? "done" : i === t.steps.findIndex((x) => !x.done) ? "active" : "todo"} onClick={() => toggleStep(i)} />
              <span onClick={() => toggleStep(i)} style={{ flex: 1, fontSize: 14, fontWeight: s.done ? 600 : 700, color: s.done ? C.faint : i === t.steps.findIndex((x) => !x.done) ? C.ink : C.mut, textDecoration: s.done ? "line-through" : "none", cursor: "pointer" }}>{s.text}</span>
              <span onClick={() => saveTask({ ...t, steps: t.steps.filter((_, j) => j !== i) })} style={{ fontSize: 13, color: C.faint, cursor: "pointer", padding: "0 4px" }}>×</span>
            </div>
          ))}
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: (t.steps || []).length ? "1px solid #f0f3ee" : "none" }}>
            <span style={{ width: 24, height: 24, borderRadius: "50%", border: "2.5px dashed " + C.line, flex: "none", boxSizing: "border-box", display: "grid", placeItems: "center", color: C.faint, fontSize: 13 }}>+</span>
            <input value={newStep} onChange={(e) => setNewStep(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && newStep.trim()) { saveTask({ ...t, steps: [...(t.steps || []), { text: newStep.trim(), done: false }] }); setNewStep(""); } }}
              placeholder="Add a step…" style={{ flex: 1, border: "none", outline: "none", fontSize: 14, fontWeight: 600, color: C.ink, fontFamily: FONT, background: "transparent" }} />
          </div>
        </div>
      )}
      {doneT && t.steps?.length > 0 && (
        <div style={{ ...card(22), padding: "8px 18px", marginTop: 13 }}>
          {t.steps.map((s, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: i ? "1px solid #f0f3ee" : "none" }}>
              <CheckCircle size={24} state="done" />
              <span style={{ fontSize: 14, fontWeight: 600, color: C.faint, textDecoration: "line-through" }}>{s.text}</span>
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
      {t.reminder?.at && !doneT && !editing && (
        <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginTop: 11, boxShadow: shadow }}>
          <span style={{ width: 34, height: 34, borderRadius: "50%", background: C.oliveSoft, display: "grid", placeItems: "center", flex: "none" }}>
            <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke={C.olive} strokeWidth="1.8" strokeLinecap="round"><path d="M12 8v5l3 2M12 3a9 9 0 109 9" /></svg>
          </span>
          <div style={{ flex: 1, fontSize: 13, color: "#54695a", lineHeight: 1.45 }}><b style={{ color: C.ink }}>{fmtDue({ dueDate: t.reminder.at.slice(0, 10) })} {t.reminder.at.slice(11, 16)} reminder</b> — and {other.name} hears the good news the moment it's done.</div>
        </div>
      )}
      {!route.create && (t.comments || []).map((c, i) => {
        const by = state.household.members.find((m) => m.id === c.by) || other;
        return (
          <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", marginTop: 12, flexDirection: c.by === me.id ? "row-reverse" : "row" }}>
            <Avatar m={by} size={30} />
            <div style={{ background: C.terraSoft, borderRadius: 18, [c.by === me.id ? "borderTopRightRadius" : "borderTopLeftRadius"]: 6, padding: "11px 14px", fontSize: 13.5, lineHeight: 1.5, color: C.terraDark }}>"{c.text}"</div>
          </div>
        );
      })}
      {!route.create && <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <input value={comment} onChange={(e) => setComment(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && comment.trim()) { saveTask({ ...t, comments: [...(t.comments || []), { by: me.id, text: comment.trim(), ts: new Date().toISOString() }] }); setComment(""); } }}
          placeholder="Add a comment…" style={{ flex: 1, background: "#fff", border: "none", outline: "none", borderRadius: 999, padding: "11px 15px", fontSize: 13, fontWeight: 600, color: C.ink, fontFamily: FONT, boxShadow: shadow }} />
      </div>}
      {route.create ? (
        <>
          <PrimaryBtn style={{ marginTop: 16 }} onClick={() => nav("home")}>Save task</PrimaryBtn>
          <div onClick={() => { removeTask(t.id); nav("home"); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 12, cursor: "pointer" }}>Discard</div>
        </>
      ) : !doneT ? (
        <>
          <div onClick={() => setEditing(!editing)} style={{ marginTop: 16, background: "#fff", borderRadius: 999, padding: 15, fontSize: 14.5, fontWeight: 800, color: C.ink, textAlign: "center", boxShadow: shadow, cursor: "pointer" }}>{editing ? "Close editor" : "Edit task"}</div>
          {nudgeTargets.length > 0 && (
            <div onClick={doNudge} style={{ marginTop: 10, background: C.terraSoft, color: C.terraDark, borderRadius: 999, padding: 14, fontSize: 14, fontWeight: 800, textAlign: "center", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
              <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke={C.terra} strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round"><path d="M21 11.3a8.3 8.3 0 01-8.3 8.3c-1.2 0-2.3-.2-3.3-.7L4.2 20.6l1.6-4A8.3 8.3 0 1121 11.3z" /><path d="M12.4 14.6l-2.7-2.6c-2.5-2.4 1.1-5.2 2.7-3 1.6-2.2 5.2.6 2.7 3z" fill={C.terra} stroke="none" /></svg>
              Nudge {state.household.members.find((m) => m.id === nudgeTargets[0])?.name} about this
            </div>
          )}
          {nudged && <div style={{ fontSize: 12.5, color: C.terraDark, fontWeight: 700, marginTop: 10, textAlign: "center" }}>{nudged}</div>}
          <div style={{ display: "flex", gap: 9, marginTop: 10 }}>
            <div onClick={() => finRef.current.click()} style={{ flex: 1.6, background: C.olive, color: "#fff", borderRadius: 999, padding: 15, fontSize: 14.5, fontWeight: 700, textAlign: "center", boxShadow: shadowBtn, cursor: "pointer" }}>Finish with a photo</div>
            <div onClick={() => { completeTask(t); nav("home"); }} style={{ flex: 1, background: "#fff", borderRadius: 999, padding: 15, fontSize: 14, fontWeight: 700, color: "#748078", textAlign: "center", boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}>Mark done</div>
          </div>
        </>
      ) : (
        <div style={{ display: "flex", gap: 9, marginTop: 16 }}>
          <div onClick={() => uncompleteTask(t)} style={{ flex: 1, background: "#fff", borderRadius: 999, padding: 15, fontSize: 14, fontWeight: 700, color: "#748078", textAlign: "center", boxShadow: "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}>Bring it back</div>
        </div>
      )}
      {!route.create && <div onClick={() => { removeTask(t.id); nav("home"); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 14, cursor: "pointer" }}>Remove task</div>}
      <input ref={finRef} type="file" accept="image/*" capture="environment" hidden onChange={pickFinish} />
    </div>
  </>);
}

function nextSaturday() { const d = new Date(); d.setDate(d.getDate() + ((6 - d.getDay() + 7) % 7 || 7)); return todayStr(d); }

/* ---------- 8d Smart capture: one camera, Klaus works out the rest ---------- */
function AddFromPhoto({ state, me, other, nav, goBack, photo, saveTask, patchTask }) {
  const [draft, setDraft] = useState(null);
  const [failed, setFailed] = useState(false);
  const [title, setTitle] = useState("");
  const stepsLabels = useMemo(() => ["Having a look at your photo", "Working out what you're up to", "Drafting it…"], []);
  const shown = useSteppedReveal(stepsLabels, !draft);
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const d = await aiSmartPhoto(photo, state.household.members);
        if (alive) { setDraft(d); setTitle(d.title || ""); }
      } catch (e) {
        if (!alive) return;
        setFailed(true);
        const d = { kind: "job", seen: "Saved your photo with the task", title: "New task from your photo", steps: [], items: [], shopping: false, assignee: me.id, why: "Give it a title and I'll keep the photo attached. Add your Claude key in Settings and I'll work photos out for you next time.", photoProof: true };
        setDraft(d); setTitle(d.title);
      }
    })();
    return () => { alive = false; };
  }, []);
  const kind = draft?.kind || "job";
  const hasItems = (draft?.items?.length || 0) > 0;
  const [saveMode, setSaveMode] = useState(null);
  const effSave = saveMode ?? (kind === "info" ? "info" : "task");
  const wantTask = effSave === "task" || effSave === "both";
  const wantInfo = effSave === "info" || effSave === "both";
  const [noteMode, setNoteMode] = useState("list");
  const [assignee, setAssignee] = useState(null);
  const [due, setDueC] = useState(null);
  const [proof, setProof] = useState(null);
  const effAssignee = assignee ?? draft?.assignee ?? me.id;
  const effDue = due === null ? (kind === "job" ? "sat" : null) : due;
  const dueDate = effDue === "sat" ? nextSaturday() : effDue === "tomorrow" ? todayStr(addDays(new Date(), 1)) : effDue === "today" ? todayStr() : null;
  const effProof = proof ?? (kind === "job" ? (draft?.photoProof ?? state.household.settings.photoProofDefault) : false);
  const assigneeM = effAssignee === "together" ? null : state.household.members.find((m) => m.id === effAssignee);
  const rows = draft
    ? [draft.seen || stepsLabels[0],
       kind === "info" ? (hasItems ? "Read " + draft.items.length + " line" + (draft.items.length > 1 ? "s" : "") : "Read it through")
         : kind === "note" ? (hasItems ? "Found " + draft.items.length + " item" + (draft.items.length > 1 ? "s" : "") : "Read the note")
         : ("Drafted the task" + (draft.steps?.length ? " and " + draft.steps.length + " steps" : "")),
       kind === "info" ? "Looks like one to keep, not to do"
         : draft.assignee === "together" ? "Suggested you take it together"
         : "Suggested " + (state.household.members.find((m) => m.id === draft.assignee)?.name || me.name)]
    : stepsLabels.slice(0, Math.max(1, shown));
  const log = () => {
    let taskId = null, infoId = null;
    const uploads = [];
    if (wantTask) {
      if (kind !== "job" && noteMode === "each" && hasItems) {
        draft.items.forEach((it) => saveTask({ id: uid(), title: it, steps: [], assignees: [effAssignee === "together" ? "together" : effAssignee], status: "todo", photoProof: false, source: "note", createdAt: todayStr(), comments: [] }));
      } else {
        const t = {
          id: uid(), title: title || "New task",
          steps: kind === "job" ? (draft?.steps || []).map((x) => ({ text: x, done: false })) : [],
          listItems: kind !== "job" ? (draft?.items || []).map((x) => ({ text: x, done: false })) : [],
          shopping: !!draft?.shopping,
          assignees: [effAssignee === "together" ? "together" : effAssignee],
          dueDate, dueTime: null, deadline: dueDate, deadlineAuto: true,
          reminder: dueDate ? { at: dueDate + "T09:00:00" } : null,
          photoProof: effProof, photoBefore: photo, status: "todo", source: kind === "job" ? "photo" : "note", createdAt: todayStr(), comments: [],
        };
        saveTask(t); uploads.push(t.id); taskId = t.id;
      }
    }
    if (wantInfo) {
      const inf = { id: uid(), kind: "info", title: title || "Worth keeping", listItems: (draft?.items || []).map((x) => ({ text: x, done: false })), photoBefore: photo, status: "todo", assignees: [], createdAt: todayStr(), comments: [] };
      saveTask(inf); uploads.push(inf.id); infoId = inf.id;
    }
    uploads.forEach((id) => sbUploadPhoto(photo, "cap-" + id).then((u) => { if (u !== photo) patchTask(id, { photoBefore: u }); }));
    if (taskId) nav("task", { id: taskId, edit: true, create: true });
    else if (infoId) nav("infoItem", { id: infoId });
    else nav("tasks");
  };
  const chip = (label, on, fn) => (
    <span onClick={fn} style={{ background: on ? C.ink : "#fff", borderRadius: 999, padding: "9px 14px", fontSize: 12.5, fontWeight: 600, color: on ? "#fff" : C.ink2, boxShadow: on ? "none" : "0 1px 4px rgba(63,56,42,.07)", cursor: "pointer" }}><span style={{ color: on ? C.terraSoft : C.terra }}>↳ </span>{label}</span>
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={goBack} /><H1 small>{kind === "info" ? "Worth keeping" : kind === "note" ? "From your note" : "New task"}</H1></div>
    <div style={{ padding: "0 20px 18px" }}>
      <Photo src={photo} h={150} label="your photo · just now" />
      <div style={{ marginTop: 13 }}><AiChecklist rows={rows} running={!draft} /></div>
      {draft && (<>
        <div style={{ ...card(22), padding: "16px 18px", marginTop: 11 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
            <input value={title} onChange={(e) => setTitle(e.target.value)} style={{ flex: 1, minWidth: 0, border: "none", outline: "none", fontSize: 16, fontWeight: 800, color: C.ink, fontFamily: FONT, background: "transparent", padding: 0 }} />
            {hasItems && kind !== "job" && <span style={{ fontSize: 12, fontWeight: 700, color: C.terra, flex: "none" }}>{draft.items.length} {kind === "info" ? "lines" : "items"}</span>}
          </div>
          {kind === "job" && <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, marginTop: 3 }}>{draft.steps?.length ? draft.steps.length + " steps drafted" : "no steps — just the one job"}</div>}
          {hasItems && kind !== "job" && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 12 }}>
              {draft.items.map((it, i) => <span key={i} style={{ fontSize: 12.5, fontWeight: 600, color: C.ink2, background: C.bg, borderRadius: 999, padding: "7px 13px" }}>{it}</span>)}
            </div>
          )}
          <div style={{ display: "flex", gap: 7, marginTop: 12, flexWrap: "wrap" }}>
            {wantTask && <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: C.ink, background: C.bg, borderRadius: 999, padding: "5px 12px", whiteSpace: "nowrap" }}>
              {assigneeM ? <><Avatar m={assigneeM} size={16} />{assigneeM.name}</> : "Together"}
            </span>}
            {wantTask && <span style={{ fontSize: 11.5, fontWeight: 700, color: C.ink2, background: C.bg, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>{dueDate ? fmtDue({ dueDate }) + " 9:00" : "Anytime"}</span>}
            {wantTask && <span style={{ fontSize: 11.5, fontWeight: 700, color: C.oliveDark, background: C.oliveSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>saved as a task</span>}
            {wantInfo && <span style={{ fontSize: 11.5, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>kept as info</span>}
            {wantTask && effProof && <span style={{ fontSize: 11.5, fontWeight: 700, color: C.terraDark, background: C.terraSoft, borderRadius: 999, padding: "6px 12px", whiteSpace: "nowrap" }}>photo when done</span>}
          </div>
        </div>
        <div style={{ display: "flex", gap: 11, alignItems: "flex-start", background: C.oliveSoft, borderRadius: 20, padding: "14px 16px", marginTop: 11 }}>
          <div><div style={{ fontSize: 10.5, letterSpacing: ".1em", textTransform: "uppercase", color: C.olive, fontWeight: 800 }}>My read</div>
            <div style={{ fontSize: 13, lineHeight: 1.5, color: C.oliveDark, marginTop: 4, fontWeight: 600 }}>{draft.why}</div></div>
        </div>
        {kind !== "job" && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginTop: 13 }}>
            {chip("Make it a task", effSave === "task", () => setSaveMode("task"))}
            {chip("Keep as info", effSave === "info", () => setSaveMode("info"))}
            {chip("Both", effSave === "both", () => setSaveMode("both"))}
          </div>
        )}
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginTop: kind !== "job" ? 8 : 13 }}>
          {wantTask && hasItems && kind !== "job" && chip("One list", noteMode === "list", () => setNoteMode("list"))}
          {wantTask && hasItems && kind !== "job" && chip("Each item its own task", noteMode === "each", () => setNoteMode("each"))}
          {wantTask && chip("Give it to " + (effAssignee === other.id ? me.name : other.name), false, () => setAssignee(effAssignee === other.id ? me.id : other.id))}
          {wantTask && chip("Do it together", effAssignee === "together", () => setAssignee("together"))}
          {wantTask && kind === "job" && chip(effDue === "sat" ? "Make it tomorrow" : "Make it Saturday", false, () => setDueC(effDue === "sat" ? "tomorrow" : "sat"))}
          {wantTask && kind === "job" && chip(effProof ? "No photo needed" : "Ask for a photo", false, () => setProof(!effProof))}
        </div>
        <PrimaryBtn style={{ marginTop: 15 }} onClick={log}>{effSave === "both" ? "Save both" : wantInfo && !wantTask ? "Keep it" : kind === "job" ? "Log it" : "Add to the list"}</PrimaryBtn>
        {failed && <div style={{ fontSize: 11.5, color: C.mut, fontWeight: 600, textAlign: "center", marginTop: 10, lineHeight: 1.5 }}>Smart drafting needs a Claude key — add one under the gear.</div>}
      </>)}
    </div>
  </>);
}

/* ---------- 8f Finish with photo ---------- *//* ---------- 8f Finish with photo ---------- *//* ---------- 8f Finish with photo ---------- */
function FinishPhoto({ state, me, other, nav, taskId, photo, saveTask, completeTask, sendNudge }) {
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
        <PrimaryBtn style={{ marginTop: 15 }} onClick={() => { sendNudge(other.id, "Done ✓ " + t.title + " — with a photo to prove it"); nav("home"); }}>Send it to {other.name}</PrimaryBtn>
        <div onClick={() => { const cur = state.tasks.find((x) => x.id === t.id); saveTask({ ...cur, status: "todo", doneAt: null, doneBy: null, doneWithPhoto: false, photoAfter: null }); nav("task", { id: t.id }); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 12, cursor: "pointer" }}>Undo</div>
      </>)}
    </div>
  </>);
}

/* ---------- 8g Reminders & nudges ---------- */
function InfoList({ state, nav, goTab, tab, addTyped }) {
  const [showArch, setShowArch] = useState(false);
  const infos = state.tasks.filter((t) => t.kind === "info");
  const bySort = (arr) => [...arr].sort((x, y) => (y.createdAt || "").localeCompare(x.createdAt || ""));
  const active = bySort(infos.filter((i) => !i.archivedAt));
  const archived = bySort(infos.filter((i) => i.archivedAt));
  const Card = ({ t }) => (
    <div onClick={() => nav("infoItem", { id: t.id })} style={{ ...card(18), marginBottom: 10, overflow: "hidden", cursor: "pointer", opacity: t.archivedAt ? 0.6 : 1 }}>
      {t.photoBefore && <div style={{ height: 88, position: "relative" }}><img src={t.photoBefore} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></div>}
      <div style={{ padding: "12px 14px" }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.ink, lineHeight: 1.25 }}>{t.title}</div>
        <div style={{ fontSize: 11, fontWeight: 600, color: C.mut, marginTop: 3 }}>
          {(t.listItems?.length ? t.listItems.length + " line" + (t.listItems.length > 1 ? "s" : "") + " · " : "")}kept {shortDay(t.createdAt || todayStr())}
        </div>
      </div>
    </div>
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}><KlausMark /><H1>Info</H1><Gear onClick={() => nav("settings")} /></div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 16px" }}>
      {active.length === 0 && archived.length === 0 && (
        <div style={{ fontSize: 13, color: C.mut, fontWeight: 600, lineHeight: 1.6, padding: "16px 4px", textAlign: "center" }}>
          Snap anything worth keeping — the weekly meal plan, a wifi code, opening hours, the boiler instructions — and Klaus files it here for both of you.
        </div>
      )}
      {active.map((t) => <Card key={t.id} t={t} />)}
      {archived.length > 0 && (
        <div onClick={() => setShowArch(!showArch)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 2px 10px", cursor: "pointer" }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: C.mut }}>Archived · {archived.length}</span>
          <span style={{ fontSize: 11, color: C.faint, transform: showArch ? "rotate(180deg)" : "none", transition: "transform .15s" }}>▾</span>
        </div>
      )}
      {showArch && archived.map((t) => <Card key={t.id} t={t} />)}
    </div>
    <BottomBar onText={addTyped} onPhoto={(p) => nav("addPhoto", { photo: p })} />
  </>);
}

function InfoDetail({ state, me, other, nav, goBack, infoId, saveTask, removeTask }) {
  const t = state.tasks.find((x) => x.id === infoId);
  const [editTitle, setEditTitle] = useState(false);
  const [newLine, setNewLine] = useState("");
  const [comment, setComment] = useState("");
  if (!t) return <div style={{ padding: 24, fontSize: 14, color: C.mut, fontWeight: 600 }}>That's gone. <span onClick={() => nav("info")} style={{ color: C.olive, cursor: "pointer" }}>Back to Info.</span></div>;
  return (<>
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}>
      <Back onClick={goBack} />
      {editTitle ? (
        <input autoFocus defaultValue={t.title} onBlur={(e) => { saveTask({ ...t, title: e.target.value.trim() || t.title }); setEditTitle(false); }}
          onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
          style={{ flex: 1, fontSize: 18, fontWeight: 800, letterSpacing: "-.01em", color: C.ink, border: "none", outline: "none", background: "#fff", borderRadius: 12, padding: "6px 10px", fontFamily: FONT, boxShadow: shadow, minWidth: 0 }} />
      ) : (
        <span onClick={() => setEditTitle(true)} style={{ flex: 1, cursor: "pointer", minWidth: 0 }}><H1 small>{t.title}</H1></span>
      )}
    </div>
    <div style={{ padding: "0 20px 18px" }}>
      {t.photoBefore && <Photo src={t.photoBefore} h={200} label={"kept · " + shortDay(t.createdAt || todayStr())} style={{ marginBottom: 13 }} />}
      {t.archivedAt && <div style={{ fontSize: 11.5, fontWeight: 700, color: C.mut, background: C.track, borderRadius: 999, padding: "6px 12px", display: "inline-block", marginBottom: 11 }}>archived</div>}
      <div style={{ ...card(22), padding: "8px 18px" }}>
        {(t.listItems || []).map((it, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: i ? "1px solid " + C.divider : "none" }}>
            <span style={{ width: 7, height: 7, borderRadius: "50%", background: C.olive, flex: "none" }} />
            <span style={{ flex: 1, fontSize: 14, fontWeight: 600, color: C.ink, lineHeight: 1.4 }}>{it.text}</span>
            <span onClick={() => saveTask({ ...t, listItems: t.listItems.filter((_, j) => j !== i) })} style={{ fontSize: 13, color: C.faint, cursor: "pointer", padding: "0 4px" }}>×</span>
          </div>
        ))}
        <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 0", borderTop: (t.listItems || []).length ? "1px solid " + C.divider : "none" }}>
          <span style={{ width: 7, height: 7, borderRadius: "50%", border: "1.5px dashed " + C.line, flex: "none", boxSizing: "border-box" }} />
          <input value={newLine} onChange={(e) => setNewLine(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && newLine.trim()) { saveTask({ ...t, listItems: [...(t.listItems || []), { text: newLine.trim(), done: false }] }); setNewLine(""); } }}
            placeholder="Add a line…" style={{ flex: 1, border: "none", outline: "none", fontSize: 14, fontWeight: 600, color: C.ink, fontFamily: FONT, background: "transparent" }} />
        </div>
      </div>
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
      <PrimaryBtn style={{ marginTop: 16, background: t.archivedAt ? C.olive : "#fff", color: t.archivedAt ? "#fff" : C.ink, boxShadow: t.archivedAt ? shadowBtn : shadow }}
        onClick={() => { saveTask({ ...t, archivedAt: t.archivedAt ? null : new Date().toISOString() }); nav("info"); }}>
        {t.archivedAt ? "Bring it back" : "Archive"}
      </PrimaryBtn>
      <div onClick={() => { removeTask(t.id); nav("info"); }} style={{ textAlign: "center", fontSize: 13, fontWeight: 700, color: C.mut, marginTop: 14, cursor: "pointer" }}>Remove</div>
    </div>
  </>);
}

/* ---------- 8h Calendar ---------- */
function CalendarView({ state, me, other, nav, goTab, tab }) {
  const [sel, setSel] = useState(todayStr());
  const [anchor, setAnchor] = useState(() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; });
  const first = new Date(anchor.y, anchor.m, 1);
  const monthLabel = first.toLocaleString("en-GB", { month: "long", year: "numeric" });
  const startOffset = (first.getDay() + 6) % 7; // Monday-start
  const gridStart = addDays(first, -startOffset);
  const daysInMonth = new Date(anchor.y, anchor.m + 1, 0).getDate();
  const nWeeks = Math.ceil((startOffset + daysInMonth) / 7);
  const weeks = Array.from({ length: nWeeks }, (_, w) => Array.from({ length: 7 }, (_, i) => addDays(gridStart, w * 7 + i)));
  const moveMonth = (d) => {
    const nm = new Date(anchor.y, anchor.m + d, 1);
    setAnchor({ y: nm.getFullYear(), m: nm.getMonth() });
    const today = new Date();
    setSel(nm.getFullYear() === today.getFullYear() && nm.getMonth() === today.getMonth() ? todayStr() : todayStr(nm));
  };
  const tasksOn = (ds) => state.tasks.filter((t) => t.kind !== "info" && (t.dueDate === ds || t.deadline === ds || (t.daily && ds >= todayStr())));
  const caches = state.outlookCache || {};
  const eventsOn = (ds) => state.household.members.flatMap((m) => {
    const c = caches[m.id];
    const acc = m.outlook?.access || "write";
    if (!c || acc === "write" || !m.outlook?.connected) return [];
    return (c.events || []).filter((e) => e.start.slice(0, 10) === ds).map((e) => ({ ...e, member: m }));
  });
  const connected = state.household.members.filter((m) => m.outlook?.connected);
  const lastSync = connected.map((m) => caches[m.id]?.updatedAt).filter(Boolean).sort().pop();
  const mins = lastSync ? Math.max(0, Math.round((Date.now() - new Date(lastSync)) / 60000)) : null;
  const selDate = new Date(sel + "T00:00:00");
  const selTasks = tasksOn(sel);
  const selEvents = eventsOn(sel).sort((x, y) => x.start.localeCompare(y.start));
  const navBtn = (dir) => (
    <span onClick={() => moveMonth(dir)} style={{ width: 30, height: 30, borderRadius: "50%", background: C.bg, display: "grid", placeItems: "center", fontSize: 14, color: "#8b8672", cursor: "pointer", flex: "none" }}>{dir < 0 ? "‹" : "›"}</span>
  );
  return (<>
    <div style={{ display: "flex", alignItems: "center", padding: "20px 20px 8px" }}>
      <KlausMark />
      <H1>Calendar</H1>
      <span onClick={() => nav("sharing")} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11, fontWeight: 700, color: connected.length ? C.oliveDark : C.mut, background: connected.length ? C.oliveSoft : C.track, borderRadius: 999, padding: "5px 11px", whiteSpace: "nowrap", cursor: "pointer" }}>
        <i style={{ width: 6, height: 6, borderRadius: "50%", background: connected.length ? C.olive : C.faint, display: "block" }} />
        Outlook · {connected.length ? (mins === null ? "connected" : mins < 1 ? "just now" : mins + "m ago") : "not yet"}
      </span>
    </div>
    <TabStrip tab={tab} goTab={goTab} />
    <div style={{ padding: "6px 20px 18px" }}>
      <div style={{ background: "#fff", borderRadius: 20, padding: "12px 10px 8px", boxShadow: shadow }}>
        <div style={{ display: "flex", alignItems: "center", padding: "0 6px 8px" }}>
          {navBtn(-1)}
          <div style={{ flex: 1, textAlign: "center", fontSize: 14, fontWeight: 800, color: C.ink }}>{monthLabel}</div>
          {navBtn(1)}
        </div>
        <div style={{ display: "flex" }}>
          {["M", "T", "W", "T", "F", "S", "S"].map((d, i) => <div key={i} style={{ flex: 1, textAlign: "center", fontSize: 9, fontWeight: 800, color: C.mut, paddingBottom: 4 }}>{d}</div>)}
        </div>
        {weeks.map((week, w) => (
          <div key={w} style={{ display: "flex" }}>
            {week.map((d) => {
              const ds = todayStr(d); const selD = ds === sel;
              const inMonth = d.getMonth() === anchor.m;
              const isToday = ds === todayStr();
              const tDots = tasksOn(ds).length, eDots = eventsOn(ds).length;
              const dots = [...Array(Math.min(tDots, 2)).fill("task"), ...Array(Math.min(eDots, 3 - Math.min(tDots, 2))).fill("ev")];
              return (
                <div key={ds} onClick={() => setSel(ds)} style={{ flex: 1, textAlign: "center", padding: "5px 0 4px", borderRadius: 12, background: selD ? C.ink : "transparent", cursor: "pointer" }}>
                  <div style={{ fontSize: 13.5, fontWeight: selD || isToday ? 800 : 600, color: selD ? C.bg : !inMonth ? C.faint : isToday ? C.terra : C.ink }}>{d.getDate()}</div>
                  <div style={{ display: "flex", gap: 2, justifyContent: "center", marginTop: 2, height: 4 }}>
                    {dots.map((k, i) => <i key={i} style={{ width: 4, height: 4, borderRadius: "50%", background: k === "task" ? (selD ? "#a3ad7a" : C.olive) : C.faint, display: "block" }} />)}
                  </div>
                </div>
              );
            })}
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 14, fontSize: 10.5, fontWeight: 700, color: C.mut, margin: "12px 2px 4px" }}>
        <span><i style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: C.olive }} /> task</span>
        <span><i style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: C.faint }} /> from Outlook</span>
      </div>
      <Kicker style={{ margin: "12px 2px 8px" }}>{DAYNAME[selDate.getDay()]} {selDate.getDate()} {selDate.toLocaleString("en-GB", { month: "short" })}</Kicker>
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
        const dlOnly = t.deadline === sel && t.dueDate !== sel;
        const sub = [
          t.daily ? "daily" : dlOnly ? "due by this day" : (t.dueTime || "any time"),
          together ? "together" : null,
          t.listItems?.length ? "list attached" : null,
          wroteMine ? "written to your Outlook (write)" : wroteOther ? "written to " + other.name + "'s Outlook (write)" : "in Klaus",
        ].filter(Boolean).join(" · ");
        return (
          <div key={t.id} onClick={() => nav("task", { id: t.id })} style={{ display: "flex", gap: 12, background: dlOnly ? C.terraSoft : C.oliveSoft, borderRadius: 20, padding: "13px 15px", marginBottom: 9, cursor: "pointer", opacity: t.status === "done" ? 0.7 : 1 }}>
            <div style={{ width: 3, borderRadius: 2, background: dlOnly ? C.terra : C.olive, flex: "none" }} />
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink, textDecoration: t.status === "done" ? "line-through" : "none" }}>{t.title}</div>
              <div style={{ fontSize: 11.5, fontWeight: 600, color: dlOnly ? C.terraDark : C.oliveDark, marginTop: 2 }}>{sub}</div>
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
function Sharing({ state, me, nav, goBack, saveHousehold }) {
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
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={goBack} /><H1 small>Sharing & access</H1></div>
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
            options={[{ value: "read", label: "Read" }, { value: "write", label: "Write" }, { value: "readwrite", label: "Read & write" }]} />
          <div style={{ fontSize: 11.5, fontWeight: 600, color: C.mut, lineHeight: 1.5, marginTop: 10 }}>
            {(m.outlook?.access || "write") === "readwrite"
              ? <>Klaus writes tasks into {m.id === me.id ? "your" : m.name + "'s"} Outlook <b style={{ color: C.ink2 }}>and</b> {m.id === me.id ? "your" : "their"} events show in the shared calendar.</>
              : (m.outlook?.access || "write") === "read"
              ? <>{m.id === me.id ? "Your" : m.name + "'s"} Outlook events show in the shared calendar, but Klaus writes nothing into {m.id === me.id ? "your" : "their"} calendar.</>
              : <>{m.id === me.id ? "Your" : m.name + "'s"} tasks land in {m.id === me.id ? "your" : "their"} Outlook, but other events stay private — only free/busy is shared.</>}
          </div>
        </div>
      ))}
      <InviteCard />
      <Kicker>Household</Kicker>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginBottom: 9, boxShadow: shadow }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Photo proof by default</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>Ask for an "after" photo on chores & projects</div></div>
        <Toggle on={h.settings.photoProofDefault} onClick={() => setSetting("photoProofDefault", !h.settings.photoProofDefault)} />
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", marginBottom: 9, boxShadow: shadow }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Daily digest</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>One summary push at 18:00 instead of many</div></div>
        <Toggle on={h.settings.dailyDigest} onClick={() => setSetting("dailyDigest", !h.settings.dailyDigest)} />
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", borderRadius: 20, padding: "14px 16px", boxShadow: shadow }}>
        <div style={{ flex: 1 }}><div style={{ fontSize: 14.5, fontWeight: 700, color: C.ink }}>Quiet hours</div><div style={{ fontSize: 12, fontWeight: 600, color: C.mut, marginTop: 2 }}>No pushes {h.settings.quietHours.from} – {h.settings.quietHours.to} · both of you</div></div>
        <Toggle on={h.settings.quietHours.on} onClick={() => setSetting("quietHours", { ...h.settings.quietHours, on: !h.settings.quietHours.on })} />
      </div>
    </div>
  </>);
}

function InviteCard() {
  const [msg, setMsg] = useState("");
  const k = loadKeys();
  const ready = k.supabaseUrl && k.supabaseAnon;
  const makeLink = () => location.origin + location.pathname + "#join=" + encodeURIComponent(btoa(k.supabaseUrl + "|" + k.supabaseAnon));
  const copy = async () => {
    if (!ready) { setMsg("Connect this phone to Supabase first — then the link can carry the connection."); return; }
    const link = makeLink();
    try { await navigator.clipboard.writeText(link); setMsg("Copied — send it however you like. Opening it connects their phone in one tap."); }
    catch (e) { setMsg(link); }
  };
  return (
    <div style={{ marginBottom: 16 }}>
      <div onClick={copy} style={{ display: "flex", alignItems: "center", gap: 12, background: "#fff", border: "1.5px dashed " + C.line, borderRadius: 20, padding: "14px 16px", cursor: "pointer" }}>
        <span style={{ fontSize: 17, color: C.terra, fontWeight: 700 }}>+</span>
        <div style={{ flex: 1, fontSize: 13.5, fontWeight: 700, color: C.ink2 }}>Invite someone with a link</div>
        <span style={{ fontSize: 10.5, fontWeight: 600, color: C.mut, fontFamily: "'JetBrains Mono',monospace" }}>{ready ? "tap to copy" : "connect first"}</span>
      </div>
      {msg && <div style={{ fontSize: 11.5, fontWeight: 600, color: C.oliveDark, lineHeight: 1.5, margin: "8px 4px 0", wordBreak: "break-all" }}>{msg}</div>}
    </div>
  );
}

/* ---------- Settings (gear) ---------- */
function Settings({ state, me, nav, goBack, syncStatus, setSyncStatus }) {
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
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "20px 20px 10px" }}><Back onClick={goBack} /><H1 small>Settings</H1></div>
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

/* ---------- join link: #join=<base64 url|anonkey> connects a new phone with one tap ---------- */
(function () {
  try {
    const m = location.hash.match(/#join=([^&]+)/);
    if (m) {
      const [u, k] = atob(decodeURIComponent(m[1])).split("|");
      if (u && k) { saveKeys({ ...loadKeys(), supabaseUrl: u.trim(), supabaseAnon: k.trim() }); history.replaceState(null, "", location.pathname); }
    }
  } catch (e) {}
})();

/* ---------- mount ---------- */
const style = document.createElement("style");
style.textContent = `@keyframes tspin{to{transform:rotate(360deg)}} .tspin{animation:tspin .9s linear infinite} body{margin:0;background:${C.bg}} input::placeholder{color:${C.mut};font-weight:600}`;
document.head.appendChild(style);
createRoot(document.getElementById("root")).render(<App />);
