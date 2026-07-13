/* ==========================================================================
   Outlook Response Tracker
   No AI, no backend, no billing. Runs entirely in the browser.
   Auth: MSAL.js (Authorization Code + PKCE, public client, SPA flow)
   Data: Microsoft Graph REST API, delegated Mail.Read scope only
   ========================================================================== */

/* --------------------------------------------------------------------
   1. CONFIGURATION — fill these in after registering your Azure AD app
      (see README.md for step-by-step instructions)
   -------------------------------------------------------------------- */
const CONFIG = {
  clientId: "62a927b7-6a3d-4881-9a6b-3a532c0e48f9", // Application (client) ID from Azure Portal
  authority: "https://login.microsoftonline.com/333c4125-7505-4b3b-b094-e77c2a4971f6", // single-tenant ("My organization only")
  redirectUri: "https://omlinconsultancy-gif.github.io/outlook-tracker/", // must exactly match the Redirect URI registered in Azure
  scopes: ["Mail.Read", "User.Read"],
};

/* --------------------------------------------------------------------
   2. MSAL SETUP
   -------------------------------------------------------------------- */
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: CONFIG.authority,
    redirectUri: CONFIG.redirectUri,
  },
  cache: {
    cacheLocation: "sessionStorage", // cleared when the tab closes — nothing lingers on disk
    storeAuthStateInCookie: false,
  },
};

let msalInstance;
let activeAccount = null;

async function initAuth() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Complete any redirect flow in progress
  const redirectResult = await msalInstance.handleRedirectPromise().catch((e) => {
    showSigninError(e.message);
    return null;
  });

  const accounts = msalInstance.getAllAccounts();
  if (redirectResult?.account) {
    activeAccount = redirectResult.account;
  } else if (accounts.length > 0) {
    activeAccount = accounts[0];
  }

  if (activeAccount) {
    msalInstance.setActiveAccount(activeAccount);
    startApp();
  }
}

async function signIn() {
  try {
    const result = await msalInstance.loginPopup({ scopes: CONFIG.scopes });
    activeAccount = result.account;
    msalInstance.setActiveAccount(activeAccount);
    startApp();
  } catch (e) {
    showSigninError(humanizeAuthError(e));
  }
}

function signOut() {
  msalInstance.logoutPopup({ account: activeAccount }).finally(() => window.location.reload());
}

async function getAccessToken() {
  const request = { scopes: CONFIG.scopes, account: activeAccount };
  try {
    // MSAL transparently uses the cached refresh token to get a fresh access
    // token here — this is the "automatic token refresh" step.
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    // Silent refresh failed (e.g. token expired past refresh window) — fall
    // back to an interactive prompt.
    const result = await msalInstance.acquireTokenPopup(request);
    return result.accessToken;
  }
}

function humanizeAuthError(e) {
  if (CONFIG.clientId === "YOUR_CLIENT_ID_HERE") {
    return "This app hasn't been configured yet — set CONFIG.clientId in app.js to your Azure AD app's client ID. See README.md.";
  }
  return e.errorMessage || e.message || "Sign-in failed. Please try again.";
}

function showSigninError(msg) {
  const el = document.getElementById("signin-error");
  if (el && msg) el.textContent = msg;
}

/* --------------------------------------------------------------------
   3. MICROSOFT GRAPH — fetching + pagination
   -------------------------------------------------------------------- */
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphGet(url) {
  const token = await getAccessToken();
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph request failed (${res.status}): ${body.slice(0, 200)}`);
  }
  return res.json();
}

// Fetches every page of a Graph messages collection, following @odata.nextLink,
// up to a safety cap so a huge mailbox doesn't run forever.
async function graphGetAllPages(initialUrl, cap = 400) {
  let url = initialUrl;
  let items = [];
  while (url && items.length < cap) {
    const page = await graphGet(url);
    items = items.concat(page.value || []);
    url = page["@odata.nextLink"] || null;
  }
  return items.slice(0, cap);
}

const MESSAGE_SELECT = [
  "id", "subject", "from", "toRecipients", "receivedDateTime", "sentDateTime",
  "conversationId", "isRead", "hasAttachments", "flag", "importance",
  "bodyPreview", "meetingMessageType", "isDraft", "internetMessageId",
].join(",");

async function fetchInboxMessages(sinceIso) {
  const url = `${GRAPH_BASE}/me/mailFolders/inbox/messages` +
    `?$select=${MESSAGE_SELECT}` +
    `&$filter=receivedDateTime ge ${sinceIso}` +
    `&$orderby=receivedDateTime desc` +
    `&$top=50`;
  return graphGetAllPages(url);
}

async function fetchSentMessages(sinceIso) {
  const url = `${GRAPH_BASE}/me/mailFolders/sentitems/messages` +
    `?$select=${MESSAGE_SELECT}` +
    `&$filter=sentDateTime ge ${sinceIso}` +
    `&$orderby=sentDateTime desc` +
    `&$top=50`;
  return graphGetAllPages(url);
}

/* --------------------------------------------------------------------
   4. REPLY-DETECTION ALGORITHM (rule-based, no AI)

   Primary signal: Microsoft's own conversationId groups every message in
   a thread — Outlook uses the same field to build its Conversation View.
   A received message counts as replied when a Sent Items message shares
   its conversationId AND was sent after it was received.

   That reply is further classified as a genuine reply vs. a forward by
   checking the sent message's subject prefix ("FW:"/"Fwd:" vs "RE:").

   Auto-replies are detected via subject patterns Outlook itself uses for
   Automatic Replies / Out of Office, since those never need a response.
   -------------------------------------------------------------------- */

const AUTO_REPLY_PATTERN = /^(automatic reply|auto[- ]?reply|out of office|away from (my )?(desk|office))/i;
const FORWARD_PATTERN = /^(fw|fwd)\s*:/i;
const REPLY_PATTERN = /^re\s*:/i;

function classifyThreadStatus(inboxMsg, sentByConversation) {
  const sentInThread = (sentByConversation[inboxMsg.conversationId] || [])
    .filter((s) => new Date(s.sentDateTime) > new Date(inboxMsg.receivedDateTime))
    .sort((a, b) => new Date(a.sentDateTime) - new Date(b.sentDateTime));

  if (sentInThread.length > 0) {
    const firstReply = sentInThread[0];
    if (FORWARD_PATTERN.test(firstReply.subject || "") && !REPLY_PATTERN.test(firstReply.subject || "")) {
      return "forwarded";
    }
    return "responded";
  }

  if (AUTO_REPLY_PATTERN.test(inboxMsg.subject || "")) return "auto-reply";

  return "awaiting";
}

/* --------------------------------------------------------------------
   5. CATEGORY + PRIORITY HEURISTICS (rule-based, no AI)
   -------------------------------------------------------------------- */

const SYSTEM_SENDER_PATTERN = /no-?reply|do-?not-?reply|notifications?@|mailer-daemon|automated/i;
const NEWSLETTER_PATTERN = /newsletter|digest|weekly roundup|unsubscribe/i;
const URGENT_WORDS = /urgent|asap|action required|deadline|overdue|immediately|time[- ]sensitive/i;

function classifyCategory(msg, userDomain) {
  if (msg.meetingMessageType && msg.meetingMessageType !== "none") return "Calendar Invite";

  const fromAddr = (msg.from?.emailAddress?.address || "").toLowerCase();
  const subject = (msg.subject || "").toLowerCase();

  if (SYSTEM_SENDER_PATTERN.test(fromAddr)) {
    return NEWSLETTER_PATTERN.test(subject) ? "Newsletter" : "System";
  }
  if (NEWSLETTER_PATTERN.test(subject)) return "Newsletter";

  const domain = fromAddr.split("@")[1] || "";
  if (userDomain && domain === userDomain) return "Internal";
  return "External";
}

function classifyPriority(msg, daysWaiting) {
  let score = 0;
  if (msg.importance === "high") score += 3;
  if (msg.flag?.flagStatus === "flagged") score += 2;
  if (URGENT_WORDS.test(msg.subject || "") || URGENT_WORDS.test(msg.bodyPreview || "")) score += 2;
  if (msg.hasAttachments) score += 1;
  if (daysWaiting >= 7) score += 2;
  else if (daysWaiting >= 2) score += 1;

  if (score >= 5) return "High";
  if (score >= 2) return "Medium";
  return "Low";
}

function isExcludedCategory(category, settings) {
  if (category === "Newsletter" && settings.exclNewsletter) return true;
  if (category === "System" && settings.exclSystem) return true;
  if (category === "Calendar Invite" && settings.exclCalendar) return true;
  return false;
}

/* --------------------------------------------------------------------
   6. STATE
   -------------------------------------------------------------------- */
const state = {
  rows: [],
  filtered: [],
  page: 1,
  pageSize: 8,
  sortKey: "daysWaiting",
  sortDir: "desc",
  search: "",
  statusFilter: "all",
  priorityFilter: "all",
  categoryFilter: "all",
  settings: {
    overdueDays: 7,
    exclNewsletter: true,
    exclSystem: true,
    exclCalendar: true,
    scanRangeDays: 90,
    theme: "light",
  },
};

/* --------------------------------------------------------------------
   7. MAILBOX SCAN
   -------------------------------------------------------------------- */
async function scanMailbox() {
  setLoading(true, "Connecting to Outlook…", 10);

  const since = new Date();
  since.setDate(since.getDate() - state.settings.scanRangeDays);
  const sinceIso = since.toISOString();

  setLoading(true, "Fetching Inbox messages…", 30);
  const inbox = await fetchInboxMessages(sinceIso);

  setLoading(true, "Fetching Sent Items…", 60);
  const sent = await fetchSentMessages(sinceIso);

  setLoading(true, "Matching conversation threads…", 85);

  const sentByConversation = {};
  for (const s of sent) {
    (sentByConversation[s.conversationId] ||= []).push(s);
  }

  const userDomain = (activeAccount?.username || "").split("@")[1]?.toLowerCase() || "";

  const rows = inbox
    .filter((m) => !m.isDraft)
    .map((m) => {
      const status = classifyThreadStatus(m, sentByConversation);
      const category = classifyCategory(m, userDomain);
      const receivedDate = new Date(m.receivedDateTime);
      const daysWaiting = Math.max(0, Math.floor((Date.now() - receivedDate.getTime()) / 86400000));
      const priority = classifyPriority(m, daysWaiting);
      return {
        id: m.id,
        sender: m.from?.emailAddress?.name || m.from?.emailAddress?.address || "Unknown",
        email: m.from?.emailAddress?.address || "",
        subject: m.subject || "(no subject)",
        receivedDateTime: m.receivedDateTime,
        daysWaiting,
        status,
        category,
        priority,
        flagged: m.flag?.flagStatus === "flagged",
        attachments: m.hasAttachments,
        preview: m.bodyPreview || "",
        conversationId: m.conversationId,
        webLink: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(m.id)}`,
      };
    })
    .filter((r) => !isExcludedCategory(r.category, state.settings));

  state.rows = rows;
  setLoading(true, "Done", 100);
  setTimeout(() => {
    setLoading(false);
    renderAll();
  }, 250);
}

/* --------------------------------------------------------------------
   8. RENDERING
   -------------------------------------------------------------------- */
function computeAnalytics() {
  const awaiting = state.rows.filter((r) => r.status === "awaiting");
  const buckets = {
    d2: awaiting.filter((r) => r.daysWaiting >= 2).length,
    d7: awaiting.filter((r) => r.daysWaiting >= 7).length,
    d14: awaiting.filter((r) => r.daysWaiting >= 14).length,
    d30: awaiting.filter((r) => r.daysWaiting >= 30).length,
  };
  const oldest = awaiting.reduce((max, r) => (r.daysWaiting > (max?.daysWaiting || -1) ? r : max), null);
  const totalTracked = state.rows.length;
  const respondedCount = state.rows.filter((r) => r.status === "responded").length;
  const responseRate = totalTracked ? Math.round((respondedCount / totalTracked) * 100) : 0;
  return { awaiting, buckets, oldest, totalTracked, respondedCount, responseRate };
}

function renderStatCards() {
  const a = computeAnalytics();
  const grid = document.getElementById("stat-grid");
  grid.innerHTML = "";
  const cards = [
    ["Awaiting response", a.awaiting.length, "in scanned range"],
    [`Overdue ${state.settings.overdueDays}+ days`, a.awaiting.filter((r) => r.daysWaiting >= state.settings.overdueDays).length, `${a.buckets.d14} over 14 days`],
    ["Response rate", `${a.responseRate}%`, `${a.respondedCount} of ${a.totalTracked} threads`],
    ["Oldest unanswered", a.oldest ? `${a.oldest.daysWaiting}d` : "—", a.oldest ? a.oldest.sender : "all caught up"],
  ];
  for (const [label, value, sub] of cards) {
    grid.insertAdjacentHTML("beforeend", `
      <div class="stat-card">
        <div class="stat-label">${label}</div>
        <div class="stat-value">${value}</div>
        <div class="stat-sub">${sub}</div>
      </div>`);
  }

  // Analytics tab mirrors the same numbers plus a total-scanned card
  const grid2 = document.getElementById("analytics-stat-grid");
  grid2.innerHTML = "";
  const cards2 = [
    ["Total scanned", a.totalTracked, `last ${state.settings.scanRangeDays} days`],
    ["Awaiting response", a.awaiting.length, ""],
    ["Response rate", `${a.responseRate}%`, ""],
    ["30+ days overdue", a.buckets.d30, ""],
  ];
  for (const [label, value, sub] of cards2) {
    grid2.insertAdjacentHTML("beforeend", `
      <div class="stat-card">
        <div class="stat-label">${label}</div>
        <div class="stat-value">${value}</div>
        <div class="stat-sub">${sub}</div>
      </div>`);
  }

  renderBucketChart(a.buckets, a.awaiting.length);
  renderStatusBreakdown();
}

function renderBucketChart(buckets, total) {
  const el = document.getElementById("bucket-chart");
  const rows = [["2+ days", buckets.d2], ["7+ days", buckets.d7], ["14+ days", buckets.d14], ["30+ days", buckets.d30]];
  el.innerHTML = rows.map(([label, count]) => {
    const pct = total ? Math.round((count / total) * 100) : 0;
    return `
      <div class="bucket-row">
        <div class="bucket-label">${label}</div>
        <div class="bucket-track"><div class="bucket-fill" style="width:${pct}%"></div></div>
        <div class="bucket-count">${count}</div>
      </div>`;
  }).join("");
}

function renderStatusBreakdown() {
  const counts = {};
  for (const r of state.rows) counts[r.status] = (counts[r.status] || 0) + 1;
  const colors = { awaiting: "var(--red)", responded: "var(--green)", forwarded: "var(--primary)", "auto-reply": "var(--text-faint)" };
  const labels = { awaiting: "Awaiting reply", responded: "Responded", forwarded: "Forwarded", "auto-reply": "Auto-reply" };
  const el = document.getElementById("status-breakdown");
  el.innerHTML = Object.entries(counts).map(([status, count]) => `
    <div class="status-item">
      <div class="status-swatch" style="background:${colors[status] || "#ccc"}"></div>
      ${labels[status] || status}: <strong>${count}</strong>
    </div>`).join("") || `<span class="stat-sub">No data yet.</span>`;
}

function applyFiltersAndSort() {
  let rows = [...state.rows];

  if (state.search.trim()) {
    const q = state.search.toLowerCase();
    rows = rows.filter((r) => r.sender.toLowerCase().includes(q) || r.subject.toLowerCase().includes(q));
  }
  if (state.statusFilter !== "all") rows = rows.filter((r) => r.status === state.statusFilter);
  if (state.priorityFilter !== "all") rows = rows.filter((r) => r.priority === state.priorityFilter);
  if (state.categoryFilter !== "all") rows = rows.filter((r) => r.category === state.categoryFilter);

  rows.sort((a, b) => {
    let av = a[state.sortKey], bv = b[state.sortKey];
    if (state.sortKey === "priority") {
      const order = { High: 3, Medium: 2, Low: 1 };
      av = order[a.priority]; bv = order[b.priority];
    }
    if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
    if (av < bv) return state.sortDir === "asc" ? -1 : 1;
    if (av > bv) return state.sortDir === "asc" ? 1 : -1;
    return 0;
  });

  state.filtered = rows;
  if (state.page > Math.ceil(rows.length / state.pageSize)) state.page = 1;
}

function renderCategoryFilterOptions() {
  const sel = document.getElementById("filter-category");
  const current = sel.value;
  const categories = [...new Set(state.rows.map((r) => r.category))].sort();
  sel.innerHTML = `<option value="all">All categories</option>` +
    categories.map((c) => `<option value="${c}">${c}</option>`).join("");
  if (categories.includes(current)) sel.value = current;
}

function badgeForPriority(p) {
  const cls = p === "High" ? "badge-high" : p === "Medium" ? "badge-medium" : "badge-low";
  return `<span class="badge ${cls}"><span class="badge-dot"></span>${p}</span>`;
}
function badgeForStatus(s) {
  const labels = { awaiting: "Awaiting reply", responded: "Responded", forwarded: "Forwarded", "auto-reply": "Auto-reply" };
  return `<span class="badge badge-${s}">${labels[s] || s}</span>`;
}
function badgeForCategory(c) {
  return `<span class="badge badge-cat">${c}</span>`;
}
function initials(name) {
  return name.split(" ").map((p) => p[0]).join("").slice(0, 2).toUpperCase();
}
function fmtDate(iso) {
  return new Date(iso).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function renderTable() {
  applyFiltersAndSort();
  renderCategoryFilterOptions();

  const tbody = document.getElementById("email-tbody");
  const start = (state.page - 1) * state.pageSize;
  const pageRows = state.filtered.slice(start, start + state.pageSize);

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr class="empty-row"><td colspan="7">No emails match these filters.</td></tr>`;
  } else {
    tbody.innerHTML = pageRows.map((r) => {
      const daysCls = r.daysWaiting >= state.settings.overdueDays ? "days-overdue" : r.daysWaiting >= 2 ? "days-warn" : "days-ok";
      return `
        <tr data-id="${r.id}">
          <td>
            <div class="cell-sender">
              <div class="avatar">${initials(r.sender)}</div>
              <div>
                <div class="sender-name">${escapeHtml(r.sender)}</div>
                <div class="sender-company">${escapeHtml(r.email)}</div>
              </div>
            </div>
          </td>
          <td><div class="subject-main">${escapeHtml(r.subject)}</div></td>
          <td>${fmtDate(r.receivedDateTime)}</td>
          <td><span class="days-waiting ${daysCls}">${r.daysWaiting}d</span></td>
          <td>${badgeForPriority(r.priority)}</td>
          <td>${badgeForStatus(r.status)}</td>
          <td>${badgeForCategory(r.category)}</td>
        </tr>`;
    }).join("");

    tbody.querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", () => openDrawer(tr.dataset.id));
    });
  }

  const totalPages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
  document.getElementById("page-label").textContent = `Page ${state.page} of ${totalPages} · ${state.filtered.length} results`;
  document.getElementById("prev-page").disabled = state.page <= 1;
  document.getElementById("next-page").disabled = state.page >= totalPages;
}

function escapeHtml(s) {
  return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function renderAll() {
  renderStatCards();
  renderTable();
}

/* --------------------------------------------------------------------
   9. CONVERSATION DRAWER
   -------------------------------------------------------------------- */
function openDrawer(id) {
  const r = state.rows.find((row) => row.id === id);
  if (!r) return;

  document.getElementById("drawer-company").textContent = r.email;
  document.getElementById("drawer-subject").textContent = r.subject;

  const daysCls = r.daysWaiting >= state.settings.overdueDays ? "days-overdue" : "days-ok";
  document.getElementById("drawer-body").innerHTML = `
    <div class="drawer-badges">
      ${badgeForPriority(r.priority)}
      ${badgeForStatus(r.status)}
      ${badgeForCategory(r.category)}
      ${r.flagged ? '<span class="badge badge-high">Flagged</span>' : ""}
    </div>
    <div>
      <div class="drawer-field-label">Preview</div>
      <div class="drawer-field-value">${escapeHtml(r.preview) || "<em>No preview available.</em>"}</div>
    </div>
    <div>
      <div class="drawer-field-label">Received</div>
      <div class="drawer-field-value">${fmtDate(r.receivedDateTime)} · <span class="${daysCls}">${r.daysWaiting} days waiting</span></div>
    </div>
    ${r.attachments ? `<div><div class="drawer-field-label">Attachments</div><div class="drawer-field-value">This message has one or more attachments.</div></div>` : ""}
    <a class="drawer-link" href="${r.webLink}" target="_blank" rel="noopener">Open in Outlook ↗</a>
  `;

  document.getElementById("drawer").classList.remove("hidden");
  document.getElementById("drawer-backdrop").classList.remove("hidden");
}
function closeDrawer() {
  document.getElementById("drawer").classList.add("hidden");
  document.getElementById("drawer-backdrop").classList.add("hidden");
}

/* --------------------------------------------------------------------
   10. LOADING SCREEN HELPERS
   -------------------------------------------------------------------- */
function setLoading(isLoading, text, pct) {
  const screen = document.getElementById("loading-screen");
  const app = document.getElementById("app");
  const signin = document.getElementById("signin-screen");
  if (isLoading) {
    signin.classList.add("hidden");
    app.classList.add("hidden");
    screen.classList.remove("hidden");
    document.getElementById("loading-text").textContent = text || "Working…";
    document.getElementById("loading-bar").style.width = `${pct ?? 0}%`;
  } else {
    screen.classList.add("hidden");
    app.classList.remove("hidden");
  }
}

/* --------------------------------------------------------------------
   11. APP BOOTSTRAP + EVENT WIRING
   -------------------------------------------------------------------- */
function startApp() {
  document.getElementById("account-email").textContent = activeAccount?.username || "";
  scanMailbox().catch((e) => {
    setLoading(false);
    alert(`Couldn't load your mailbox: ${e.message}`);
  });
}

function wireEvents() {
  document.getElementById("signin-btn").addEventListener("click", signIn);
  document.getElementById("signout-btn").addEventListener("click", signOut);
  document.getElementById("refresh-btn").addEventListener("click", () => scanMailbox().then(renderAll));

  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".nav-item").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
      document.getElementById(`tab-${btn.dataset.tab}`).classList.remove("hidden");
    });
  });

  document.getElementById("search-input").addEventListener("input", (e) => {
    state.search = e.target.value;
    state.page = 1;
    renderTable();
  });
  document.getElementById("filter-status").addEventListener("change", (e) => { state.statusFilter = e.target.value; state.page = 1; renderTable(); });
  document.getElementById("filter-priority").addEventListener("change", (e) => { state.priorityFilter = e.target.value; state.page = 1; renderTable(); });
  document.getElementById("filter-category").addEventListener("change", (e) => { state.categoryFilter = e.target.value; state.page = 1; renderTable(); });

  document.querySelectorAll("#email-table thead th[data-sort]").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.sort;
      if (state.sortKey === key) state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      else { state.sortKey = key; state.sortDir = "desc"; }
      renderTable();
    });
  });

  document.getElementById("prev-page").addEventListener("click", () => { state.page--; renderTable(); });
  document.getElementById("next-page").addEventListener("click", () => { state.page++; renderTable(); });

  document.getElementById("drawer-close").addEventListener("click", closeDrawer);
  document.getElementById("drawer-backdrop").addEventListener("click", closeDrawer);

  // Settings
  const overdueSlider = document.getElementById("overdue-days");
  overdueSlider.addEventListener("input", (e) => {
    document.getElementById("overdue-days-val").textContent = `${e.target.value} days`;
  });

  document.getElementById("apply-settings-btn").addEventListener("click", () => {
    state.settings.overdueDays = Number(document.getElementById("overdue-days").value);
    state.settings.exclNewsletter = document.getElementById("excl-newsletter").checked;
    state.settings.exclSystem = document.getElementById("excl-system").checked;
    state.settings.exclCalendar = document.getElementById("excl-calendar").checked;
    state.settings.scanRangeDays = Number(document.getElementById("scan-range").value);
    scanMailbox().then(renderAll);
  });

  document.getElementById("theme-btn").addEventListener("click", toggleTheme);
  document.getElementById("theme-light-btn").addEventListener("click", () => setTheme("light"));
  document.getElementById("theme-dark-btn").addEventListener("click", () => setTheme("dark"));
}

function toggleTheme() {
  setTheme(state.settings.theme === "light" ? "dark" : "light");
}
function setTheme(theme) {
  state.settings.theme = theme;
  document.body.setAttribute("data-theme", theme);
  document.getElementById("theme-light-btn").classList.toggle("active", theme === "light");
  document.getElementById("theme-dark-btn").classList.toggle("active", theme === "dark");
}

wireEvents();
initAuth().catch((e) => showSigninError(e.message));
