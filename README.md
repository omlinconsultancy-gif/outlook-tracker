# Outlook Response Tracker

A free, static, client-side dashboard that scans your Outlook Inbox and Sent
Items and flags emails you haven't replied to yet. No backend, no AI API, no
billing of any kind — it's plain HTML/CSS/JS calling Microsoft Graph directly
from your browser, authenticated with Microsoft's own MSAL.js library.

## How it decides "awaiting reply" vs "responded"

For every Inbox message it looks at Microsoft's own `conversationId` (the
same field Outlook uses to build its Conversation View) and checks whether
any Sent Items message shares that ID **and** was sent after the message was
received.

- No matching sent message → **Awaiting reply**
- A matching sent message whose subject starts with `RE:` → **Responded**
- A matching sent message whose subject starts with `FW:`/`Fwd:` → **Forwarded**
- Subject matches "Automatic reply" / "Out of Office" patterns → **Auto-reply**

Category (Client/Internal/Newsletter/System/Calendar Invite) and priority
(High/Medium/Low) are also rule-based — sender domain, importance flag,
urgency keywords, attachments, and how long the email has waited. All of this
logic lives in `app.js` under `classifyThreadStatus`, `classifyCategory`, and
`classifyPriority` — tune the patterns there if it's over- or under-flagging
anything in your mailbox.

## 1. Register a free Azure AD app (one-time, no billing)

Graph API access requires an app registration, but registration itself is
free — no credit card, no paid tier.

1. Go to [portal.azure.com](https://portal.azure.com) and sign in.
2. Search for **App registrations** → **New registration**.
3. Name it anything, e.g. "Outlook Response Tracker".
4. Under **Supported account types**, choose:
   - "Accounts in any organizational directory and personal Microsoft accounts" if you want to sign in with either a work/school or personal account, or
   - "Accounts in this organizational directory only" if it's just for your work account.
5. Under **Redirect URI**, select platform **Single-page application (SPA)**
   and enter the URL you'll host this app at, e.g.:
   - `http://localhost:5500/` for local testing
   - `https://yourname.github.io/outlook-tracker/` for GitHub Pages
   - (You can add multiple redirect URIs later under Authentication.)
6. Click **Register**.
7. Copy the **Application (client) ID** shown on the Overview page.
8. Go to **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → search for and add `Mail.Read` (and
   `User.Read`, usually already present). Personal read access typically
   doesn't require admin consent — if you're on a work/school account and it
   does, ask your IT admin to grant consent.

## 2. Configure the app

Open `app.js` and set your client ID:

```js
const CONFIG = {
  clientId: "YOUR_CLIENT_ID_HERE", // <-- paste the Application (client) ID here
  authority: "https://login.microsoftonline.com/common",
  ...
};
```

If you registered the app as single-tenant, change `authority` to
`https://login.microsoftonline.com/<your-tenant-id>`.

## 3. Run it locally

This is a static site — any local web server works (it can't be opened as a
`file://` URL because MSAL's redirect flow needs a real origin).

```bash
# Option A: Node's built-in static server
npx serve .

# Option B: Python
python3 -m http.server 5500
```

Then open the URL (matching the redirect URI you registered) in your browser
and click **Sign in with Microsoft**.

## 4. Deploy for free (optional)

Any static host works, all with free tiers:

- **GitHub Pages** — push this folder to a repo, enable Pages in settings.
- **Netlify** / **Vercel** — drag-and-drop deploy, free tier.
- **Azure Static Web Apps** — free tier, integrates naturally with Azure AD.

Whichever URL you deploy to, add it as a **Redirect URI** (SPA platform) in
your app registration under **Authentication**.

## Notes and limitations

- **Read-only.** The app only requests `Mail.Read`. It never sends, replies
  to, or modifies anything in your mailbox.
- **Nothing is stored outside your browser.** Tokens live in
  `sessionStorage` and disappear when you close the tab; nothing is sent to
  any third-party server.
- **Scan range.** To keep things fast, the app fetches messages from the
  last N days (configurable in Settings, default 90). Increase it if you
  need to look further back — larger ranges take longer to fetch.
- **Junk/Clutter is excluded automatically** simply by only reading the
  Inbox and Sent Items folders.
- **Heuristics, not certainty.** The category and priority rules are
  starting points — read them in `app.js` and adjust the regular
  expressions to match your own mailbox's patterns (e.g. your company's
  newsletter senders).
