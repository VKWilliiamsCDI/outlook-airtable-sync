# outlook-airtable-sync

Reads event invites and deadlines out of an Outlook inbox (via Microsoft Graph)
and syncs them into an Airtable table so they show up alongside the rest of
your planning data.

- **Event invites** come from your Outlook calendar (`/me/calendarView`) —
  i.e. meetings you've received, not a guess parsed out of email text.
- **Deadlines** are detected by scanning recent inbox messages for
  deadline-style language ("due by", "RSVP by", "submit by", etc.) and
  parsing the nearby date with `chrono-node`.
- Already-synced items are skipped on each run (tracked by Outlook
  message/event ID in the "Outlook Message ID" column), so it's safe to run
  repeatedly, e.g. on a schedule.

Writes to the **Inbox Dates** table in the **Marketing & Sales Ops Hub**
Airtable base, with columns: Title, Date, Type, Source Subject, Sender,
Email Link, Notes, Synced At, Outlook Message ID.

## Setup

### 1. Register an Azure AD app (one-time, in your Microsoft 365 admin/account)

1. Go to https://portal.azure.com → **App registrations** → **New registration**.
2. Name it anything (e.g. "Outlook Airtable Sync"). Supported account types:
   "Accounts in any organizational directory and personal Microsoft
   accounts" (or narrower, if your org requires it).
3. Redirect URI: platform **Mobile and desktop applications**,
   `http://localhost:3000/redirect`.
4. Under **API permissions**, add delegated Microsoft Graph permissions:
   `Mail.Read`, `Calendars.Read`, `offline_access`, `User.Read`. Grant admin
   consent if your tenant requires it.
5. Copy the **Application (client) ID** and your **Directory (tenant) ID**
   (or use `common` if this is a personal/multi-tenant account).

### 2. Get an Airtable API key

Create a personal access token at https://airtable.com/create/tokens with
`data.records:read` and `data.records:write` scope on the
**Marketing & Sales Ops Hub** base.

### 3. Configure

```bash
cp .env.example .env
# fill in MS_CLIENT_ID, MS_TENANT_ID, AIRTABLE_API_KEY
bun install   # or: npm install
```

### 4. Sign in to Microsoft (one-time, refreshes automatically after)

```bash
bun run login
```

This prints a device-login URL and code — open it in a browser and sign in.
Tokens are cached in `.cache/` (gitignored) and refreshed silently after
that.

### 5. Run a sync

```bash
bun run sync
```

Run this on a schedule (cron, a scheduled GitHub Action, Task Scheduler,
etc.) to keep Airtable current. `LOOKBACK_DAYS` in `.env` controls how far
back it scans messages and how far forward it looks for calendar events
(default 14).

## Notes / next steps

- This only reads mail and calendar (`Mail.Read`, `Calendars.Read`) — it
  never sends or modifies anything in Outlook.
- Deadline detection is keyword + NLP-date based, not perfect; tune
  `DEADLINE_KEYWORDS` in `src/extract.ts` as you see false positives/negatives.
- To restrict which mail folder is scanned, set `MS_MAIL_FOLDER` in `.env`
  (defaults to `inbox`).
