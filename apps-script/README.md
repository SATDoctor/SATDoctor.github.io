# SAT Doctor — Booking API (Apps Script v1)

Group-class calendar backed by Google Sheets. Students pick tutor-preset sessions on the website; confirmations include the Zoom link from the slot row (or Script Properties defaults).

## Setup (one-time)

### 1. Google Sheet + script files

1. Create a spreadsheet named **SAT Doctor Bookings** (or reuse your existing enrollment sheet).
2. Open **Extensions → Apps Script**.
3. Replace `Code.gs` with the file from this folder.
4. *(Optional)* Add **HTML** file **`Calendar`** if you want a standalone scheduling page at `?page=calendar`. The public website embeds the calendar directly and does **not** use an iframe.
5. Save the project.

### 2. Initialize tabs and spreadsheet ID

In the spreadsheet, refresh the page. Use menu **SAT Doctor → Initialize sheets**, then **SAT Doctor → Save spreadsheet ID for web app** (required so the calendar works when embedded on your website).

You should see three tabs:

| Tab | Purpose |
|-----|---------|
| **Slots** | Tutor-created group sessions (availability) |
| **Bookings** | Student reservations |
| **Enrollments** | Legacy contact form submissions |

### 3. Script properties (optional but recommended)

**Project Settings → Script properties** add:

| Property | Example |
|----------|---------|
| `ZOOM_LEVEL1` | `https://zoom.us/j/...` for Level 1 group |
| `ZOOM_LEVEL2` | `https://zoom.us/j/...` for Level 2 group |
| `NOTIFY_EMAIL` | `thesatdoctor1600@gmail.com` (booking alerts) |

Per-slot Zoom URLs in the **Slots** sheet override these defaults.

### 4. Add sessions

**Option A — Menu:** **SAT Doctor → Add sample group slots** (includes June 14, 2026 and other demos).

**Option A2 — One session:** **SAT Doctor → Add June 14, 2026 session (10 AM – 12 PM ET)** adds only:

| Field | Value |
|-------|--------|
| Date | Sunday, June 14, 2026 |
| Time | 10:00 AM – 12:00 PM Eastern |
| Level | Level 1 — Beginner |
| Tutor | Krutant Mehta |
| Duration | 2 hours |

**Option B — Manual rows** on **Slots**:

| Column | Example |
|--------|---------|
| slot_id | `SLOT-A1B2C3D4` (unique) |
| start_at | `2026-06-07T14:00:00.000Z` (ISO 8601) |
| end_at | `2026-06-07T15:30:00.000Z` |
| lesson_type | `Level 1` or `Level 2` |
| title | `SAT Group — Beginner (Level 1)` |
| tutor | `Krutant Mehta` |
| capacity | `6` |
| booked_count | `0` |
| status | `open` (use `full` or `cancelled` to hide) |
| zoom_join_url | Full Zoom join URL (optional if property set) |

Times are stored in UTC in the sheet; emails display **America/New_York**.

### 5. Deploy web app

1. **Deploy → New deployment → Web app**
2. Execute as: **Me** (your account)
3. **Who has access: Anyone** — must be exactly this, **not** “Anyone with a Google account” or “Only myself”
4. Copy the `/exec` URL into `index.html` as `APPS_SCRIPT_URL`.
5. After any code change: **Deploy → Manage deployments → Edit → Version: New version → Deploy**.

### Fix “You need access” / Google Drive sign-in in the calendar box

That happens when the web app is **not** public. Parents do not sign in with Google.

1. Open Apps Script → **Deploy → Manage deployments**
2. Click the pencil (Edit) on your Web app deployment
3. Set **Who has access** → **Anyone**
4. Click **Deploy** (new version)
5. Hard-refresh the SAT Doctor website (Ctrl+F5)

Test anonymously in an incognito window:

| URL | Expected |
|-----|----------|
| `.../exec` | `{"ok":true,"message":"SAT Doctor booking API is running.",...}` |
| `.../exec?action=slots&from=2026-01-01&to=2027-12-31` | `{"ok":true,"slots":[...]}` (empty array if no sessions yet) |

Append `&callback=test` to either URL for JSONP: `test({"ok":true,...})`.

**Note:** Opening only `/exec` used to show `"Unknown action"` — that still meant the app was live; you just needed `?action=slots` or `?action=ping`.

### 6. Authorize

First run: open the deployed URL with `?action=ping` and complete Google authorization (Gmail send for confirmations).

## API (used by the website)

### List slots (JSONP)

```
GET .../exec?action=slots&from=2026-05-01&to=2026-08-01&level=Level%201&callback=cb
```

`level` is optional (`Level 1` | `Level 2`).

### Book a slot (JSONP)

```
GET .../exec?action=book&slotId=SLOT-...&name=...&email=...&phone=...&course=...&requestId=uuid&callback=cb
```

### Enrollment (POST JSON, no-cors — unchanged)

```json
{ "name": "...", "email": "...", "phone": "...", "course": "...", "message": "...", "timestamp": "..." }
```

## Managing classes day-to-day

- **Add a new group time:** new row on **Slots** with `status=open`, `booked_count=0`.
- **Cancel a session:** set `status` to `cancelled`.
- **See who booked:** **Bookings** tab filtered by `slot_id`.
- **Full session:** `booked_count` reaches `capacity` → status auto-set to `full`.

## Calendar on the website (important)

GitHub Pages often **cannot load** the Apps Script JSONP API (ad blockers, browser security). The reliable fix:

### 1. Share your booking spreadsheet

1. Open your **SAT Doctor Bookings** Google Sheet.
2. Click **Share** → **General access** → **Anyone with the link** → **Viewer**.
3. Copy the Sheet ID from the URL:  
   `https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`

### 2. Paste the ID into the website

In `index.html`, set:

```javascript
const BOOKING_SPREADSHEET_ID = 'paste-your-sheet-id-here';
```

Commit and push to GitHub Pages.

The calendar will load sessions from the public **Slots** tab (no Apps Script needed for viewing). **Booking** still uses Apps Script in the background.

### 3. Test the public sheet feed

Open in a browser (replace `SHEET_ID`):

`https://docs.google.com/spreadsheets/d/SHEET_ID/gviz/tq?tqx=out:csv&sheet=Slots`

You should see CSV text with a `slot_id` header row — not a sign-in page.

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| **Cannot reach booking API** | Set `BOOKING_SPREADSHEET_ID` in `index.html` (above); share sheet as Anyone with link can view |
| **“You need access” / Google Drive** | Web app deployment must be **Who has access: Anyone** (see above) |
| JSONP / script errors | Allow `script.google.com`; use `BOOKING_SPREADSHEET_ID` so the calendar does not depend on JSONP |
| Calendar blank, no error | Add rows to **Slots** with future dates and `status=open` |
| Calendar empty | Add rows to **Slots**; check `status=open` and future `start_at` |
| “Session full” | Increase `capacity` or add a new slot |
| No Zoom in email | Set `zoom_join_url` on the slot or `ZOOM_LEVEL1` / `ZOOM_LEVEL2` properties |
| Ad blocker | Allow `script.google.com` |
