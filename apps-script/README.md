# Sid's Gym — Google Sheets Two-Way Sync (v3.2)

Starting in **v3.2**, the app no longer relies on `localStorage` to keep your data.
Google Sheets is now the **source of truth**: your data is **written** to the sheet
after every change and **read back** automatically every time the app loads. This is
why a hard refresh (or Safari/Chrome clearing site storage) no longer wipes your data.

`localStorage` is still used, but only as a fast local cache and offline fallback.

## One-time setup

You must (re)deploy the Apps Script because the endpoints changed.

1. Open your Google Sheet (or create a new blank one).
2. **Extensions → Apps Script**.
3. Delete any existing code, then paste the full contents of [`Code.gs`](./Code.gs). **Save**.
4. **Deploy → New deployment**.
   - Click the gear → **Web app**.
   - **Execute as:** Me
   - **Who has access:** Anyone
   - **Deploy**, then authorize when prompted.
5. Copy the **Web app URL** (ends in `/exec`).
6. In the app, go to the **Health** tab → **Google Sheets Sync**, paste the URL.
   The app immediately recalls anything already stored in the sheet.
7. Tap **Test Connection** to confirm it shows "Connected — script is live".

> **Updating the script later:** any time you edit `Code.gs`, you must redeploy:
> **Deploy → Manage deployments → (edit / pencil) → Version: New version → Deploy.**
> Editing the code alone does not update the live Web App.

## How it works

| Direction | When | Endpoint |
|-----------|------|----------|
| **Read (recall)** | Every app load, and right after you paste the URL | `GET ?action=getAll` → `{ status, state }` |
| **Write (save)**  | After finishing a workout, saving health metrics, body measurements, library or custom-metric changes (debounced) | `POST { action:'saveAll', state }` |
| Connection test   | "Test Connection" button | `GET ?action=ping` |

### Where the data lives in the sheet
- **`_state`** (hidden sheet, cell A1): the entire app state as one JSON document.
  This is what the app reads and writes — do not edit it by hand.
- **`Sessions`** and **`Health`** sheets: human-readable mirrors, auto-rebuilt on every
  save so you can browse your workouts and metrics as a normal spreadsheet. These are
  view-only; editing them does not change what the app reads.

### Merge behavior
On save, the script merges the incoming data with whatever is already in the cloud:
sessions and health entries are unioned by key (so data recorded offline on another
device is never dropped), and library / custom metrics / body data take the latest
values. The same union happens on the client when recalling, so nothing is lost.
