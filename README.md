# Excel Daily (Rust / Windows)

Windows hotkey utility for modifying **the Excel workbook already open in Microsoft Excel**.
It does **not** edit an XLSX file directly or open/save a separate copy.

## Install

1. On GitHub, open **Actions → Build Windows EXE → latest successful run → Artifacts**.
2. Extract \`excel-daily-windows.zip\` to a writable folder, e.g. \`C:\ExcelDaily\`.
3. Set \`settings.json\`. Microsoft Excel 2013+ (desktop, Windows) is required.
4. Start \`excel-daily.exe\` once. It installs a user-level Run registry entry only when \`app.start_with_windows=true\`.
5. Keep Excel open, then use **Ctrl+1** (day) or **Ctrl+2** (night), or **Ctrl+3** (update pending times only in both shift sheets). Check \`excel-daily.log\` for diagnostics.
6. To turn off startup: set \`start_with_windows=false\` and restart the EXE once.
7. To stop the background app, end **excel-daily.exe** in Task Manager. This initial version does not provide a tray menu.

## Settings (JSON)

\`\`\`json
"app": {
  "run_in_background": true,
  "start_with_windows": true,
  "show_status_on_start": false
}
\`\`\`

- \`run_in_background=true\`: starts without a status dialog. The application stays resident and receives global hotkeys.
- \`start_with_windows=true\`: adds \`HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ExcelDaily\`; no administrator access required.
- \`start_with_windows=false\`: removes that user-level startup entry when the EXE runs.
- Edit \`sheets.day\` and \`sheets.night\` to exactly match sheet tab names, including spaces.
- \`workbook.filename_pattern\` e.g. \`成形スケジュール*.xlsx\` matches names with variable dates.
- \`workbook.selected_workbook\` can pin an exact open name. If empty, multiple matching workbooks prompt a numbered selection.
- \`rows\` and \`columns\` are configurable. By default, day starts row 9, night row 7, step 2, 12 slots each.
- \`save_automatically=false\` leaves edits visible in Excel but unsaved until Ctrl+S.

## Data input

Use either \`data_source.events_file\` with a local JSON file, or \`data_source.api_url\` pointing at an HTTP GET API returning **exactly**:

\`\`\`json
{"events":[{"id":"order-123","shift":"day","kind":"mold","machine":"19","time":"10:30","scheduled_at":"2026-11-16T10:30:00","completed":false}]}
\`\`\`

Color events use \`"kind":"color"\`, plus optional \`color_before\` and \`color_after\`. Set \`shift\` to \`day\` or \`night\`. \`time\` must be \`HH:mm\`. An \`id\` **must remain stable** across synchronizations. Date/time values must represent local Japan time; do not send a UTC timestamp without conversion. The D1/MICS endpoint is **not hardcoded** because its API contract has not been provided.

On update, existing IDs keep their row and only change their time (while the scheduled event is still in the future). Completed or past entries remain unchanged. New future IDs append to empty rows. Existing non-empty cells are never overwritten by new IDs. The mapping is stored in \`excel-daily-state.json\` alongside the EXE. For Excel files that have been changed manually, verify the mapping before reusing them.

## Limitations

- Uses Excel COM and the first registered Excel.Application instance; a workbook open in a separate Excel process may not be visible.
- This version has no native settings window or tray menu. Edit JSON and restart via Task Manager.
- The GitHub workflow builds on Windows; a successful artifact/build is not proof the Excel integration has been tested on your work PC.
- Keep the app folder writable for its state and log. Avoid putting secrets in publicly committed \`settings.json\`; configure API tokens locally.
- Ctrl+1 and Ctrl+2 are global and might override other shortcuts. If another app already owns one, the agent reports an error.


## Shift handover (status 100, 410, 440, 490)

A MICS change that began in the previous shift but remains unfinished should be
included in the next shift's event feed with the **same stable id** and the next
shift name. Example:

\`\`\`json
{
  "events": [{
    "id": "order-19-mold-0113A10C",
    "shift": "night",
    "kind": "mold",
    "machine": "19",
    "time": "18:30",
    "scheduled_at": "2026-11-16T18:30:00",
    "status_code": "100",
    "carryover": true,
    "completed": false
  }]
}
\`\`\`

The Excel agent explicitly permits a past scheduled timestamp only when
\`carryover=true\`, \`completed=false\`, and \`status_code\` is one of
\`100,410,440,490\`. It writes to the chosen shift sheet without deleting
historical entries on the previous sheet. Repeated syncs use the same event ID
to update only the time cell of the existing row. The upstream API must mark
carryover accurately, and cease sending it as active after completion.
The machine-control web shift view independently carries live unfinished work
into the current shift; this version of the Excel agent does **not** automatically
extract events from the web shift list. Configure a compatible API data feed.


## Automatic updates (v0.2.0+)

Configure `auto_update.enabled=true`, `auto_update.check_interval_hours=6` and
`auto_update.repository="maedandori/excel-daily"` in `settings.json`.
The agent checks GitHub's **latest published Release** after the first 60 seconds
and then at the configured interval. A newer semantic version is downloaded
to a temporary folder and verified against its published SHA-256 checksum.
Only when no Excel hotkey operation is in progress does the agent launch the
detached installer, exit, replace its EXE, and restart automatically.
The old EXE is backed up as `excel-daily.exe.bak`. If replacement or restart
fails, the helper restores the backup. Logs are written to
`excel-daily-updater.log`.

**Important:** This is not an automatic download of every commit. A successful
Windows build must publish a new numbered GitHub Release with both
`excel-daily.exe` and `excel-daily.exe.sha256`. To ship a future version,
increase `version` in `Cargo.toml` before pushing. The workflow deliberately
does not silently overwrite an existing published release.

Only the executable is replaced: `settings.json`, `events.json`,
`excel-daily-state.json`, and existing Excel workbooks are preserved.
The agent does not need administrator privileges if installed in a folder
writable by the current user. Windows startup can remain enabled independently.


## Fixing HTTP 401 / missing events.json (v0.2.2)

The default \`settings.json\` now points to
\`https://machine-control.maedandoris.workers.dev/api/excel-daily/events\`.
This route requires a dedicated Cloudflare Worker secret named
\`EXCEL_DAILY_TOKEN\`, intentionally **not** the MICS_SYNC_TOKEN and not the
login password/cookie.

1. Deploy the updated \`machine-control\` Worker source including
   \`worker/excel-daily-events.js\`. A GitHub commit alone does not make the
   API available to the live Worker.
2. Set \`EXCEL_DAILY_TOKEN\` as a secret in the Worker deployment's actual
   Cloudflare account/environment. For a Wrangler-managed Worker:
   \`npx wrangler secret put EXCEL_DAILY_TOKEN\`, and paste a long random value.
3. On the Windows PC only, put the **same exact value** in
   \`settings.json → data_source.api_token\`. Do not commit this token.
4. Set \`data_source.api_url\` to the Worker URL (already the default),
   and leave \`data_source.events_file\` blank.
5. Restart Excel Daily and press Ctrl+1 or Ctrl+2.

The agent appends \`?shift=day|night&date=YYYY-MM-DD\`. For a night
shift between midnight and 07:30, the selected date is yesterday.
The Worker only permits read access to MICS schedule snapshots for this route.
A 401 means a secret/token mismatch. A 503 can mean a missing Worker secret
or unavailable D1 data. No Excel cells are changed when the request fails.

Security note: Secret checksums for GitHub self-updates only verify the
integrity of the downloaded files against the release's published checksum;
they are **not** a substitute for release signing.

## Ctrl+3, silent mode and UTF-8 (v0.2.3)

Ctrl+3 checks the day and night sheets and only changes existing **time cells** for events previously managed by Excel Daily whose scheduled time has not passed and that are not marked completed. It never adds entries, changes machine or color text, or edits entries with a past planned time. Ctrl+1 and Ctrl+2 retain add/update functionality and may repair clearly garbled Japanese color text on tracked rows without altering ordinary manually edited labels. HTTP JSON is explicitly decoded as UTF-8, the embedded script is UTF-8 with BOM for Windows PowerShell 5.1, and PowerShell runs without an opened CMD window. No settings.json change is required.

## Excel Daily v0.3.0 — two Windows EXEs, Admin updates, ten-minute rounding

**First installation:** download the complete `excel-daily-windows.zip` Release and extract
both `excel-daily.exe` (silent hotkey Agent) and `excel-daily-settings.exe` (Settings GUI)
into the same writable folder. Keep your existing `settings.json`,
`excel-daily-state.json`, `excel-daily-agent-id.txt`, and Excel workbooks. If updating
from v0.2.3, use the full ZIP once so that the new Settings EXE is installed too.

- Settings EXE edits the workbook, sheet names, API URL and token, device display name,
  Windows startup, automatic updates, and permission to receive Admin update requests.
  It can save, start/restart, and stop the local Agent or request a GitHub update.
- The Agent sends authenticated heartbeats to Machine-Control using the existing
  `data_source.api_token`; no new secret is needed. Admin can view individual
  device/version/online state and request a newer published Release for that device.
- The Agent checks remote update requests about every 90 seconds, only when not
  editing Excel. The existing periodic GitHub auto-update is independent and optional.
- Releases package **both EXEs** and a SHA-256 checksum for the complete ZIP.
  The updater stages and verifies the ZIP, then replaces both EXEs. It retains
  backups beside them and never overwrites local settings, state or Excel files.
- All hotkeys round **up** the displayed time to the next 10-minute boundary:
  `5:01 -> 5:10`, `5:10 -> 5:10`, `5:59 -> 6:00`, `23:59 -> 0:00`.
  Original MICS timestamps remain unchanged for completed/past-event protection.
- In the Excel Daily feed, if a schedule transition changes **both mold and color**,
  the transition is recorded as mold only, not duplicated in the color table.
  Color-only changes remain visible. Old already-written Excel rows are preserved
  (no destructive automatic deletion).

## Excel Daily v0.3.1: fast writes and fill colors

The Agent caches Excel COM cell references for one operation, scans available
slots only once per table, and temporarily disables ScreenUpdating/EnableEvents
during writes. Both settings are restored in a finally block, including on
errors; calculation mode is not changed. It preserves merged cells, open
workbooks and the original row tracking state.

Ctrl+1/Ctrl+2 automatically fill the entire merged color-before/color-after
cells with the recognized Japanese color; dark colors use white font. Clear
(クリア) is displayed as pale blue. Unknown colors retain template formatting.
Ctrl+3 remains time-only and does not change color cells. A broken label
repaired by the existing UTF-8 fix receives matching formatting.
Optional `settings.json` flags `formatting.fast_excel_writes` and
`formatting.color_cells` default to true when absent.

## v0.3.2 — Settings startup fix, displayed versions, editable hotkeys and tomorrow schedules

- The Settings EXE now reports PowerShell startup errors in a visible dialog and
  saves details to `excel-daily-settings.log`; it does not silently exit.
  A Windows CI self-test constructs all Settings controls before a Release is published.
- Settings header displays the compiled **Settings version** and the last
  Agent version recorded in `excel-daily-version.json`, with a running indicator.
- Choose four unique global hotkeys in Settings using `Ctrl+0..9` or
  `Ctrl+Shift+0..9`. Defaults are: Ctrl+1 for day, Ctrl+2 for night,
  Ctrl+3 for pending times only, and Ctrl+4 for tomorrow's full schedule.
  Save and restart Agent for hotkey registration to change.
- The tomorrow action imports the **next calendar day's** mold and color change
  events for *both* day and night sheets, including when invoked after midnight.
  To avoid mixing days, open the intended tomorrow workbook first. The Agent
  asks for confirmation. It only fills free Excel rows, preserving all existing
  entries. It does not automatically clear or create workbooks.
- Settings includes an **Apply tomorrow** button. This sends a local command
  to the Agent under the same busy lock as hotkeys, so no concurrent Excel
  operations are started. If the Agent is not running, start it first.
- Files `settings.json`, `excel-daily-state.json`, and existing workbooks
  are preserved by the updater. An existing `settings.json` with only day and
  night shortcuts receives default refresh/tomorrow shortcuts in memory; saving
  Settings persists the new keys.

## v0.3.3: Settings closing unexpectedly

The previous Settings timer could read a left-over
`excel-daily-update-status.txt` containing `UPDATE_READY` and immediately close
the GUI without an error or diagnostic log. Settings now uses a **unique status
file for each actual button-initiated update** and ignores old update results.
Normal startup does not automatically terminate Settings. The update monitor
closes the GUI only after a fresh and verified update request is staged.

Settings now records `EXE_START`, PowerShell start/exit and Windows Forms
show/close events, including normal exits, in both the installation folder
(`excel-daily-settings.log`) and the per-user fallback
`%LOCALAPPDATA%\\ExcelDaily\\settings.log`. No API tokens are logged.
When errors occur, Settings still shows the error dialog. The new CI test
also checks that stale `UPDATE_READY` is ignored.

## v0.3.4 — Excel 2019 Ribbon, Turso-backed refresh, no VBA

The bundled Excel-DNA add-in provides a new **EXCEL-DAILY** Ribbon tab in
Microsoft Excel 2019 (32-bit and 64-bit). No VBA, macro-enabled workbook,
MICS database connection, LAN access, or Turso authentication token is needed
on the Excel PC. The normal `.xlsx` workbook stays open.

- **データ更新**: add missing mold/color transitions to **both** shift sheets,
  update the times of recognized existing events, and preserve other cells.
- **時間更新**: update only times for previously tracked, uncompleted events.
- **翌日予定**: schedule tomorrow's day/night sheets (confirm the workbook first).

The existing agent continues to fetch data from:
`https://machine-control.maedandoris.workers.dev/api/excel-daily/events`
using the per-device `data_source.api_token` (Worker secret `EXCEL_DAILY_TOKEN`).
The Worker reads the **currently selected MICS backend**, including Turso.
Never put `TURSO_AUTH_TOKEN` in Excel or commit `api_token`.

**One-time installation:** extract `excel-daily-windows.zip` to a writable
folder. Close Excel, open PowerShell in that folder and run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-ribbon.ps1
```

The installer registers either `ExcelDailyRibbon.xll` (32-bit) or
`ExcelDailyRibbon64.xll` (64-bit) for **only the current user** in Excel
16.0. If bitness auto-detection fails, supply `-Bitness x86` or
`-Bitness x64`. Restart Excel; enable the trusted Add-in in Excel's security
dialog if required by the organization's policies. To remove the Ribbon, close
Excel and run `install-ribbon.ps1 -Remove`. Its `.xll` and Agent must
remain in the same folder. The Ribbon sends a small local command to the Agent;
the Agent and PowerShell Excel COM code remain responsible for all writes and
for the existing busy lock. An accepted command is **not** a success signal:
check `excel-daily.log` for the API/Excel result. A stale or incompatible
Excel Add-in can require closing Excel before a binary update is applied.

**30-minute pre-shift rule:** for day shift the next-sheet event feed also
considers mold/color changes scheduled 07:00–07:29; for night shift
18:30–18:59. Future planned changes are included. A past change is included
only if its exact MICS order is still in an active change status. The
read-only Worker endpoint must include the corresponding
`machine-control` update. Without that deployment, the old API cannot
supply the newly requested pre-shift events.

**Safe row matching:** stored stable event IDs always take precedence. If a
mapping file has been lost, the refresh can reuse an unambiguous existing
machine/time/color row and update only its time, not re-add it; ambiguous
manual rows remain untouched and are reported in the log. Full refresh never
deletes prior rows. Completed events and manual edits remain protected.
The original Ctrl+1/Ctrl+2/Ctrl+3/Ctrl+4 hotkeys remain available.

**Verification:** GitHub Windows CI checks that both packed 32/64-bit XLLs
exist. Manual validation on a real installed Excel 2019 instance and a live
authorized Worker is still necessary before production rollout.

## v0.3.5 — automatic Excel 2019 Ribbon registration fix

In v0.3.4 the Ribbon was included in the ZIP but only appeared after the user
manually ran `install-ribbon.ps1`. On first launch of `excel-daily.exe`,
v0.3.5 automatically registers the bundled XLL for the detected Excel bitness
in the current user's Office 16.0 `OPEN*` options, with no VBA or admin rights.
Existing unrelated Excel add-ins are preserved. Auto-update to v0.3.5
copies the XLLs and installer before restarting the Agent. **Excel must be
fully closed and restarted for the Ribbon tab to become visible.** If
`excel-daily.exe` has not been started, run it once.

Use the full ZIP rather than downloading only the EXE. For failures, run
the following in PowerShell from the installation folder and inspect the
`RIBBON_REGISTER_` lines in `excel-daily.log`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-ribbon.ps1 -Diagnose
```

This diagnostic is read-only: reports whether bundled XLL files exist,
detected Excel architecture, existing Add-in registry entries, and whether
Excel is open. If the XLL is registered and still missing, use
**ファイル > オプション > アドイン** to check Excel Add-ins and Disabled Items.
An organizational security policy may block an XLL regardless of correct
registration. The program does not disable Trust Center protections.
