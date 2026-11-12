# Excel Daily (Rust / Windows)

Windows hotkey utility for modifying **the Excel workbook already open in Microsoft Excel**.
It does **not** edit an XLSX file directly or open/save a separate copy.

## Install

1. On GitHub, open **Actions → Build Windows EXE → latest successful run → Artifacts**.
2. Extract \`excel-daily-windows.zip\` to a writable folder, e.g. \`C:\ExcelDaily\`.
3. Set \`settings.json\`. Microsoft Excel 2013+ (desktop, Windows) is required.
4. Start \`excel-daily.exe\` once. It installs a user-level Run registry entry only when \`app.start_with_windows=true\`.
5. Keep Excel open, then use **Ctrl+1** (day) or **Ctrl+2** (night). Check \`excel-daily.log\` for diagnostics.
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
