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
