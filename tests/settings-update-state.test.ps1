$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/settings_update_state.ps1')
if(ShouldCloseSettingsForUpdate $false 'UPDATE_READY'){throw 'Stale UPDATE_READY closed Settings'}
if(ShouldCloseSettingsForUpdate $false 'UPDATE_FAILED'){throw 'Failed old update closed Settings'}
if(ShouldCloseSettingsForUpdate $true 'DOWNLOADING'){throw 'Pending update closed Settings too early'}
if(!(ShouldCloseSettingsForUpdate $true 'UPDATE_READY')){throw 'Fresh update must close Settings for EXE replacement'}
Write-Output 'PASS: no auto-close from stale updater state'
