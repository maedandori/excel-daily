# Only a newly requested update may close Excel Daily Settings.
function ShouldCloseSettingsForUpdate([bool]$active,[string]$status){
 return ($active -and $status -ceq 'UPDATE_READY')
}
