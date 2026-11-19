# Select production shift date in Japan local time.
# DateOffset=1 always means tomorrow's calendar day, independently of shift.
function ShiftDate([string]$shift,[int]$dateOffset=0,[datetime]$now=[DateTime]::Now) {
    if($dateOffset -eq 1){return $now.Date.AddDays(1).ToString('yyyy-MM-dd')}
    if($dateOffset -ne 0){throw 'DateOffset must be 0 or 1'}
    if($shift -eq 'night' -and $now.TimeOfDay -lt [TimeSpan]::FromHours(7.5)) {
        return $now.AddDays(-1).ToString('yyyy-MM-dd')
    }
    return $now.ToString('yyyy-MM-dd')
}
