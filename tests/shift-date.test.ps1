$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/shift_date.ps1')
$early=[datetime]'2026-11-20T02:15:00'
$evening=[datetime]'2026-11-20T19:15:00'
$cases=@(
 @('day',0,$early,'2026-11-20'),
 @('night',0,$early,'2026-11-19'),
 @('night',1,$early,'2026-11-21'),
 @('day',1,$early,'2026-11-21'),
 @('night',0,$evening,'2026-11-20'),
 @('night',1,$evening,'2026-11-21')
)
foreach($case in $cases){
 $actual=ShiftDate $case[0] $case[1] $case[2]
 if($actual -cne $case[3]){throw "ShiftDate $($case[0]) $($case[1]) => $actual, expected $($case[3])"}
}
Write-Output 'PASS: tomorrow and night-shift date selection'
