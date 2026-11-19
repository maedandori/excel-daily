$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/round_time.ps1')
$cases=@(@('5:01','5:10'),@('5:10','5:10'),@('5:59','6:00'),@('5:00','5:00'),@('5:51','6:00'),@('0:01','0:10'),@('23:59','0:00'),@('23:50','23:50'),@('19:09','19:10'))
foreach($case in $cases){$actual=RoundEventTimeUp $case[0];if($actual -cne $case[1]){throw "Expected $($case[0]) -> $($case[1]); got $actual"}}
foreach($bad in @('25:00','5:60','abc')){$fail=$false;try{$null=RoundEventTimeUp $bad}catch{$fail=$true};if(!$fail){throw "Expected invalid time: $bad"}}
Write-Output 'PASS: Excel Daily time rounding'
