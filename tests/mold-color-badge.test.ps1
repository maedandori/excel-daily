$ErrorActionPreference='Stop'
$source=Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot '../scripts/update_excel.ps1')
foreach($pattern in @('function SetScheduleBadges','color_changed','ExcelDaily managed color badge','ExcelDaily managed drying badge','Shapes.AddShape','Placement=1','HorizontalAlignment=-4131','HorizontalAlignment=-4108','SetScheduleBadges \$sheet \$machineCol \$row','SetScheduleBadges \$sheet \$machineCol \$adopt','SetScheduleBadges \$sheet \$machineCol \$emptyRow')){
 if($source -notmatch $pattern){throw "Missing badge contract: $pattern"}
}
if($source -match 'UpdateDryingTable -sheet'){throw 'Legacy drying table must no longer be written'}
Write-Output 'COLOR_BADGE_CONTRACT_OK'

# Extract the pure planner from the actual production script.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
$fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'PlanPendingScheduleOrder'},$true)
if(!$fn){throw 'Missing pending schedule planner'}
Invoke-Expression $fn.Extent.Text
$pending=@(
 @{kind='color';row=9;key='10';when=[datetime]'2026-11-25T10:00:00'},
 @{kind='color';row=11;key='11';when=[datetime]'2026-11-25T09:50:00'},
 @{kind='color';row=15;key='13';when=[datetime]'2026-11-25T10:30:00'},
 @{kind='mold';row=9;key='m1';when=[datetime]'2026-11-26T00:10:00'},
 @{kind='mold';row=11;key='m2';when=[datetime]'2026-11-25T23:50:00'}
)
$plan=@(PlanPendingScheduleOrder $pending)
if($plan.Count -ne 5){throw 'Expected five placements'}
$colors=@($plan|Where-Object{$_.item.kind -eq 'color'})
$actualColorOrder=($colors|ForEach-Object{$_.item.key}) -join ','
if($actualColorOrder -ne '11,10,13'){throw "Color order must follow updated time, actual=$actualColorOrder"}
if($colors[0].destination -ne 9 -or $colors[0].item.row -ne 11 -or $colors[1].destination -ne 11 -or $colors[1].item.row -ne 9){throw 'Machines must swap their original rows'}
if((($colors|ForEach-Object{$_.destination}) -join ',') -ne '9,11,15'){throw 'Fixed row 13 must remain untouched'}
$molds=@($plan|Where-Object{$_.item.kind -eq 'mold'})
if((($molds|ForEach-Object{$_.item.key}) -join ',') -ne 'm2,m1'){throw 'Night order must use full date and time'}
$tied=@(PlanPendingScheduleOrder @(
 @{kind='color';row=11;key='b';when=[datetime]'2026-11-25T10:00:00'},
 @{kind='color';row=9;key='a';when=[datetime]'2026-11-25T10:00:00'}
))
if((($tied|ForEach-Object{$_.item.key}) -join ',') -ne 'a,b'){throw 'Equal times must retain row order'}
$duplicates=@(PlanPendingScheduleOrder @(
 @{kind='color';row=9;key='a';when=[datetime]'2026-11-25T10:00:00'},
 @{kind='color';row=9;key='b';when=[datetime]'2026-11-25T09:50:00'}
))
if($duplicates.Count){throw 'Ambiguous row identities must not move'}
Write-Output 'PENDING_ORDER_TESTS_OK'

# Simulate machine-6 with missing/stale metadata, without installed Excel.
foreach($name in @('MachineId','TimeStillPending','FindExistingSlot','SetScheduleRowIdentity','CanAppendScheduleEvent')){
 $fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
 if(!$fn){throw "Missing $name"}
 Invoke-Expression $fn.Extent.Text
}
function CellValue($sheet,$column,$row){return [string]$sheet["$column$row"]}
function ExcelTimeText($sheet,$column,$row){return [string]$sheet["$column$row"]}
function EventTime($event){return [string]$event.time}
$cfg=[pscustomobject]@{columns=[pscustomobject]@{mold_machine='AN';mold_time='AO';color_machine='AR';color_time='AS';color_before='AU';color_after='AW'}}
$book=[pscustomobject]@{FullName='C:\test.xlsx'}
$sheet=@{Name='day';AN9='06';AO9='10:00'}
$event=[pscustomobject]@{id='new';kind='mold';machine='6';time='9:50'}
$now=[datetime]'2026-11-25T08:00:00'
$when=[datetime]'2026-11-25T09:50:00'
$prefix="$($book.FullName)|$($sheet.Name)|mold|"
$state=@{}
if((FindExistingSlot $sheet $cfg $event $state $book 9 2 12 @($event) $now $when $false) -ne 9){throw 'Machine 6 must reuse untracked existing row'}
$state[$prefix+'old']=@{row=9;scheduled_at='2026-11-25T10:00:00'}
if((FindExistingSlot $sheet $cfg $event $state $book 9 2 12 @($event) $now $when $false) -ne 9){throw 'Machine 6 must reuse stale pending ID'}
SetScheduleRowIdentity $state ($prefix+'new') $prefix 9 '2026-11-25T09:50:00'
if($state.Count -ne 1 -or $state.ContainsKey($prefix+'old') -or $state[$prefix+'new'].row -ne 9){throw 'Stale identity must be replaced'}
$state=@{}
$sheet.AO9='07:00'
if((FindExistingSlot $sheet $cfg $event $state $book 9 2 12 @($event) $now $when $false) -ne -1){throw 'Past unknown row must not cause duplicate append'}
$sheet.AO9='09:50'
if((FindExistingSlot $sheet $cfg $event $state $book 9 2 12 @($event) $now $when $false) -ne 9){throw 'Rounded hour with leading zero must match'}
$other=[pscustomobject]@{id='other';kind='mold';machine='6';time='12:00'}
$state[$prefix+'other']=@{row=9;scheduled_at='2026-11-25T12:00:00'}
if((FindExistingSlot $sheet $cfg $event $state $book 9 2 12 @($event,$other) $now $when $false) -ne 0){throw 'Distinct scheduled order remains eligible for a separate row'}
if($source -match 'if \(\$TimeOnly -and -not \$previous\)'){throw 'Time-only must recover missing row identity'}
Write-Output 'ROW_IDENTITY_RECOVERY_TESTS_OK'

# Verify badge geometry for one, two and three tags, including narrow cells.
$fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'GetScheduleBadgeLayout'},$true)
if(!$fn){throw 'Missing vertical badge layout'}
Invoke-Expression $fn.Extent.Text
foreach($n in 1..3){
 $layout=GetScheduleBadgeLayout 55 32 20 $n
 if($null -eq $layout){throw "Layout missing for $n badges"}
 if($layout.left -lt 20 -or $layout.left+$layout.width -gt 55){throw 'Machine number must stay visible'}
 if($layout.top -lt 0 -or $layout.top+($n-1)*$layout.step+$layout.height -gt 32){throw 'Badges must fit in row'}
 if($layout.step -le $layout.height){throw 'Vertical badges overlap'}
 if($layout.width -ne 33){throw 'Multiple badges must use the same full right-column width'}
}
if($null -ne (GetScheduleBadgeLayout 24 32 20 2)){throw 'Insufficient width must not overlap machine number'}
if($null -ne (GetScheduleBadgeLayout 55 12 20 3)){throw 'Insufficient height must not overlap adjacent rows'}
if($source -notmatch '\$top\+=\$layout\.step'){throw 'Badge rendering must advance vertically'}
if($source -match '\$left\+=\$width\+1'){throw 'Horizontal stacking must be removed'}
Write-Output 'VERTICAL_BADGE_LAYOUT_TESTS_OK'

# Regression from company screenshot: 23:20 (finished), 03:30, 04:50, 04:00.
foreach($name in @('VisibleScheduleDate','GetVisiblePendingSchedule')){
 $fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
 if(!$fn){throw "Missing $name"}
 Invoke-Expression $fn.Extent.Text
}
function TestScheduleRowMarked($sheet,$machineCol,$timeCol,$row){return ($row -eq 15)}
$sheet=@{Name='night';AR7='11';AS7='23:20';AR9='10';AS9='03:30';AR11='13';AS11='04:50';AR13='19';AS13='04:00';AR15='8';AS15='05:00'}
$state=@{}
$nightDate=[datetime]'2026-11-24'
$now=[datetime]'2026-11-25T01:32:00'
$visible=@(GetVisiblePendingSchedule $sheet $cfg $state $book 7 2 12 $nightDate 'night' $now @())
if($visible.Count -ne 3){throw "Expected three unfinished rows, got $($visible.Count)"}
$plan=@(PlanPendingScheduleOrder $visible)
$sequence=($plan|ForEach-Object{$sheet["AR$($_.item.row)"]}) -join ','
if($sequence -ne '10,19,13'){throw "Visible sheet order incorrect: $sequence"}
if((($plan|ForEach-Object{$_.destination}) -join ',') -ne '9,11,13'){throw 'Finished and manually marked rows must not move'}
if((VisibleScheduleDate '23:20' $nightDate 'night') -ne [datetime]'2026-11-24T23:20:00'){throw 'Night evening date wrong'}
if((VisibleScheduleDate '04:00' $nightDate 'night') -ne [datetime]'2026-11-25T04:00:00'){throw 'Night morning date wrong'}
if($source -notmatch '(?m)^MovePendingScheduleOrder \$sheet \$cfg \$state \$visiblePending'){throw 'Both refresh modes must sort visible rows'}
Write-Output 'VISIBLE_SCHEDULE_ORDER_TESTS_OK'

# Historical data is allowed only in the full data refresh, never in time-only.
$history=[pscustomobject]@{history_replay=$true;carryover=$false}
if(-not (CanAppendScheduleEvent $history $true $false)){throw 'Missing historical machine 6 must be recoverable'}
if(CanAppendScheduleEvent $history $true $true){throw 'Time-only must never append historical rows'}
if(CanAppendScheduleEvent ([pscustomobject]@{history_replay=$false}) $true $false){throw 'Unmarked legacy past event must remain protected'}
if(CanAppendScheduleEvent ([pscustomobject]@{history_replay=$true;carryover=$true}) $true $false){throw 'Do not replay cross-shift historical rows'}
if(-not (CanAppendScheduleEvent ([pscustomobject]@{}) $false $false)){throw 'Future event must remain eligible'}
