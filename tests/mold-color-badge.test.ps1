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
if(($colors.item.key -join ',') -ne '11,10,13'){throw 'Color order must follow updated time'}
if(($colors.destination -join ',') -ne '9,11,15'){throw 'Fixed row 13 must remain untouched'}
$molds=@($plan|Where-Object{$_.item.kind -eq 'mold'})
if(($molds.item.key -join ',') -ne 'm2,m1'){throw 'Night order must use full date and time'}
$tied=@(PlanPendingScheduleOrder @(
 @{kind='color';row=11;key='b';when=[datetime]'2026-11-25T10:00:00'},
 @{kind='color';row=9;key='a';when=[datetime]'2026-11-25T10:00:00'}
))
if(($tied.item.key -join ',') -ne 'a,b'){throw 'Equal times must retain row order'}
$duplicates=@(PlanPendingScheduleOrder @(
 @{kind='color';row=9;key='a';when=[datetime]'2026-11-25T10:00:00'},
 @{kind='color';row=9;key='b';when=[datetime]'2026-11-25T09:50:00'}
))
if($duplicates.Count){throw 'Ambiguous row identities must not move'}
Write-Output 'PENDING_ORDER_TESTS_OK'
