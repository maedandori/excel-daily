$ErrorActionPreference='Stop'
$source=Join-Path $PSScriptRoot '..\scripts\drying.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Resolve-Path $source).Path,[ref]$tokens,[ref]$errors)
if($errors.Count){foreach($parseError in $errors){Write-Output ('PARSER_ERROR: '+$parseError.Message+' TEXT='+$parseError.Extent.Text+' LINE='+$parseError.Extent.StartLineNumber)};throw 'Drying script parsing failed on Windows PowerShell 5.1'}
$fn=$ast.Find({param($x) $x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'DryingTimeFromConfig'},$true)
if(!$fn){throw 'Missing DryingTimeFromConfig'}
Invoke-Expression $fn.Extent.Text
$cfg=Get-Content (Join-Path $PSScriptRoot '..\settings.json') -Raw -Encoding UTF8|ConvertFrom-Json
if((DryingTimeFromConfig $cfg @('ABS')) -ne '要確認'){throw 'Unconfigured material cannot use assumed hours'}
$cfg.drying.time_by_material.ABS='2h'
if((DryingTimeFromConfig $cfg @('ABS')) -ne '2h'){throw 'Configured duration not written'}
$cfg.drying.time_by_material.PC='3h'
if((DryingTimeFromConfig $cfg @('ABS','PC')) -ne '要確認'){throw 'Conflicting grades must require confirmation'}
$cfg.drying.time_by_material.PC='2h'
if((DryingTimeFromConfig $cfg @('ABS','PC')) -ne '2h'){throw 'Equal grade durations can be applied'}
$main=Get-Content (Join-Path $PSScriptRoot '..\scripts\update_excel.ps1') -Raw -Encoding UTF8
if($main -notmatch 'SetScheduleBadges' -or $main -notmatch 'DryingEntries' -or $main -notmatch 'UpdateDryingTable'){throw 'Drying badges and table updates must both run'}
$agent=Get-Content (Join-Path $PSScriptRoot '..\src\main.rs') -Raw -Encoding UTF8
if($agent -notmatch 'include_str!\("\.\./scripts/drying\.ps1"\)'){throw 'Auto-update EXE must embed drying.ps1; legacy updater only replaces EXE and XLL'}
Write-Output 'DRYING_TESTS_OK'

$fn=$ast.Find({param($x) $x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'GetDryingMaterialsForEvent'},$true)
if(!$fn){throw 'Missing drying identity matcher'}
Invoke-Expression $fn.Extent.Text
$e=[pscustomobject]@{id='13:mold:ORD1:1:TOOL1:MOLD1';kind='mold';shift='day';machine='13';scheduled_at='2026-11-25T10:00:00'}
$entries=@(
 [pscustomobject]@{id='13:drying:ORD2:1:PART2:TOOL1:MOLD1';machine='13';shift='day';scheduled_at=$e.scheduled_at;materials=@('ABS')},
 [pscustomobject]@{id='13:drying:ORD1:1:PART1:TOOL1:MOLD1';machine='13';shift='day';scheduled_at=$e.scheduled_at;materials=@('PC','AS')},
 [pscustomobject]@{id='13:drying:ORD1:1:PART1:TOOL1:MOLD1';machine='13';shift='night';scheduled_at=$e.scheduled_at;materials=@('ナイロン')}
)
$result=@(GetDryingMaterialsForEvent $e $entries)
if(($result -join ',') -ne 'PC,AS'){throw 'Drying must match the exact order, not only the machine/time'}
$e.id='13:color:ORD1:1:TOOL1:MOLD1';$e.kind='color'
if((@(GetDryingMaterialsForEvent $e $entries) -join ',') -ne 'PC,AS'){throw 'Color changes must also get drying badges'}
$e.scheduled_at='2026-11-25T11:00:00'
if(@(GetDryingMaterialsForEvent $e $entries).Count -ne 0){throw 'Different schedule time must not get a badge'}
Write-Output 'DRYING_IDENTITY_TESTS_OK'
