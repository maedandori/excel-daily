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
if($main -notmatch 'UpdateDryingTable' -or $main -notmatch 'DryingEntries'){throw 'Main update has lost drying support'}
$agent=Get-Content (Join-Path $PSScriptRoot '..\src\main.rs') -Raw -Encoding UTF8
if($agent -notmatch 'include_str!\("\.\./scripts/drying\.ps1"\)'){throw 'Auto-update EXE must embed drying.ps1; legacy updater only replaces EXE and XLL'}
Write-Output 'DRYING_TESTS_OK'
