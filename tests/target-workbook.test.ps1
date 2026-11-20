$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PSScriptRoot '..\scripts\update_excel.ps1'
$tokens=$null;$errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath).Path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
$function=$ast.Find({param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'PickWorkbook'
},$true)
if(!$function){throw 'Missing PickWorkbook'}
Invoke-Expression $function.Extent.Text
$bookA=[pscustomobject]@{Name='A.xlsx';FullName='C:\\plans\\A.xlsx'}
$bookB=[pscustomobject]@{Name='B.xlsx';FullName='C:\\plans\\B.xlsx'}
$excel=[pscustomobject]@{Workbooks=@($bookA,$bookB)}
$cfg=[pscustomobject]@{workbook=[pscustomobject]@{filename_pattern='A*.xlsx';selected_workbook='';choose_if_multiple=$false}}
$selected=PickWorkbook $excel $cfg $bookB.FullName
if($selected.FullName -ne $bookB.FullName){throw 'Explicit Shape target B should override filename pattern A'}
$rejected=$false
try{ $null=PickWorkbook $excel $cfg 'C:\\plans\\missing.xlsx' }catch{$rejected=$true}
if(!$rejected){throw 'Missing clicked workbook must never fall back to A'}
$selected=PickWorkbook $excel $cfg
if($selected.FullName -ne $bookA.FullName){throw 'Original hotkey pattern selection must remain unchanged'}
Write-Output 'TARGET_WORKBOOK_TESTS_OK'
