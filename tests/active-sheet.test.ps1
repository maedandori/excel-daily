$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PSScriptRoot '..\scripts\update_excel.ps1'
$tokens=$null;$errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath).Path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
foreach($name in @('NormalizeName','ResolveShiftForSheet','PickSheetExactly')){
 $func=$ast.Find({param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
 },$true)
 if(!$func){throw "Missing $name"}
 Invoke-Expression $func.Extent.Text
}
$cfg=[pscustomobject]@{sheets=[pscustomobject]@{day='1班　昼勤';night='1班　夜勤 '}}
if((ResolveShiftForSheet $cfg '1班　昼勤') -ne 'day'){throw 'Configured day shift not recognized'}
if((ResolveShiftForSheet $cfg '1班　夜勤 ') -ne 'night'){throw 'Configured night shift not recognized'}
if((ResolveShiftForSheet $cfg '2班　昼勤') -ne 'day'){throw 'Another team day sheet should work'}
if((ResolveShiftForSheet $cfg '2班　夜勤') -ne 'night'){throw 'Another team night sheet should work'}
if((ResolveShiftForSheet $cfg '3班　日勤') -ne 'day'){throw '日勤 must mean day'}
foreach($invalid in @('単なる予定','昼勤・夜勤','予定表')){
 $bad=$false
 try{$null=ResolveShiftForSheet $cfg $invalid}catch{$bad=$true}
 if(!$bad){throw "Ambiguous/unknown shift should fail closed: $invalid"}
}
$day=[pscustomobject]@{Name='1班　昼勤'}
$night=[pscustomobject]@{Name='1班　夜勤'}
$other=[pscustomobject]@{Name='2班　昼勤'}
$book=[pscustomobject]@{Worksheets=@($day,$night,$other)}
$target=PickSheetExactly $book '2班　昼勤'
if(![object]::ReferenceEquals($target,$other)){throw 'Clicked sheet must be selected instead of configured default'}
$target=PickSheetExactly $book '1班　夜勤'
if(![object]::ReferenceEquals($target,$night)){throw 'Night sheet identity changed'}
$bad=$false
try{$null=PickSheetExactly $book '3班　昼勤'}catch{$bad=$true}
if(!$bad){throw 'Missing clicked sheet must fail, not fall back to any other sheet'}
$rust=Get-Content (Join-Path $PSScriptRoot '..\src\main.rs') -Raw -Encoding UTF8
$ribbon=Get-Content (Join-Path $PSScriptRoot '..\ribbon\RibbonController.cs') -Raw -Encoding UTF8
if($rust -notmatch '&\["auto"\]' -or $rust -notmatch '-WorksheetName' -or $rust -notmatch 'sheet_target\.is_some'){
 throw 'Agent is not routing a single clicked sheet'
}
if($ribbon -notmatch 'ActiveSheetTarget\(' -or $ribbon -notmatch 'ClickedShapeTarget\(' -or $ribbon -notmatch '"sheet_name"'){
 throw 'Ribbon/Shape does not send clicked worksheet to Agent'
}
if($rust -notmatch '&\["day", "night"\]'){
 throw 'Legacy hotkey day/night behavior must remain available'
}
Write-Output 'ACTIVE_SHEET_TESTS_OK'
