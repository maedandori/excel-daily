$ErrorActionPreference='Stop'
$source=Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot '../scripts/update_excel.ps1')
foreach($pattern in @('function SetMoldColorBadge','color_changed','ExcelDaily managed color badge','Shapes.AddShape','Placement=1','SetMoldColorBadge \$sheet \$machineCol \$row','SetMoldColorBadge \$sheet \$machineCol \$adopt','SetMoldColorBadge \$sheet \$machineCol \$emptyRow')){
 if($source -notmatch $pattern){throw "Missing badge contract: $pattern"}
}
Write-Output 'COLOR_BADGE_CONTRACT_OK'
