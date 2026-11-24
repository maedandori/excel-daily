$ErrorActionPreference='Stop'
$source=Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot '../scripts/update_excel.ps1')
foreach($pattern in @('function SetScheduleBadges','color_changed','ExcelDaily managed color badge','ExcelDaily managed drying badge','Shapes.AddShape','Placement=1','HorizontalAlignment=-4131','HorizontalAlignment=-4108','SetScheduleBadges \$sheet \$machineCol \$row','SetScheduleBadges \$sheet \$machineCol \$adopt','SetScheduleBadges \$sheet \$machineCol \$emptyRow')){
 if($source -notmatch $pattern){throw "Missing badge contract: $pattern"}
}
if($source -match 'UpdateDryingTable -sheet'){throw 'Legacy drying table must no longer be written'}
Write-Output 'COLOR_BADGE_CONTRACT_OK'
