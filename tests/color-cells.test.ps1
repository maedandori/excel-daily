$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/color_cells.ps1')
$cases=@(
 @('ｼｰﾙｽﾞGy',0xB5B5B5),@('ｸﾘｱ',0xFFF0DF),@('ｲｵﾝW',0xFFFFFF),
 @('ｼｰﾙｽﾞBk',0x252525),@('ﾀﾞｰｸﾌﾞﾗｳﾝ',0x394C70),
 @('ｲｴﾛｰ',0x7BEBFF),@('ﾁｬｺｰﾙGy',0x666666),@('STC Be',0xBFE4F4)
)
foreach($case in $cases){
 $actual=ExcelDailyColorStyle $case[0]
 if(!$actual -or $actual.Background -ne $case[1]){throw "Wrong color $($case[0])"}
}
if($null -ne (ExcelDailyColorStyle '未登録')){throw 'Unknown color must not be recolored'}
Write-Output 'PASS: Japanese color palette'
