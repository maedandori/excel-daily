param(
    [ValidateSet('auto','x86','x64')][string]$Bitness='auto',
    [switch]$Remove,
    [switch]$Diagnose
)
$ErrorActionPreference='Stop'
$dir=Split-Path -Parent $MyInvocation.MyCommand.Path
# Office 2019 is version 16.0 (also used by Microsoft 365).
$key='HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
function ReadArchitecture {
    # Most Office 2019 Click-to-Run installs expose the actual Excel bitness here.
    foreach($p in @('HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
                   'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration')){
        try {
            $platform=[string](Get-ItemProperty -LiteralPath $p -Name Platform -ErrorAction Stop).Platform
            if($platform -in @('x86','x64')){return $platform}
        }catch{}
    }
    # MSI installations typically have separate registry views.
    $office32='HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\InstallRoot'
    $office64='HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot'
    if(Test-Path $office32){return 'x86'}
    if(Test-Path $office64){return 'x64'}
    throw 'Excel 2019の32/64ビットを判定できません。ビット数を指定して再実行してください: .\install-ribbon.ps1 -Bitness x86 hoặc x64.'
}

if($Diagnose){
    Write-Output 'Excel Daily リボンの診断（読み取り専用）'
    Write-Output ('インストール先: '+$dir)
    foreach($name in @('excel-daily.exe','settings.json','ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll')){
        Write-Output ($name+': '+(Test-Path -LiteralPath (Join-Path $dir $name)))
    }
    try{ Write-Output ('Excelのビット数: '+(ReadArchitecture)) }
    catch{ Write-Output ('Excelのビット数を判定できません: '+$_.Exception.Message) }
    if(Test-Path -LiteralPath $key){
        $item=Get-Item -LiteralPath $key
        foreach($entry in @($item.GetValueNames()|Where-Object{$_ -match '^OPEN\d*$'})){
            $value=[string]$item.GetValue($entry)
            if($value -like '*ExcelDailyRibbon*.xll*'){
                Write-Output ($entry+' = '+$value)
            }
        }
    } else {
        Write-Output 'Excelのアドイン登録情報が見つかりません。'
    }
    Write-Output ('Excelの起動数: '+@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count)
    Write-Output 'リボンが表示されない場合は、Excelの「ファイル→オプション→アドイン」および「セキュリティセンター」を確認してください。'
    exit 0
}
if($Bitness -eq 'auto' -and -not $Remove){$Bitness=ReadArchitecture}
$xllName=if($Bitness -eq 'x64'){'ExcelDailyRibbon64.xll'}else{'ExcelDailyRibbon.xll'}
$xll=Join-Path $dir $xllName
if(-not $Remove -and -not (Test-Path -LiteralPath $xll)){throw "$xllNameが見つかりません。ZIPをすべて展開してください。"}
if(-not $Remove -and -not (Test-Path -LiteralPath (Join-Path $dir 'excel-daily.exe'))){
    throw 'Excel Daily本体とアドインを同じフォルダーに配置してください。'
}
New-Item -Path $key -Force|Out-Null
$registry=Get-Item -LiteralPath $key
# Removing only our own Ribbon registration protects other Excel Add-ins.
$existing=@($registry.GetValueNames()|Where-Object{$_ -match '^OPEN\d*$'})
$found=0
foreach($name in $existing){
    $value=[string]$registry.GetValue($name)
    if($value -like '*ExcelDailyRibbon*.xll*'){
        if($Remove){
            Remove-ItemProperty -LiteralPath $key -Name $name
            Write-Output "登録を解除しました: $name"
        } else {
            if($value -eq ('/R "'+$xll+'"')){
                Write-Output "登録済み: $xll"
                $found++
            }else{
                # No automatic replacement of an XLL loaded in a running Excel process.
                Write-Output "別のExcel Dailyリボンが登録されています（$name）。先に古い登録を解除してください。"
                throw '既存のアドイン登録と競合しています。'
            }
        }
    }
}
if($Remove){Write-Output 'リボンの登録を解除しました。Excelを再起動してください。';exit 0}
if($found -gt 0){exit 0}
for($i=0;$i -lt 100;$i++){
    $name=if($i -eq 0){'OPEN'}else{'OPEN'+$i}
    if($name -notin $existing){
        New-ItemProperty -LiteralPath $key -Name $name -Value ('/R "'+$xll+'"') -PropertyType String -Force|Out-Null
        Write-Output "Excel Dailyリボンを登録しました（$Bitness, $name）。Excelを再起動すると表示されます。"
        exit 0
    }
}
throw 'Excelのアドイン登録枠が不足しています。'
