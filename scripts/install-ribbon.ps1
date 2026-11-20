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
    throw 'Không xác định được Excel 2019 32/64 bit. Chạy lại: .\install-ribbon.ps1 -Bitness x86 hoặc x64.'
}
if($Diagnose){
    Write-Output 'Excel Daily Ribbon diagnostics (read-only)'
    Write-Output ('Install folder: '+$dir)
    foreach($name in @('excel-daily.exe','settings.json','ExcelDailyRibbon.xll','ExcelDailyRibbon64.xll')){
        Write-Output ($name+': '+(Test-Path -LiteralPath (Join-Path $dir $name)))
    }
    try{
        Write-Output ('Detected Excel architecture: '+(ReadArchitecture))
    }catch{Write-Output ('Excel architecture unknown: '+$_.Exception.Message)}
    if(Test-Path -LiteralPath $key){
        $item=Get-Item -LiteralPath $key
        foreach($entry in @($item.GetValueNames()|Where-Object{$_ -match '^OPEN\d*$xllName=if($Bitness -eq 'x64'){'ExcelDailyRibbon64.xll'}else{'ExcelDailyRibbon.xll'}
$xll=Join-Path $dir $xllName
if(-not $Remove -and -not (Test-Path -LiteralPath $xll)){throw "Không tìm thấy $xllName trong $dir. Extract the complete v0.3.5 ZIP, not only EXE."}
if(-not $Remove -and -not (Test-Path -LiteralPath (Join-Path $dir 'excel-daily.exe'))){
    throw 'Cần cài Excel Daily Agent cùng thư mục với Add-in.'
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
            Write-Output "Removed $name"
        } else {
            if($value -eq ('/R "'+$xll+'"')){
                Write-Output "Already installed: $xll"
                $found++
            }else{
                # No automatic replacement of an XLL loaded in a running Excel process.
                Write-Output "Another Excel Daily Ribbon is registered ($name). Remove it explicitly first."
                throw 'Existing Excel Daily Ribbon registration conflicts with this installation.'
            }
        }
    }
}
if($Remove){Write-Output 'Ribbon unregistered. Restart Excel.';exit 0}
if($found -gt 0){exit 0}
for($i=0;$i -lt 100;$i++){
    $name=if($i -eq 0){'OPEN'}else{'OPEN'+$i}
    if($name -notin $existing){
        New-ItemProperty -LiteralPath $key -Name $name -Value ('/R "'+$xll+'"') -PropertyType String -Force|Out-Null
        Write-Output "Excel Daily Ribbon registered ($Bitness, $name). Restart Excel to show the tab."
        exit 0
    }
}
throw 'Không còn chỗ trống OPEN trong Excel registry options.'
})){
            $value=[string]$item.GetValue($entry)
            if($value -like '*ExcelDailyRibbon*.xll*'){
                Write-Output ($entry+' = '+$value)
            }
        }
    }else{
        Write-Output 'Excel Options registry key is not present.'
    }
    Write-Output ('Excel processes currently open: '+@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count)
    Write-Output 'If installed but missing, inspect Excel > File > Options > Add-ins > Disabled Items, and Trust Center.'
    exit 0
}
if($Bitness -eq 'auto' -and -not $Remove){$Bitness=ReadArchitecture}
$xllName=if($Bitness -eq 'x64'){'ExcelDailyRibbon64.xll'}else{'ExcelDailyRibbon.xll'}
$xll=Join-Path $dir $xllName
if(-not $Remove -and -not (Test-Path -LiteralPath $xll)){throw "Không tìm thấy $xllName trong $dir"}
if(-not $Remove -and -not (Test-Path -LiteralPath (Join-Path $dir 'excel-daily.exe'))){
    throw 'Cần cài Excel Daily Agent cùng thư mục với Add-in.'
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
            Write-Output "Removed $name"
        } else {
            if($value -eq ('/R "'+$xll+'"')){
                Write-Output "Already installed: $xll"
                $found++
            }else{
                # No automatic replacement of an XLL loaded in a running Excel process.
                Write-Output "Another Excel Daily Ribbon is registered ($name). Remove it explicitly first."
                throw 'Existing Excel Daily Ribbon registration conflicts with this installation.'
            }
        }
    }
}
if($Remove){Write-Output 'Ribbon unregistered. Restart Excel.';exit 0}
if($found -gt 0){exit 0}
for($i=0;$i -lt 100;$i++){
    $name=if($i -eq 0){'OPEN'}else{'OPEN'+$i}
    if($name -notin $existing){
        New-ItemProperty -LiteralPath $key -Name $name -Value ('/R "'+$xll+'"') -PropertyType String -Force|Out-Null
        Write-Output "Excel Daily Ribbon registered ($Bitness, $name). Restart Excel."
        exit 0
    }
}
throw 'Không còn chỗ trống OPEN trong Excel registry options.'
