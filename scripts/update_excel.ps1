param(
    [Parameter(Mandatory=$true)][string]$Settings,
    [Parameter(Mandatory=$true)][ValidateSet("day","night","auto")][string]$Shift,
    [switch]$TimeOnly,
    [ValidateRange(0,1)][int]$DateOffset=0,
    [string]$WorkbookFullName='',
    [string]$WorksheetName=''
)
$ErrorActionPreference = "Stop"
# PowerShell 5.1 otherwise writes stdout/stderr using the OEM code page,
# which corrupts Japanese file paths when Rust reads them as UTF-8.
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Windows.Forms

# Cache Excel COM references; merged cells use their actual top-left value cell.
$script:ExcelCells=@{}
function ExcelCell($sheet,[string]$column,[int]$row){
 $address="$column$row"
 if(-not $script:ExcelCells.ContainsKey($address)){
  $cell=$sheet.Range($address)
  if($cell.MergeCells){$cell=$cell.MergeArea.Cells.Item(1,1)}
  $script:ExcelCells[$address]=$cell
 }
 return $script:ExcelCells[$address]
}
function CellValue($sheet,[string]$column,[int]$row){
 $v=(ExcelCell $sheet $column $row).Value2
 if($null -eq $v){return ""}
 return [string]$v
}
function PutValue($sheet,[string]$column,[int]$row,$value){
 (ExcelCell $sheet $column $row).Value2=$value
}
function PutTextValue($sheet,[string]$column,[int]$row,[string]$value){
 $cell=ExcelCell $sheet $column $row
 $nf=''
 try{$nf=[string]$cell.NumberFormat}catch{}
 if($nf -ne '@' -and $value -match '^[=+\-@\t\r]'){$cell.Value2="'"+$value}
 else{$cell.Value2=$value}
}
# Use the .NET COM dynamic binder with explicitly typed scalar values.
# Avoid PowerShell's Value2 coercion and the failing reflection InvokeMember route.
function SetScheduleRawValue($cell,$value){
 if(-not ('ExcelDailyNativeCellWriter' -as [type])){
  Add-Type -ReferencedAssemblies @('System.Core.dll','Microsoft.CSharp.dll') -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelDailyNativeCellWriter {
 public static void Write(object cell, object value) {
  dynamic range = cell;
  bool com = Marshal.IsComObject(cell);
  if (value == null) {
   if (com) {
    if (Convert.ToBoolean(range.MergeCells)) range.MergeArea.ClearContents();
    else range.ClearContents();
   } else range.Value2 = null;
  } else if (value is double) range.Value2 = (double)value;
  else if (value is string) {
   // Preserve literal text; @-formatted cells must not receive a literal apostrophe.
   string nf = com ? Convert.ToString(range.NumberFormat) : null;
   bool isTextFormat = string.Equals(nf, "@", StringComparison.Ordinal);
   range.Value2 = (com && !isTextFormat) ? "'" + (string)value : (string)value;
  } else if (value is bool) range.Value2 = (bool)value;
  else throw new ArgumentException("Unsupported Excel Value2 type: " + value.GetType().FullName);
  object actual = range.Value2;
  if (!Object.Equals(value, actual)) {
   // Excel may round serials at floating-point precision.
   if (value is double && actual is double &&
       Math.Abs((double)value - (double)actual) <= 1e-12) return;
   throw new InvalidOperationException("Excel Value2 read-back mismatch");
  }
 }
}
'@
 }
 [ExcelDailyNativeCellWriter]::Write($cell.PSObject.BaseObject,$value)
}
function PaintColorCell($sheet,[string]$column,[int]$row,[string]$label){
 if(-not $script:ColorCellsEnabled -or -not $column){return $false}
 $style=ExcelDailyColorStyle $label
 if($null -eq $style){return $false}
 $cell=ExcelCell $sheet $column $row
 if($cell.MergeCells){$cell=$cell.MergeArea}
 $cell.Interior.Color=[int]$style.Background
 $cell.Font.Color=[int]$style.Foreground
 return $true
}

# Managed printable badges; preserve any unrelated workbook Shapes.
# All badges share a single right-aligned column; multiple badges stack vertically.
function GetScheduleBadgeLayout([double]$cellWidth,[double]$cellHeight,[double]$machineWidth,[int]$count){
 if($count -lt 1){return $null}
 $gap=1.0
 $width=[Math]::Min(36.0,$cellWidth-$machineWidth-2.0)
 $height=[Math]::Min(14.0,($cellHeight-2.0-($count-1)*$gap)/$count)
 if($width -lt 7.0 -or $height -lt 6.0){return $null}
 $totalHeight=$height*$count+($count-1)*$gap
 return [pscustomobject]@{width=$width;height=$height;left=$cellWidth-1.0-$width;top=($cellHeight-$totalHeight)/2.0;step=$height+$gap}
}
function SetScheduleBadges($sheet,[string]$column,[int]$row,$event,$cfg){
 if($TimeOnly){return}
 try {
  $cell=ExcelCell $sheet $column $row
  if($cell.MergeCells){$cell=$cell.MergeArea}
  $tags=@()
  if($event.kind -eq 'mold' -and $event.color_changed -eq $true){
   $tags+=@{key='Color';label='色替え';short='色';rgb=0x0067F9;marker='ExcelDaily managed color badge';detail='色替え'}
  }
  if($event.stamp_changed -eq $true){
   $tags+=@{key='Stamp';label='刻印替え';short='刻';rgb=0xB36325;marker='ExcelDaily managed stamp badge';detail='刻印替え'}
  }
  $materials=@()
  if($cfg.drying.enabled -ne $false){
   $materials=@(GetDryingMaterialsForEvent $event @($script:DryingEntries) | Select-Object -Unique)
   if($event.PSObject.Properties['drying_materials']){
    $materials=@($event.drying_materials | Where-Object{$_ -in @('ABS','PS-TH21','PC','ナイロン','AS')} | Select-Object -Unique)
   }
  }
  if($materials.Count){
   $tags+=@{key='Drying';label='乾燥';short='乾';rgb=0x2626DC;marker='ExcelDaily managed drying badge';detail=('乾燥: '+($materials -join ' / '))}
  }
  $active=@{}
  foreach($tag in $tags){$active[$tag.key]=$true}
  $markers=@{Color='ExcelDaily managed color badge';Stamp='ExcelDaily managed stamp badge';Drying='ExcelDaily managed drying badge'}
  foreach($key in $markers.Keys){
   $shape=$null
   try{$shape=$sheet.Shapes.Item("ExcelDaily_${key}Badge_${column}_${row}")}catch{}
   if($null -ne $shape -and $shape.AlternativeText -eq $markers[$key] -and !$active.ContainsKey($key)){$shape.Delete()}
  }
  if(!$tags.Count){$cell.HorizontalAlignment=-4108;return}
  # Reserve space for the machine number, including merged cells.
  $machineWidth=[Math]::Max(14.0,([string]$event.machine).Length*[double]$cell.Font.Size*0.8+5)
  $layout=GetScheduleBadgeLayout ([double]$cell.Width) ([double]$cell.Height) $machineWidth $tags.Count
  if($null -eq $layout){
   foreach($tag in $tags){
    try{$shape=$sheet.Shapes.Item("ExcelDaily_$($tag.key)Badge_${column}_${row}");if($shape.AlternativeText -eq $tag.marker){$shape.Delete()}}catch{}
   }
   $cell.HorizontalAlignment=-4108
   Write-Output "予定マークの表示幅が不足しています: $column$row"
   return
  }
  $cell.HorizontalAlignment=-4131
  $height=$layout.height;$width=$layout.width
  $top=[double]$cell.Top+$layout.top
  $left=[double]$cell.Left+$layout.left
  foreach($tag in $tags){
   $name="ExcelDaily_$($tag.key)Badge_${column}_${row}"
   $shape=$null
   try{$shape=$sheet.Shapes.Item($name)}catch{}
   if($null -ne $shape -and $shape.AlternativeText -ne $tag.marker){$top+=$layout.step;continue}
   if($null -eq $shape){
    $shape=$sheet.Shapes.AddShape(5,[single]$left,[single]$top,[single]$width,[single]$height)
    $shape.Name=$name
    $shape.AlternativeText=$tag.marker
   }
   $shape.Left=$left;$shape.Top=$top;$shape.Width=$width;$shape.Height=$height
   $shape.Placement=1
   $shape.Fill.ForeColor.RGB=[int]$tag.rgb
   $shape.Line.Visible=0
   $label=if($width -ge ($tag.label.Length*6+2)){$tag.label}else{$tag.short}
   $shape.TextFrame.Characters().Text=$label
   $shape.TextFrame.Characters().Font.Color=0xFFFFFF
   $shape.TextFrame.Characters().Font.Bold=$true
   $shape.TextFrame.Characters().Font.Size=[Math]::Min(8,[Math]::Max(4,[Math]::Min($height-2,($width-1)/$label.Length)))
   $shape.TextFrame.HorizontalAlignment=-4108
   $shape.TextFrame.VerticalAlignment=-4108
   $shape.TextFrame.MarginLeft=0;$shape.TextFrame.MarginRight=0
   $shape.TextFrame.MarginTop=0;$shape.TextFrame.MarginBottom=0
   try{$shape.Title=$tag.detail}catch{}
   try{$shape.DrawingObject.PrintObject=$true}catch{}
   if($tag.key -eq 'Drying'){$script:DryingBadgeCount++}
   $top+=$layout.step
  }
 }catch{Write-Output "予定マーク更新失敗: $column$row / $($_.Exception.Message)"}
}
function RemoveUnusedScheduleBadges($sheet,[string]$column,[int]$start,[int]$step,[int]$count){
 if($TimeOnly){return}
 $markers=@{Color='ExcelDaily managed color badge';Stamp='ExcelDaily managed stamp badge';Drying='ExcelDaily managed drying badge'}
 for($i=0;$i -lt $count;$i++){
  $row=$start+$i*$step
  if((CellValue $sheet $column $row) -ne ''){continue}
  foreach($key in $markers.Keys){
   try{$shape=$sheet.Shapes.Item("ExcelDaily_${key}Badge_${column}_${row}");if($shape.AlternativeText -eq $markers[$key]){$shape.Delete()}}catch{}
  }
  $cell=ExcelCell $sheet $column $row
  if($cell.MergeCells){$cell=$cell.MergeArea}
  $cell.HorizontalAlignment=-4108
 }
}

function NormalizeName([string]$name) { return $name.Trim().Replace([char]0x3000, ' ') }
function PickWorkbook($excel, $config, [string]$requestedFullName='') {
    if (-not [string]::IsNullOrWhiteSpace($requestedFullName)) {
        if (-not [IO.Path]::IsPathRooted($requestedFullName)) {
            throw '更新対象のファイルは保存済みの絶対パスで指定してください。'
        }
        # This overrides the filename pattern ONLY for a Shape/Ribbon action.
        # Never pick a different workbook if the clicked one is not open.
        $found=@()
        foreach($opened in $excel.Workbooks) {
            if([string]::Equals([string]$opened.FullName,$requestedFullName,
                    [StringComparison]::OrdinalIgnoreCase)) {
                $found+= $opened
            }
        }
        if($found.Count -ne 1) {
            throw "指定したファイルがExcelで開かれていません: $requestedFullName"
        }
        return $found[0]
    }
    $pat = [string]$config.workbook.filename_pattern
    if (-not $pat) { $pat = '成形スケジュール*.xlsx' }
    $matches = @()
    foreach ($book in $excel.Workbooks) {
        if ($book.Name -like $pat) { $matches += $book }
    }
    if ($matches.Count -eq 0) { throw "一致するExcelファイルが開かれていません: $pat。開いているファイルのみ更新できます。" }
    $selected = [string]$config.workbook.selected_workbook
    if ($selected) {
        $found = @($matches | Where-Object { $_.Name -eq $selected })
        if ($found.Count -eq 1) { return $found[0] }
        throw "選択したファイル「$selected」が開かれていません。他のファイルは変更しません。"
    }
    if ($matches.Count -eq 1) { return $matches[0] }
    # No modal workbook picker. An ambiguous legacy hotkey must stop safely;
    # a clicked Shape or Ribbon button always passes its exact workbook path.
    throw "対象ファイルを特定できません。対象のシート上のボタンから更新するか、設定でファイルを指定してください。"
}
function PickSheet($book, [string]$wanted) {
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -ceq $wanted) { return $sheet }
    }
    # Fallback only for invisible trailing spaces/full-width spaces.
    $normalized = NormalizeName $wanted
    $found = @()
    foreach ($sheet in $book.Worksheets) {
        if ((NormalizeName $sheet.Name) -eq $normalized) { $found += $sheet }
    }
    if ($found.Count -eq 1) { return $found[0] }
    throw "シート「$wanted」が見つかりません。settings.jsonのシート名を確認してください。"
}
# A clicked Shape/Ribbon command is scoped to the currently selected
# sheet, even if the workbook contains many shifts. Detect which source
# schedule to read from the exact configured name or a clear shift marker.
# Do not guess for unrelated or ambiguous sheet names.
function ResolveShiftForSheet($cfg,[string]$name){
    if([string]::IsNullOrWhiteSpace($name)){throw '更新するシート名を指定してください。'}
    $normalized=NormalizeName $name
    $day=NormalizeName ([string]$cfg.sheets.day)
    $night=NormalizeName ([string]$cfg.sheets.night)
    if($normalized -eq $day -and $normalized -ne $night){return 'day'}
    if($normalized -eq $night -and $normalized -ne $day){return 'night'}
    $isDay=($normalized -match '昼勤|日勤')
    $isNight=($normalized -match '夜勤')
    if($isDay -and -not $isNight){return 'day'}
    if($isNight -and -not $isDay){return 'night'}
    throw "シート「$name」の勤務区分（昼勤・夜勤）を判定できません。名前に昼勤または夜勤を含めるか、設定のシート名を確認してください。"
}
function PickSheetExactly($book,[string]$name){
    foreach($sheet in $book.Worksheets){
        if($sheet.Name -ceq $name){return $sheet}
    }
    throw "指定されたシート「$name」が対象ブックに見つかりません。更新は実行されませんでした。"
}
. (Join-Path $PSScriptRoot 'shift_date.ps1')
function Test-ExcelDailyApiUri([string]$raw) {
    try {
        $u = [System.Uri]$raw
        if (-not $u.IsAbsoluteUri -or $u.Scheme -ne 'https' -or $u.AbsolutePath -ne '/api/excel-daily/events' -or $u.UserInfo -or $u.Query -or $u.Fragment) { return $false }
        $h = $u.DnsSafeHost.ToLowerInvariant()
        if (!$h -or $h -eq 'localhost' -or $h.EndsWith('.localhost') -or $h.EndsWith('.local') -or $h.EndsWith('.internal') -or $h.Contains(':')) { return $false }
        $ip = $null
        if ([System.Net.IPAddress]::TryParse($h, [ref]$ip)) {
            $b = $ip.GetAddressBytes()
            if ($b.Length -ne 4 -or $b[0] -eq 0 -or $b[0] -eq 10 -or $b[0] -eq 127 -or ($b[0] -eq 169 -and $b[1] -eq 254) -or ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) -or ($b[0] -eq 192 -and $b[1] -eq 168) -or ($b[0] -eq 255 -and $b[1] -eq 255)) { return $false }
        }
        return $true
    } catch { return $false }
}
function Resolve-ExcelDailyToken([string]$raw) {
    if (!$raw) { return '' }
    $t = $raw.Trim()
    if ($t.StartsWith('dpapi:')) {
        Add-Type -AssemblyName System.Security
        $enc = [Convert]::FromBase64String($t.Substring(6))
        $plain = [System.Security.Cryptography.ProtectedData]::Unprotect($enc, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Text.Encoding]::UTF8.GetString($plain)
    }
    return $t
}
function FetchEvents($cfg, [string]$baseDir, [string]$shift,[int]$dateOffset=0) {
    $script:DryingEntries=@()
    $url = [string]$cfg.data_source.api_url
    if ($url) {
        $token = Resolve-ExcelDailyToken ([string]$cfg.data_source.api_token)
        # API must return {"events":[...]} with the documented fields.
        if (!$token) { throw '認証トークンが設定されていません。Excel Daily用トークンを設定してください。' }
        if (-not (Test-ExcelDailyApiUri $url)) { throw 'APIの接続先URLが正しくありません（HTTPS / /api/excel-daily/events のみ許可）。' }
        $headers = @{ Authorization = "Bearer $token" }
        $date = ShiftDate $shift $dateOffset
        $separator = if ($url.Contains('?')) { '&' } else { '?' }
        $requestUrl = $url + $separator + 'shift=' + [uri]::EscapeDataString($shift) + '&date=' + [uri]::EscapeDataString($date)
        if ($TimeOnly) { $requestUrl += '&time_only=1' }
        try {
                        # Decode original HTTP bytes, not Windows PowerShell 5.1's legacy text charset.
            $httpResponse = Invoke-WebRequest -UseBasicParsing -Uri $requestUrl -Headers $headers -Method Get -TimeoutSec 15 -MaximumRedirection 0
            if (-not $httpResponse.RawContentStream) { throw 'APIからデータを取得できませんでした。' }
            $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
            $payload = $utf8.GetString($httpResponse.RawContentStream.ToArray())
            $response = $payload | ConvertFrom-Json
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            if ($statusCode -eq 401) { throw 'API認証に失敗しました（401）。トークンを確認してください。' }
            if ($statusCode -eq 503) { throw 'APIが利用できません（503）。Workerの設定とMICSの接続状態を確認してください。' }
            throw
        }
        if ($null -eq $response.events) { throw "API応答に予定データがありません。Excelは変更しません。" }
        if ($null -ne $response.drying) { $script:DryingEntries=@($response.drying) }
        return @($response.events)
    }
    $path = [string]$cfg.data_source.events_file
    if (-not $path) { throw "データ取得先を設定してください（events_fileまたはapi_url）。" }
    if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path $baseDir $path }
    if (-not (Test-Path -LiteralPath $path)) { throw "予定データが見つかりません: 「$path」。Excelは変更しません。" }
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $json.events) { throw "予定データにevents配列がありません。" }
    if ($null -ne $json.drying) { $script:DryingEntries=@($json.drying) }
    return @($json.events)
}
# Repair only visibly corrupted color labels in rows already tracked by Excel Daily.
function LooksMojibake([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    return ($value.Contains([char]0xFFFD) -or $value.Contains([char]0x00EF) -or
        $value.Contains([char]0x00E3) -or $value.Contains([char]0x00C3) -or
        $value.Contains([char]0x00C2) -or $value.Contains('縺') -or $value.Contains('繧'))
}
function RepairColorIfBroken($sheet, $cfg, $event, [int]$row) {
    $pairs = @(
        @([string]$cfg.columns.color_before,[string]$event.color_before),
        @([string]$cfg.columns.color_after,[string]$event.color_after)
    )
    foreach ($pair in $pairs) {
        $column = [string]$pair[0]
        $correct = [string]$pair[1]
        if ($column -and $correct -and (LooksMojibake (CellValue $sheet $column $row))) {
            PutTextValue $sheet $column $row $correct
            $null=PaintColorCell $sheet $column $row $correct
        }
    }
}
# Ctrl+1, Ctrl+2 and Ctrl+3 share this rounding rule.
. (Join-Path $PSScriptRoot 'round_time.ps1')
function EventTime($event) {
    $t = [string]$event.time
    try { return RoundEventTimeUp $t }
    catch { throw "時刻が正しくありません ($($event.id)): $t" }
}
. (Join-Path $PSScriptRoot 'color_cells.ps1')
. (Join-Path $PSScriptRoot 'drying.ps1')
# Recover rows if a workbook was already filled manually or the Agent state
# was not available. Never overwrite ambiguous or previously tracked rows.
function ExcelTimeText($sheet,[string]$column,[int]$row) {
    $value=(ExcelCell $sheet $column $row).Value2
    if ($null -eq $value) { return '' }
    if ($value -is [double] -or $value -is [decimal]) {
        $number=[double]$value
        if ($number -ge 0 -and $number -lt 1) {
            $minutes=[int][Math]::Round($number*1440)
            return ('{0:00}:{1:00}' -f ([int][Math]::Floor($minutes/60)%24),($minutes%60))
        }
    }
    $raw=[string]$value
    if ($raw -match '^\s*(\d{1,2}):(\d{2})') {
        return ('{0:00}:{1:00}' -f [int]$matches[1],[int]$matches[2])
    }
    return ''
}
function MachineId([string]$value) {
    $digits=($value.Trim() -replace '[^0-9]','')
    return ($digits -replace '^0+','')
}
function TimeStillPending([string]$existing,[datetime]$scheduled,[datetime]$now) {
    try {
        $parsed=[TimeSpan]::Parse($existing,[Globalization.CultureInfo]::InvariantCulture)
        $candidate=$scheduled.Date.Add($parsed)
        if (($candidate-$scheduled).TotalHours -gt 12) { $candidate=$candidate.AddDays(-1) }
        if (($candidate-$scheduled).TotalHours -lt -12) { $candidate=$candidate.AddDays(1) }
        return $candidate -gt $now
    } catch { return $false }
}
# Only a server-authorized native-shift history event can fill an empty past slot.
# Time-only refresh, legacy JSON and completed/manual records remain protected.
function CanRefreshTrackedTime($event,[bool]$past,[bool]$carryover,[bool]$preservePast,$previous,[datetime]$now) {
    if($event.completed -eq $true){return $false}
    return $true
}
function CanAppendScheduleEvent($event,[bool]$past,[bool]$timeOnly) {
    if($timeOnly){return $false}
    if(-not $past){return $true}
    return ($event.history_replay -eq $true -and $event.carryover -ne $true)
}

function FindExistingSlot($sheet,$cfg,$event,$state,$book,[int]$start,[int]$step,[int]$count,
                          [object[]]$ordered,[datetime]$now,[datetime]$scheduled,[bool]$carryover) {
    $kind=[string]$event.kind
    $machine=[string]$event.machine
    $machineCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_machine}else{[string]$cfg.columns.color_machine}
    $timeCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_time}else{[string]$cfg.columns.color_time}
    $targetTime=([TimeSpan]::Parse((EventTime $event))).ToString('hh\:mm')
    $prefix="$($book.FullName)|$($sheet.Name)|$kind|"
    $incoming=@($ordered|Where-Object{
        [string]$_.kind -eq $kind -and (MachineId ([string]$_.machine)) -eq (MachineId $machine)
    })
    $liveKeys=@{}
    foreach($item in $ordered){
        if([string]$item.kind -eq $kind){$liveKeys[$prefix+[string]$item.id]=$true}
    }
    $occupied=@{}
    $stale=@{}
    foreach($entry in $state.GetEnumerator()) {
        if(([string]$entry.Key).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and $entry.Value.row) {
            $r=[int]$entry.Value.row
            if($liveKeys.ContainsKey([string]$entry.Key)){$occupied[$r]=$true}
            else{$stale[$r]=$true}
        }
    }
    $candidates=@()
    $unresolved=$false
    for($i=0;$i -lt $count;$i++) {
        $row=$start+$i*$step
        if($occupied.ContainsKey($row)) { continue }
        if((MachineId (CellValue $sheet $machineCol $row)) -ne (MachineId $machine)){continue}
        $time=ExcelTimeText $sheet $timeCol $row
        if(-not $time){$unresolved=$true;continue}
        if($stale.ContainsKey($row) -and ($incoming.Count -ne 1 -or (!$carryover -and !(TimeStillPending $time $scheduled $now)))){
            $unresolved=$true;continue
        }
        $colorMatched=$false
        if($kind -eq 'color') {
            $old=[string]$event.color_before
            $new=[string]$event.color_after
            $actualOld=CellValue $sheet ([string]$cfg.columns.color_before) $row
            $actualNew=CellValue $sheet ([string]$cfg.columns.color_after) $row
            if($old -and $new -and $actualOld -and $actualNew) {
                $colorMatched=($actualOld.Trim() -eq $old.Trim() -and $actualNew.Trim() -eq $new.Trim())
                if(-not $colorMatched){$unresolved=$true;continue}
            }
        }
        $candidates+=@{row=$row;time=$time;colorMatched=$colorMatched}
    }
    if($candidates.Count -eq 0){if($unresolved){return -1};return 0}
    $exact=@($candidates|Where-Object{$_.time -eq $targetTime})
    if($exact.Count -eq 1){return [int]$exact[0].row}
    if($exact.Count -gt 1){return -1}
    $sameColors=@($candidates|Where-Object{$_.colorMatched})
    if($sameColors.Count -eq 1 -and ($carryover -or (TimeStillPending $sameColors[0].time $scheduled $now))) {
        return [int]$sameColors[0].row
    }
    if($sameColors.Count -gt 1){return -1}
    if($candidates.Count -eq 1 -and $incoming.Count -eq 1 -and
       ($carryover -or (TimeStillPending $candidates[0].time $scheduled $now))) {
        return [int]$candidates[0].row
    }
    return -1
}

function SetScheduleRowIdentity($state,[string]$key,[string]$prefix,[int]$row,[string]$scheduledAt){
    foreach($oldKey in @($state.Keys)){
        if($oldKey -ne $key -and ([string]$oldKey).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and [int]$state[$oldKey].row -eq $row){
            $state.Remove($oldKey)
        }
    }
    $state[$key]=@{row=$row;scheduled_at=$scheduledAt}
}

# Derive chronological order from the visible sheet after refreshing times.
# A night shift crosses midnight; 18:30 is the earliest normal lookback slot.
function VisibleScheduleDate([string]$time,[datetime]$date,[string]$shift){
 if($time -notmatch '^([01]?\d|2[0-3]):([0-5]\d)\z'){return $null}
 $minutes=[int]$Matches[1]*60+[int]$Matches[2]
 if($shift -eq 'night' -and $minutes -lt 1110){$minutes+=1440}
 return $date.Date.AddMinutes($minutes)
}
# Recognize row-local completion marks. Do not drag an ambiguous drawing.
function GetScheduleRowMarkInfo($sheet,[string]$machineCol,[string]$timeCol,[int]$row){
 $marked=$false
 $unsafe=$false
 $marks=@()
 foreach($column in @($machineCol,$timeCol)){
  $cell=ExcelCell $sheet $column $row
  if($cell.Font.Strikethrough -eq $true){$marked=$true}
 }
 $cell=ExcelCell $sheet $machineCol $row
 if($cell.MergeCells){$cell=$cell.MergeArea}
 $timeCell=ExcelCell $sheet $timeCol $row
 if($timeCell.MergeCells){$timeCell=$timeCell.MergeArea}
 $cx=[double]$cell.Left+[double]$cell.Width/2
 $rowTop=[double]$cell.Top
 $rowBottom=$rowTop+[double]$cell.Height
 $cy=($rowTop+$rowBottom)/2
 $tolerance=[double]$cell.Height*0.35
 foreach($shape in $sheet.Shapes){
  try{
   if([string]$shape.AlternativeText -like 'ExcelDaily managed * badge'){continue}
   $left=[double]$shape.Left;$top=[double]$shape.Top
   $w=[double]$shape.Width;$h=[double]$shape.Height
   if($w -lt [double]$cell.Width*0.5 -or $left -gt $cx -or ($left+$w) -lt $cx -or
      $top -gt ($cy+$tolerance) -or ($top+$h) -lt ($cy-$tolerance)){continue}
   $marked=$true
   $right=[double]$timeCell.Left+[double]$timeCell.Width
   if([int]$shape.Type -notin @(5,9,22,23) -or $top -lt $rowTop -or
      ($top+$h) -gt $rowBottom -or $h -gt [double]$cell.Height*0.8 -or
      ($left+$w) -gt ($right+3) -or $left -lt ([double]$cell.Left-[double]$cell.Width*0.75)){
    $unsafe=$true;continue
   }
   $marks+=@{shape=$shape;top=$top;dy=($top-$rowTop);height=$h}
  }catch{$unsafe=$true}
 }
 return [pscustomobject]@{marked=$marked;unsafe=$unsafe;shapes=$marks}
}
function TestScheduleRowMarked($sheet,[string]$machineCol,[string]$timeCol,[int]$row){
 $info=GetScheduleRowMarkInfo $sheet $machineCol $timeCol $row
 return ($info.marked -or $info.unsafe)
}
function GetVisiblePendingSchedule($sheet,$cfg,$state,$book,[int]$start,[int]$step,[int]$count,[datetime]$date,[string]$shift,[datetime]$now,[object[]]$events){
 foreach($kind in @('mold','color')){
  $machineCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_machine}else{[string]$cfg.columns.color_machine}
  $timeCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_time}else{[string]$cfg.columns.color_time}
  $prefix="$($book.FullName)|$($sheet.Name)|$kind|"
  for($i=0;$i -lt $count;$i++){
   $row=$start+$i*$step
   $machine=MachineId (CellValue $sheet $machineCol $row)
   if(!$machine){continue}
   $when=VisibleScheduleDate (ExcelTimeText $sheet $timeCol $row) $date $shift
   if($null -eq $when){continue}
   # Completion protects the time value, not its position in the schedule.
   $mark=GetScheduleRowMarkInfo $sheet $machineCol $timeCol $row
   if($mark.unsafe){
    Write-Verbose "予定行のマークを安全に移動できません: $kind / $row"
    continue
   }
   [pscustomobject]@{kind=$kind;row=$row;key='';when=$when;scheduled_at=$when.ToString('s');prefix=$prefix}
  }
 }
}
# Plan a chronological permutation of safely movable visible rows.
function PlanPendingScheduleOrder([object[]]$pending){
 foreach($kind in @('mold','color')){
  # Normalize hashtables to explicit objects before PowerShell 5.1 sorting.
  $items=@(foreach($p in $pending){
   if([string]$p.kind -eq $kind){
    [pscustomobject]@{kind=$kind;row=[int]$p.row;key=[string]$p.key;when=[datetime]$p.when;scheduled_at=[string]$p.scheduled_at;prefix=[string]$p.prefix}
   }
  })
  if($items.Count -lt 2){continue}
  $slots=@($items | ForEach-Object{[int]$_.row} | Sort-Object -Unique)
  if($slots.Count -ne $items.Count){continue}
  $ordered=@($items | Sort-Object -Property @{Expression={$_.when};Ascending=$true},@{Expression={$_.row};Ascending=$true})
  for($i=0;$i -lt $slots.Count;$i++){
   [pscustomobject]@{item=$ordered[$i];destination=[int]$slots[$i]}
  }
 }
}
# Move managed badges and positively identified row-local completion marks.
function MovePendingScheduleOrder($sheet,$cfg,$state,[object[]]$pending){
 $moves=0
 foreach($kind in @('mold','color')){
  $plan=@(PlanPendingScheduleOrder @($pending|Where-Object{$_.kind -eq $kind}))
  $moving=@($plan|Where-Object{[int]$_.item.row -ne [int]$_.destination})
  if(!$moving.Count){continue}
  $machineCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_machine}else{[string]$cfg.columns.color_machine}
  $timeCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_time}else{[string]$cfg.columns.color_time}
  $columns=@($machineCol,$timeCol)
  $colorColumns=@()
  if($kind -eq 'color'){
   $colorColumns=@([string]$cfg.columns.color_before,[string]$cfg.columns.color_after)
   $columns+=@($colorColumns)
  }
  if(@($columns|Where-Object{!$_}).Count -or @($columns|Select-Object -Unique).Count -ne $columns.Count){continue}
  $markers=@{Color='ExcelDaily managed color badge';Stamp='ExcelDaily managed stamp badge';Drying='ExcelDaily managed drying badge'}
  $snap=@{}
  $shapes=@()
  $completionMarks=@()
  $addresses=@{}
  $blocked=$false
  # Stage every source before overwriting a single cell. Preserve number formats,
  # borders and merged-cell layout. Only transfer values and event-specific colors.
  foreach($move in $moving){
   $row=[int]$move.item.row
   $cells=@{}
   foreach($col in $columns){
    $cell=ExcelCell $sheet $col $row
    $address="$($cell.Row):$($cell.Column)"
    if($addresses.ContainsKey($address) -or $cell.HasFormula -or ($sheet.ProtectContents -and $cell.Locked)){
     $blocked=$true;break
    }
    $addresses[$address]=$true
    $rawVal=$cell.Value2
    if($null -ne $rawVal -and -not ($rawVal -is [double] -or $rawVal -is [string] -or $rawVal -is [bool])){$blocked=$true;break}
    $entry=@{value=$rawVal}
    $range=$cell
    if($cell.MergeCells){$range=$cell.MergeArea}
    if($null -eq $range.Font.Strikethrough -or $range.Font.Strikethrough -is [DBNull]){$blocked=$true;break}
    $entry.strike=[bool]$range.Font.Strikethrough
    if($col -eq $machineCol){
     $align=$range.HorizontalAlignment
     if($null -eq $align -or $align -is [DBNull]){$blocked=$true;break}
     $entry.alignment=$align
    }
    if($colorColumns -contains $col){
     $pat=$range.Interior.Pattern
     $bg=$range.Interior.Color
     $fg=$range.Font.Color
     if($null -eq $pat -or $pat -is [DBNull] -or $null -eq $bg -or $bg -is [DBNull] -or $null -eq $fg -or $fg -is [DBNull]){$blocked=$true;break}
     $entry.pattern=$pat
     $entry.background=$bg
     $entry.foreground=$fg
    }
    $cells[$col]=$entry
   }
   if($blocked){break}
   $snap[$row]=$cells
   $machineCell=ExcelCell $sheet $machineCol $row
   if($machineCell.MergeCells){$machineCell=$machineCell.MergeArea}
   $markInfo=GetScheduleRowMarkInfo $sheet $machineCol $timeCol $row
   if($markInfo.unsafe -or ($markInfo.shapes.Count -and $sheet.ProtectDrawingObjects)){$blocked=$true;break}
   $target=ExcelCell $sheet $machineCol ([int]$move.destination)
   if($target.MergeCells){$target=$target.MergeArea}
   foreach($mark in $markInfo.shapes){
    if($mark.dy+$mark.height -gt [double]$target.Height){$blocked=$true;break}
    $completionMarks+=@{shape=$mark.shape;top=$mark.top;targetTop=([double]$target.Top+$mark.dy)}
   }
   if($blocked){break}
   foreach($tag in $markers.Keys){
    $name="ExcelDaily_${tag}Badge_${machineCol}_${row}"
    $shape=$null
    try{$shape=$sheet.Shapes.Item($name)}catch{}
    if($null -eq $shape){continue}
    if($shape.AlternativeText -ne $markers[$tag] -or $sheet.ProtectDrawingObjects){$blocked=$true;break}
    $shapes+=@{shape=$shape;name=$name;tag=$tag;row=$row;destination=[int]$move.destination;
      left=[double]$shape.Left;top=[double]$shape.Top;
      dx=([double]$shape.Left-[double]$machineCell.Left);dy=([double]$shape.Top-[double]$machineCell.Top);
      width=[double]$shape.Width;height=[double]$shape.Height}
   }
   if($blocked){break}
  }
  if($blocked){Write-Output "並べ替え対象に保護セル・数式・未管理マークがあります: $kind";continue}
  # Destination slots are a permutation of source slots; temporary names
  # avoid name collisions when two adjacent machines exchange places.
  $writeStep='shape names'
  $attemptedCells=@{}
  $attemptedMarks=@()
  $attemptedBadges=@()
  try{
   foreach($badge in $shapes){$attemptedBadges+=,$badge;$badge.shape.Name='ExcelDaily_Tmp_'+[guid]::NewGuid().ToString('N')}
   foreach($move in $moving){
    $source=$snap[[int]$move.item.row]
    $row=[int]$move.destination
    foreach($col in $columns){
     $cell=ExcelCell $sheet $col $row
     $v=$source[$col]
     $writeStep="Value2 $col$row"
     $attemptedCells["$col$row"]=$true
     SetScheduleRawValue $cell $v.value
     $writeStep="format $col$row"
     $range=$cell
     if($cell.MergeCells){$range=$cell.MergeArea}
     $range.Font.Strikethrough=$v.strike
     if($col -eq $machineCol){$range.HorizontalAlignment=[int]$v.alignment}
     if($colorColumns -contains $col){
      $range.Interior.Color=[int]$v.background
      $range.Interior.Pattern=[int]$v.pattern
      $range.Font.Color=[int]$v.foreground
     }
    }
   }
   $writeStep='badge positions'
   foreach($badge in $shapes){
    $target=ExcelCell $sheet $machineCol $badge.destination
    if($target.MergeCells){$target=$target.MergeArea}
    $badge.shape.Left=[single]([double]$target.Left+$badge.dx)
    $badge.shape.Top=[single]([double]$target.Top+$badge.dy)
    $badge.shape.Name="ExcelDaily_$($badge.tag)Badge_${machineCol}_$($badge.destination)"
   }
   $writeStep='completion marks'
   foreach($mark in $completionMarks){$attemptedMarks+=,$mark;$mark.shape.Top=[single]$mark.targetTop}
  }catch{
   $failure=$_
   $rollbackErrors=@()
   foreach($mark in $attemptedMarks){try{$mark.shape.Top=[single]$mark.top}catch{$rollbackErrors+=$_.Exception.Message}}
   # Best-effort rollback to the post-time-update values and original badges.
   foreach($badge in $attemptedBadges){try{$badge.shape.Name='ExcelDaily_Tmp_'+[guid]::NewGuid().ToString('N')}catch{$rollbackErrors+=$_.Exception.Message}}
   foreach($row in $snap.Keys){
    foreach($col in $columns){
     if(-not $attemptedCells.ContainsKey("$col$row")){continue}
     try{
      $cell=ExcelCell $sheet $col $row;$v=$snap[$row][$col];SetScheduleRawValue $cell $v.value
      $range=$cell;if($cell.MergeCells){$range=$cell.MergeArea}
      $range.Font.Strikethrough=$v.strike
      if($col -eq $machineCol){$range.HorizontalAlignment=[int]$v.alignment}
      if($colorColumns -contains $col){$range.Interior.Color=[int]$v.background;$range.Interior.Pattern=[int]$v.pattern;$range.Font.Color=[int]$v.foreground}
     }catch{$rollbackErrors+="${col}${row}: $($_.Exception.Message)"}
    }
   }
   foreach($badge in $attemptedBadges){try{$badge.shape.Left=[single]$badge.left;$badge.shape.Top=[single]$badge.top;$badge.shape.Name=$badge.name}catch{$rollbackErrors+=$_.Exception.Message}}
   $rollbackInfo=if($rollbackErrors.Count){' 復元失敗: '+($rollbackErrors -join '; ')}else{''}
   throw "予定の並べ替えに失敗しました。保存せずにシートを確認してください: [$kind / $writeStep] $($failure.Exception.Message)$rollbackInfo"
  }
  # Remap existing identities using original rows, without assigning IDs to manual rows.
  $rowMap=@{}
  foreach($move in $moving){$rowMap[[int]$move.item.row]=[int]$move.destination}
  $prefix=[string]$moving[0].item.prefix
  foreach($key in @($state.Keys)){
   if(!$prefix -or !([string]$key).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)){continue}
   $old=$state[$key]
   if($old.row -and $rowMap.ContainsKey([int]$old.row)){
    $old.row=$rowMap[[int]$old.row]
   }
  }
  $moves+=$moving.Count
 }
 $script:ReorderedCount=$moves
}

$cfg = Get-Content -LiteralPath $Settings -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not [string]::IsNullOrWhiteSpace($WorksheetName)){
    if([string]::IsNullOrWhiteSpace($WorkbookFullName)){
        throw 'シートを指定する場合は更新対象のExcelファイルも指定してください。'
    }
    if($Shift -ne 'auto'){
        throw 'シートを指定する場合はShift=autoを使用してください。'
    }
    $Shift=ResolveShiftForSheet $cfg $WorksheetName
}elseif($Shift -eq 'auto'){
    throw '更新するシートが指定されていません。'
}
$script:ColorCellsEnabled=($null -eq $cfg.formatting -or $cfg.formatting.color_cells -ne $false)
$fastWrites=($null -eq $cfg.formatting -or $cfg.formatting.fast_excel_writes -ne $false)
$baseDir = Split-Path -Parent $Settings
$recoveryLock=Join-Path $baseDir 'excel-daily-recovery.lock'
if(Test-Path -LiteralPath $recoveryLock){
 throw '前回の更新が中断されました。Excelを保存せずに確認し、安全を確認した後に excel-daily-recovery.lock を削除してください。'
}
$events = @(FetchEvents $cfg $baseDir $Shift $DateOffset)
# Validate entire input before touching Excel.
$usable = @()
foreach ($e in $events) {
    if ([string]$e.shift -ne $Shift) { continue }
    # A carried-over event must explicitly target this shift. It retains its
    # stable event ID and original timestamp, despite the previous shift date.
    if ([string]$e.kind -notin @("mold","color")) { throw "予定種別が正しくありません: $($e.kind)" }
    if (-not $e.id -or -not $e.machine -or -not $e.scheduled_at) {
        throw "予定にはid、machine、scheduled_atが必要です。"
    }
    $null = EventTime $e
    $usable += $e
}
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
if ($null -eq $excel) { throw "Excelが起動していません。" }
$book = PickWorkbook $excel $cfg $WorkbookFullName
$sheetName = if (-not [string]::IsNullOrWhiteSpace($WorksheetName)) {
    $WorksheetName
} elseif ($Shift -eq "day") { [string]$cfg.sheets.day } else { [string]$cfg.sheets.night }
$sheet = if (-not [string]::IsNullOrWhiteSpace($WorksheetName)) {
    PickSheetExactly $book $sheetName
} else {
    PickSheet $book $sheetName
}
$start = if ($Shift -eq "day") { [int]$cfg.rows.day_start } else { [int]$cfg.rows.night_start }
$step = [int]$cfg.rows.step
$count = [int]$cfg.rows.max_entries
if ($step -lt 1 -or $count -lt 1 -or $count -gt 100) { throw "行の設定が正しくありません。" }
$metadataPath = Join-Path $baseDir "excel-daily-state.json"
$state = @{}
if (Test-Path -LiteralPath $metadataPath) {
    $loaded = Get-Content -LiteralPath $metadataPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($loaded) { foreach ($p in $loaded.PSObject.Properties) { $state[$p.Name] = $p.Value } }
}
$changed = 0
$added = 0
$kept = 0
$now = [DateTime]::Now
$ordered = @($usable | Sort-Object { [datetime]::Parse([string]$_.scheduled_at, [Globalization.CultureInfo]::InvariantCulture) })
$started=[DateTime]::UtcNow
$colored=0
$nextSlot=@{mold=0;color=0}
$script:DryingBadgeCount=0
$script:ReorderedCount=0
if (-not ('ExcelDailyComMessageFilter' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IExcelDailyOleMessageFilter {
 [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
 [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
 [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}
public sealed class ExcelDailyComMessageFilter : IExcelDailyOleMessageFilter {
 [DllImport("Ole32.dll")]
 private static extern int CoRegisterMessageFilter(IExcelDailyOleMessageFilter newFilter, out IExcelDailyOleMessageFilter oldFilter);
 [ThreadStatic] private static IExcelDailyOleMessageFilter _prior;
 [ThreadStatic] private static bool _registered;
 public static void Register() {
  if (_registered || System.Threading.Thread.CurrentThread.GetApartmentState() != System.Threading.ApartmentState.STA) return;
  IExcelDailyOleMessageFilter old;
  if (CoRegisterMessageFilter(new ExcelDailyComMessageFilter(), out old) == 0) { _prior = old; _registered = true; }
 }
 public static void Revoke() {
  if (!_registered) return;
  IExcelDailyOleMessageFilter dummy;
  CoRegisterMessageFilter(_prior, out dummy);
  _prior = null; _registered = false;
 }
 int IExcelDailyOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
 int IExcelDailyOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) {
  if (dwRejectType == 2 && dwTickCount < 5000) return 120;
  return -1;
 }
 int IExcelDailyOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
'@
}
[ExcelDailyComMessageFilter]::Register()
if (-not $excel.Ready) {
 [ExcelDailyComMessageFilter]::Revoke()
 throw 'Excelが編集中のため更新できません。セルの編集を確定してから再実行してください。'
}
# Persistent fail-closed marker before workbook mutations. Only success removes it.
[IO.File]::WriteAllText($recoveryLock,("$($book.FullName)`r`n$($sheet.Name)`r`n$([DateTime]::Now.ToString('o'))"),(New-Object System.Text.UTF8Encoding($false)))
$screenChanged=$false;$eventsChanged=$false
if($fastWrites){
 try{$oldScreen=[bool]$excel.ScreenUpdating;$excel.ScreenUpdating=$false;$screenChanged=$true}catch{}
 try{$oldEvents=[bool]$excel.EnableEvents;$excel.EnableEvents=$false;$eventsChanged=$true}catch{}
}
try {
foreach ($e in $ordered) {
    $kind = [string]$e.kind
    $machine = [string]$e.machine
    $id = [string]$e.id
    $key = "$($book.FullName)|$($sheet.Name)|$kind|$id"
    $machineCol = if ($kind -eq "mold") { [string]$cfg.columns.mold_machine } else { [string]$cfg.columns.color_machine }
    $timeCol = if ($kind -eq "mold") { [string]$cfg.columns.mold_time } else { [string]$cfg.columns.color_time }
    $when = [DateTime]::Parse([string]$e.scheduled_at, [System.Globalization.CultureInfo]::InvariantCulture)
    $completed = ($e.completed -eq $true)
    $handoverStatuses = @('100','410','440','490')
    $carryover = ($e.carryover -eq $true) -and
        ($handoverStatuses -contains ([string]$e.status_code)) -and (-not $completed)
    $past = (($when -le $now) -and (-not $carryover)) -or $completed
    $previous = $state[$key]
    $prefix="$($book.FullName)|$($sheet.Name)|$kind|"
    if ($previous) {
        $row = [int]$previous.row
        if ($row -lt $start -or (($row-$start) % $step) -ne 0 -or
            $row -ge ($start + $count*$step)) {
            $state.Remove($key)
            $previous = $null
        }
    }
    if ($previous) {
        $row = [int]$previous.row
        $previousMachine = CellValue $sheet $machineCol $row
        if ((MachineId $previousMachine) -eq (MachineId $machine)) {
            if(TestScheduleRowMarked $sheet $machineCol $timeCol $row){$kept++;continue}
            if(-not $TimeOnly){SetScheduleBadges $sheet $machineCol $row $e $cfg}
            if (-not (CanRefreshTrackedTime $e $past $carryover ([bool]$cfg.update_policy.preserve_past) $previous $now)) {
                if (-not $TimeOnly -and $kind -eq 'color') { RepairColorIfBroken $sheet $cfg $e $row }
                $kept++
                continue
            }
            $newTime = EventTime $e
            if ((ExcelTimeText $sheet $timeCol $row) -ne ([TimeSpan]::Parse($newTime)).ToString('hh\:mm')) {
                PutValue $sheet $timeCol $row $newTime
                $changed++
            }
            if (-not $TimeOnly -and $kind -eq 'color') { RepairColorIfBroken $sheet $cfg $e $row }
            $state[$key] = @{row=$row; scheduled_at=[string]$e.scheduled_at}
            continue
        }
        # Recover the actual row instead of overwriting an operator's edit.
        $state.Remove($key)
    }
    # If this same machine/change is already in the Excel sheet, only refresh
    # the time cell. Don't create a duplicate when the local ID mapping was lost.
    $adopt=FindExistingSlot $sheet $cfg $e $state $book $start $step $count $ordered $now $when $carryover
    if($adopt -lt 0) {
        Write-Output "対応する行を特定できません: 種別=$kind / 設備=$machine / 予定=$id。変更しません。"
        $kept++
        continue
    }
    if($adopt -gt 0) {
        if(TestScheduleRowMarked $sheet $machineCol $timeCol $adopt){$kept++;continue}
        SetScheduleBadges $sheet $machineCol $adopt $e $cfg
        $oldTime=ExcelTimeText $sheet $timeCol $adopt
        $canUpdate=($e.completed -ne $true)
        if($canUpdate) {
            $newTime=EventTime $e
            if($oldTime -ne ([TimeSpan]::Parse($newTime)).ToString('hh\:mm')) {
                PutValue $sheet $timeCol $adopt $newTime
                $changed++
            }
            if(-not $TimeOnly -and $kind -eq 'color'){RepairColorIfBroken $sheet $cfg $e $adopt}
        }
        SetScheduleRowIdentity $state $key $prefix $adopt ([string]$e.scheduled_at)
        $kept++
        continue
    }
    # Full refresh can recover missing historical rows explicitly supplied by
    # the Worker; existing rows, uncertain matches and time-only stay protected.
    if(-not (CanAppendScheduleEvent $e $past ([bool]$TimeOnly))){$kept++;continue}
    $emptyRow = 0
    for ($i=[int]$nextSlot[$kind]; $i -lt $count; $i++) {
        $row = $start + ($i*$step)
        if ((CellValue $sheet $machineCol $row) -eq "" -and (CellValue $sheet $timeCol $row) -eq "") {
            $emptyRow = $row
            $nextSlot[$kind]=$i+1
            break
        }
    }
    if ($emptyRow -eq 0) {
        Write-Output "空き行がありません: 種別=$kind / 設備=$machine"
        continue
    }
    PutTextValue $sheet $machineCol $emptyRow $machine
    PutValue $sheet $timeCol $emptyRow (EventTime $e)
    SetScheduleBadges $sheet $machineCol $emptyRow $e $cfg
    if ($kind -eq "color") {
        if ($e.color_before) {
            $col=[string]$cfg.columns.color_before
            PutTextValue $sheet $col $emptyRow ([string]$e.color_before)
            if (PaintColorCell $sheet $col $emptyRow ([string]$e.color_before)) { $colored++ }
        }
        if ($e.color_after) {
            $col=[string]$cfg.columns.color_after
            PutTextValue $sheet $col $emptyRow ([string]$e.color_after)
            if (PaintColorCell $sheet $col $emptyRow ([string]$e.color_after)) { $colored++ }
        }
    }
    $state[$key] = @{row=$emptyRow; scheduled_at=[string]$e.scheduled_at}
    $added++
}
$dryingStats = UpdateDryingTable -sheet $sheet -cfg $cfg -book $book -state $state -now $now -timeOnly ([bool]$TimeOnly) -entries $script:DryingEntries -ResolvedShift $Shift
$added += [int]$dryingStats.added
$changed += [int]$dryingStats.updated
$kept += [int]$dryingStats.preserved
$shiftDate=[datetime]::ParseExact((ShiftDate $Shift $DateOffset $now),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
$visiblePending=@(GetVisiblePendingSchedule $sheet $cfg $state $book $start $step $count $shiftDate $Shift $now @($ordered))
MovePendingScheduleOrder $sheet $cfg $state $visiblePending
RemoveUnusedScheduleBadges $sheet ([string]$cfg.columns.mold_machine) $start $step $count
RemoveUnusedScheduleBadges $sheet ([string]$cfg.columns.color_machine) $start $step $count
$pruneCutoff = $now.AddDays(-14)
foreach ($stateKey in @($state.Keys)) {
    $entry = $state[$stateKey]
    if ($entry -and $entry.scheduled_at) {
        try {
            $entryTime = [datetime]::Parse([string]$entry.scheduled_at, [Globalization.CultureInfo]::InvariantCulture)
            if ($entryTime -lt $pruneCutoff) {
                $state.Remove($stateKey)
            }
        } catch {}
    }
}
# Save only our local row-ID state atomically. The open Excel workbook remains unsaved.
$json = ConvertTo-Json -InputObject $state -Depth 5
$tmpState = Join-Path $baseDir ('excel-daily-state-' + [guid]::NewGuid().ToString('N') + '.tmp')
try {
 [System.IO.File]::WriteAllText($tmpState, $json, (New-Object System.Text.UTF8Encoding($false)))
 [System.IO.File]::Copy($tmpState, $metadataPath, $true)
} finally {
 Remove-Item -LiteralPath $tmpState -Force -ErrorAction SilentlyContinue
}
if ($cfg.update_policy.save_automatically -eq $true) {
    $book.Save()
}
Remove-Item -LiteralPath $recoveryLock -ErrorAction Stop
$elapsed=[int]([DateTime]::UtcNow-$started).TotalMilliseconds
Write-Output "date_offset=$DateOffset; shift=$Shift; Workbook=$($book.Name); sheet=$($sheet.Name); added=$added; updated_times=$changed; colored_cells=$colored; preserved=$kept; drying_badges=$script:DryingBadgeCount; reordered=$script:ReorderedCount; elapsed_ms=$elapsed; excel_saved=$($cfg.update_policy.save_automatically)"
} finally {
 [ExcelDailyComMessageFilter]::Revoke()
 if($eventsChanged){try{$excel.EnableEvents=$oldEvents}catch{}}
 if($screenChanged){try{$excel.ScreenUpdating=$oldScreen}catch{}}
 $script:ExcelCells.Clear()
}
