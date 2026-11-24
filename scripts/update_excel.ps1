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
function FetchEvents($cfg, [string]$baseDir, [string]$shift,[int]$dateOffset=0) {
    $script:DryingEntries=@()
    $url = [string]$cfg.data_source.api_url
    if ($url) {
        $headers = @{}
        if ($cfg.data_source.api_token) { $headers.Authorization = "Bearer $($cfg.data_source.api_token)" }
        # API must return {"events":[...]} with the documented fields.
        if (!$cfg.data_source.api_token) { throw '認証トークンが設定されていません。Excel Daily用トークンを設定してください。' }
        if ($url -notmatch '^https://') { throw 'APIの接続先にはHTTPSを使用してください。' }
        $date = ShiftDate $shift $dateOffset
        $separator = if ($url.Contains('?')) { '&' } else { '?' }
        $requestUrl = $url + $separator + 'shift=' + [uri]::EscapeDataString($shift) + '&date=' + [uri]::EscapeDataString($date)
        try {
                        # Decode original HTTP bytes, not Windows PowerShell 5.1's legacy text charset.
            $httpResponse = Invoke-WebRequest -UseBasicParsing -Uri $requestUrl -Headers $headers -Method Get -TimeoutSec 15
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
            PutValue $sheet $column $row $correct
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
function TestScheduleRowMarked($sheet,[string]$machineCol,[string]$timeCol,[int]$row){
 # Protect font strikethrough and horizontal manual line/ink crossing the number.
 foreach($column in @($machineCol,$timeCol)){
  $cell=ExcelCell $sheet $column $row
  if($cell.Font.Strikethrough -eq $true){return $true}
 }
 $cell=ExcelCell $sheet $machineCol $row
 if($cell.MergeCells){$cell=$cell.MergeArea}
 $cx=[double]$cell.Left+[double]$cell.Width/2
 $cy=[double]$cell.Top+[double]$cell.Height/2
 $tolerance=[double]$cell.Height*0.35
 foreach($shape in $sheet.Shapes){
  try{
   if([string]$shape.AlternativeText -like 'ExcelDaily managed * badge'){continue}
   $left=[double]$shape.Left;$top=[double]$shape.Top
   $w=[double]$shape.Width;$h=[double]$shape.Height
   if($w -ge [double]$cell.Width*0.5 -and $h -le [double]$cell.Height*0.8 -and
     $left -le $cx -and ($left+$w) -ge $cx -and
     $top -le ($cy+$tolerance) -and ($top+$h) -ge ($cy-$tolerance)){return $true}
  }catch{}
 }
 return $false
}
function GetVisiblePendingSchedule($sheet,$cfg,$state,$book,[int]$start,[int]$step,[int]$count,[datetime]$date,[string]$shift,[datetime]$now,[object[]]$events){
 foreach($kind in @('mold','color')){
  $machineCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_machine}else{[string]$cfg.columns.color_machine}
  $timeCol=if($kind -eq 'mold'){[string]$cfg.columns.mold_time}else{[string]$cfg.columns.color_time}
  $prefix="$($book.FullName)|$($sheet.Name)|$kind|"
  $fixed=@{}
  foreach($e in $events){
   if([string]$e.kind -ne $kind -or ($e.completed -ne $true -and $e.carryover -ne $true)){continue}
   $entry=$state[$prefix+[string]$e.id]
   if($entry -and $entry.row){$fixed[[int]$entry.row]=$true}
  }
  for($i=0;$i -lt $count;$i++){
   $row=$start+$i*$step
   if($fixed.ContainsKey($row)){continue}
   $machine=MachineId (CellValue $sheet $machineCol $row)
   if(!$machine){continue}
   $when=VisibleScheduleDate (ExcelTimeText $sheet $timeCol $row) $date $shift
   if($null -eq $when -or $when -le $now){continue}
   if(TestScheduleRowMarked $sheet $machineCol $timeCol $row){continue}
   [pscustomobject]@{kind=$kind;row=$row;key='';when=$when;scheduled_at=$when.ToString('s');prefix=$prefix}
  }
 }
}
# Plan a permutation of only the unfinished visible rows.
function PlanPendingScheduleOrder([object[]]$pending){
 foreach($kind in @('mold','color')){
  # Normalize hashtables to explicit objects before PowerShell 5.1 sorting.
  $items=@(foreach($p in $pending){
   if([string]$p.kind -eq $kind){
    [pscustomobject]@{kind=$kind;row=[int]$p.row;key=[string]$p.key;when=[datetime]$p.when;scheduled_at=[string]$p.scheduled_at}
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
# Move the existing managed Shapes, not unrelated worksheet objects.
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
    $entry=@{value=$cell.Value2}
    $range=$cell
    if($cell.MergeCells){$range=$cell.MergeArea}
    if($col -eq $machineCol){$entry.alignment=$range.HorizontalAlignment}
    if($colorColumns -contains $col){
     $entry.pattern=$range.Interior.Pattern
     $entry.background=$range.Interior.Color
     $entry.foreground=$range.Font.Color
    }
    $cells[$col]=$entry
   }
   if($blocked){break}
   $snap[$row]=$cells
   $machineCell=ExcelCell $sheet $machineCol $row
   if($machineCell.MergeCells){$machineCell=$machineCell.MergeArea}
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
  try{
   foreach($badge in $shapes){$badge.shape.Name='ExcelDaily_Tmp_'+[guid]::NewGuid().ToString('N')}
   foreach($move in $moving){
    $source=$snap[[int]$move.item.row]
    $row=[int]$move.destination
    foreach($col in $columns){
     $cell=ExcelCell $sheet $col $row
     $v=$source[$col]
     $cell.Value2=$v.value
     $range=$cell
     if($cell.MergeCells){$range=$cell.MergeArea}
     if($col -eq $machineCol){$range.HorizontalAlignment=$v.alignment}
     if($colorColumns -contains $col){
      $range.Interior.Color=$v.background
      $range.Interior.Pattern=$v.pattern
      $range.Font.Color=$v.foreground
     }
    }
   }
   foreach($badge in $shapes){
    $target=ExcelCell $sheet $machineCol $badge.destination
    if($target.MergeCells){$target=$target.MergeArea}
    $badge.shape.Left=[double]$target.Left+$badge.dx
    $badge.shape.Top=[double]$target.Top+$badge.dy
    $badge.shape.Name="ExcelDaily_$($badge.tag)Badge_${machineCol}_$($badge.destination)"
   }
  }catch{
   $failure=$_
   # Best-effort rollback to the post-time-update values and original badges.
   foreach($badge in $shapes){try{$badge.shape.Name='ExcelDaily_Tmp_'+[guid]::NewGuid().ToString('N')}catch{}}
   foreach($row in $snap.Keys){
    foreach($col in $columns){
     try{
      $cell=ExcelCell $sheet $col $row;$v=$snap[$row][$col];$cell.Value2=$v.value
      $range=$cell;if($cell.MergeCells){$range=$cell.MergeArea}
      if($col -eq $machineCol){$range.HorizontalAlignment=$v.alignment}
      if($colorColumns -contains $col){$range.Interior.Color=$v.background;$range.Interior.Pattern=$v.pattern;$range.Font.Color=$v.foreground}
     }catch{}
    }
   }
   foreach($badge in $shapes){try{$badge.shape.Left=$badge.left;$badge.shape.Top=$badge.top;$badge.shape.Name=$badge.name}catch{}}
   throw "予定の並べ替えに失敗しました。保存せずにシートを確認してください: $($failure.Exception.Message)"
  }
  foreach($move in $moving){
   $state[$move.item.key]=@{row=[int]$move.destination;scheduled_at=[string]$move.item.scheduled_at}
   $moves++
  }
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
$ordered = @($usable | Sort-Object scheduled_at)
$started=[DateTime]::UtcNow
$colored=0
$nextSlot=@{mold=0;color=0}
$script:DryingBadgeCount=0
$script:ReorderedCount=0
$pendingOrder=@()
# Existing state must identify a slot uniquely before it may move.
$rowOwners=@{}
foreach($mapping in $state.GetEnumerator()){
 foreach($k in @('mold','color')){
  $prefix="$($book.FullName)|$($sheet.Name)|$k|"
  if(([string]$mapping.Key).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and $mapping.Value.row){
   $slot="$k|$([int]$mapping.Value.row)"
   if(!$rowOwners.ContainsKey($slot)){$rowOwners[$slot]=0}
   $rowOwners[$slot]++
  }
 }
}
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
            if(-not $TimeOnly){SetScheduleBadges $sheet $machineCol $row $e $cfg}
            if ($past -or ([bool]$cfg.update_policy.preserve_past -and (-not $carryover) -and
                $previous.scheduled_at -and ([DateTime]::Parse([string]$previous.scheduled_at) -le $now))) {
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
            if($TimeOnly -and !$carryover -and $when -gt $now -and $rowOwners["$kind|$row"] -eq 1){
                $pendingOrder+=[pscustomobject]@{kind=$kind;row=$row;key=$key;when=$when;scheduled_at=[string]$e.scheduled_at}
            }
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
        SetScheduleBadges $sheet $machineCol $adopt $e $cfg
        $oldTime=ExcelTimeText $sheet $timeCol $adopt
        $canUpdate=(-not $past) -and ($carryover -or -not [bool]$cfg.update_policy.preserve_past -or (TimeStillPending $oldTime $when $now))
        if($canUpdate) {
            $newTime=EventTime $e
            if($oldTime -ne ([TimeSpan]::Parse($newTime)).ToString('hh\:mm')) {
                PutValue $sheet $timeCol $adopt $newTime
                $changed++
            }
            if(-not $TimeOnly -and $kind -eq 'color'){RepairColorIfBroken $sheet $cfg $e $adopt}
        }
        if($canUpdate -or $past){
            SetScheduleRowIdentity $state $key $prefix $adopt ([string]$e.scheduled_at)
        }
        if($TimeOnly -and $canUpdate -and !$carryover -and $when -gt $now){
            $pendingOrder+=[pscustomobject]@{kind=$kind;row=$adopt;key=$key;when=$when;scheduled_at=[string]$e.scheduled_at}
        }
        $kept++
        continue
    }
    # Time-only refresh can recover existing rows but never add a new one.
    if($TimeOnly){$kept++;continue}
    # Never append historical/completed work, except an explicitly active
    # handover item with status 100/410/440/490 targeting this shift.
    if ($past) { $kept++; continue }
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
    PutValue $sheet $machineCol $emptyRow $machine
    PutValue $sheet $timeCol $emptyRow (EventTime $e)
    SetScheduleBadges $sheet $machineCol $emptyRow $e $cfg
    if ($kind -eq "color") {
        if ($e.color_before) {
            $col=[string]$cfg.columns.color_before
            PutValue $sheet $col $emptyRow ([string]$e.color_before)
            if (PaintColorCell $sheet $col $emptyRow ([string]$e.color_before)) { $colored++ }
        }
        if ($e.color_after) {
            $col=[string]$cfg.columns.color_after
            PutValue $sheet $col $emptyRow ([string]$e.color_after)
            if (PaintColorCell $sheet $col $emptyRow ([string]$e.color_after)) { $colored++ }
        }
    }
    $state[$key] = @{row=$emptyRow; scheduled_at=[string]$e.scheduled_at}
    $added++
}
if($TimeOnly){MovePendingScheduleOrder $sheet $cfg $state @($pendingOrder)}
RemoveUnusedScheduleBadges $sheet ([string]$cfg.columns.mold_machine) $start $step $count
RemoveUnusedScheduleBadges $sheet ([string]$cfg.columns.color_machine) $start $step $count
# Save only our local row-ID state. The open Excel workbook remains unsaved.
$json = ConvertTo-Json -InputObject $state -Depth 5
[System.IO.File]::WriteAllText($metadataPath, $json, (New-Object System.Text.UTF8Encoding($false)))
if ($cfg.update_policy.save_automatically -eq $true) {
    $book.Save()
}
$elapsed=[int]([DateTime]::UtcNow-$started).TotalMilliseconds
Write-Output "date_offset=$DateOffset; shift=$Shift; Workbook=$($book.Name); sheet=$($sheet.Name); added=$added; updated_times=$changed; colored_cells=$colored; preserved=$kept; drying_badges=$script:DryingBadgeCount; reordered=$script:ReorderedCount; elapsed_ms=$elapsed; excel_saved=$($cfg.update_policy.save_automatically)"
} finally {
 if($eventsChanged){try{$excel.EnableEvents=$oldEvents}catch{}}
 if($screenChanged){try{$excel.ScreenUpdating=$oldScreen}catch{}}
 $script:ExcelCells.Clear()
}
