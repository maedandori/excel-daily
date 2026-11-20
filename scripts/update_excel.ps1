param(
    [Parameter(Mandatory=$true)][string]$Settings,
    [Parameter(Mandatory=$true)][ValidateSet("day","night")][string]$Shift,
    [switch]$TimeOnly,
    [ValidateRange(0,1)][int]$DateOffset=0,
    [string]$WorkbookFullName=''
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
. (Join-Path $PSScriptRoot 'shift_date.ps1')
function FetchEvents($cfg, [string]$baseDir, [string]$shift,[int]$dateOffset=0) {
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
        return @($response.events)
    }
    $path = [string]$cfg.data_source.events_file
    if (-not $path) { throw "データ取得先を設定してください（events_fileまたはapi_url）。" }
    if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path $baseDir $path }
    if (-not (Test-Path -LiteralPath $path)) { throw "予定データが見つかりません: 「$path」。Excelは変更しません。" }
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $json.events) { throw "予定データにevents配列がありません。" }
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
    $targetTime=EventTime $event
    $prefix="$($book.FullName)|$($sheet.Name)|$kind|"
    $occupied=@{}
    foreach($entry in $state.GetEnumerator()) {
        if(([string]$entry.Key).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and $entry.Value.row) {
            $occupied[[int]$entry.Value.row]=$true
        }
    }
    $candidates=@()
    for($i=0;$i -lt $count;$i++) {
        $row=$start+$i*$step
        if($occupied.ContainsKey($row)) { continue }
        if((MachineId (CellValue $sheet $machineCol $row)) -ne (MachineId $machine)){continue}
        $time=ExcelTimeText $sheet $timeCol $row
        if(-not $time){continue}
        $colorMatched=$false
        if($kind -eq 'color') {
            $old=[string]$event.color_before
            $new=[string]$event.color_after
            $actualOld=CellValue $sheet ([string]$cfg.columns.color_before) $row
            $actualNew=CellValue $sheet ([string]$cfg.columns.color_after) $row
            if($old -and $new -and $actualOld -and $actualNew) {
                $colorMatched=($actualOld.Trim() -eq $old.Trim() -and $actualNew.Trim() -eq $new.Trim())
                if(-not $colorMatched){continue}
            }
        }
        $candidates+=@{row=$row;time=$time;colorMatched=$colorMatched}
    }
    if($candidates.Count -eq 0){return 0}
    $exact=@($candidates|Where-Object{$_.time -eq $targetTime})
    if($exact.Count -eq 1){return [int]$exact[0].row}
    if($exact.Count -gt 1){return -1}
    $sameColors=@($candidates|Where-Object{$_.colorMatched})
    if($sameColors.Count -eq 1 -and ($carryover -or (TimeStillPending $sameColors[0].time $scheduled $now))) {
        return [int]$sameColors[0].row
    }
    if($sameColors.Count -gt 1){return -1}
    $incoming=@($ordered|Where-Object{
        [string]$_.kind -eq $kind -and (MachineId ([string]$_.machine)) -eq (MachineId $machine)
    })
    if($candidates.Count -eq 1 -and $incoming.Count -eq 1 -and
       ($carryover -or (TimeStillPending $candidates[0].time $scheduled $now))) {
        return [int]$candidates[0].row
    }
    if($candidates.Count -gt 1){return -1}
    return 0
}
$cfg = Get-Content -LiteralPath $Settings -Raw -Encoding UTF8 | ConvertFrom-Json
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
$sheetName = if ($Shift -eq "day") { [string]$cfg.sheets.day } else { [string]$cfg.sheets.night }
$sheet = PickSheet $book $sheetName
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
    if ($TimeOnly -and -not $previous) { $kept++; continue }
    if ($previous) {
        $row = [int]$previous.row
        if ($row -lt $start -or (($row-$start) % $step) -ne 0 -or
            $row -ge ($start + $count*$step)) {
            if ($TimeOnly) { $kept++; continue }
            $state.Remove($key)
            $previous = $null
        }
    }
    if ($previous) {
        $row = [int]$previous.row
        $previousMachine = CellValue $sheet $machineCol $row
        if ($previousMachine -eq $machine) {
            if ($past -or ([bool]$cfg.update_policy.preserve_past -and (-not $carryover) -and
                $previous.scheduled_at -and ([DateTime]::Parse([string]$previous.scheduled_at) -le $now))) {
                if (-not $TimeOnly -and $kind -eq 'color') { RepairColorIfBroken $sheet $cfg $e $row }
                $kept++
                continue
            }
            $newTime = EventTime $e
            if ((CellValue $sheet $timeCol $row) -ne $newTime) {
                PutValue $sheet $timeCol $row $newTime
                $changed++
            }
            if (-not $TimeOnly -and $kind -eq 'color') { RepairColorIfBroken $sheet $cfg $e $row }
            $state[$key] = @{row=$row; scheduled_at=[string]$e.scheduled_at}
            continue
        }
        # The user edited/deleted a previously tracked row. Do not overwrite it.
        if ($TimeOnly) { $kept++; continue }
        $state.Remove($key)
    }
    if ($TimeOnly) { $kept++; continue }
    # If this same machine/change is already in the Excel sheet, only refresh
    # the time cell. Don't create a duplicate when the local ID mapping was lost.
    $adopt=FindExistingSlot $sheet $cfg $e $state $book $start $step $count $ordered $now $when $carryover
    if($adopt -lt 0) {
        Write-Output "対応する行を特定できません: 種別=$kind / 設備=$machine / 予定=$id。変更しません。"
        $kept++
        continue
    }
    if($adopt -gt 0) {
        if(-not $past) {
            $newTime=EventTime $e
            if((ExcelTimeText $sheet $timeCol $adopt) -ne $newTime) {
                PutValue $sheet $timeCol $adopt $newTime
                $changed++
            }
            if($kind -eq 'color'){RepairColorIfBroken $sheet $cfg $e $adopt}
        }
        $state[$key]=@{row=$adopt;scheduled_at=[string]$e.scheduled_at}
        $kept++
        continue
    }
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
# Save only our local row-ID state. The open Excel workbook remains unsaved.
$json = ConvertTo-Json -InputObject $state -Depth 5
[System.IO.File]::WriteAllText($metadataPath, $json, (New-Object System.Text.UTF8Encoding($false)))
if ($cfg.update_policy.save_automatically -eq $true) {
    $book.Save()
}
$elapsed=[int]([DateTime]::UtcNow-$started).TotalMilliseconds
Write-Output "date_offset=$DateOffset; Workbook=$($book.Name); sheet=$($sheet.Name); added=$added; updated_times=$changed; colored_cells=$colored; preserved=$kept; elapsed_ms=$elapsed; excel_saved=$($cfg.update_policy.save_automatically)"
} finally {
 if($eventsChanged){try{$excel.EnableEvents=$oldEvents}catch{}}
 if($screenChanged){try{$excel.ScreenUpdating=$oldScreen}catch{}}
 $script:ExcelCells.Clear()
}
