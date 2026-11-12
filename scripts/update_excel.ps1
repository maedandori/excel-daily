param(
    [Parameter(Mandatory=$true)][string]$Settings,
    [Parameter(Mandatory=$true)][ValidateSet("day","night")][string]$Shift
)
$ErrorActionPreference = "Stop"
# PowerShell 5.1 otherwise writes stdout/stderr using the OEM code page,
# which corrupts Japanese file paths when Rust reads them as UTF-8.
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

function CellValue($sheet, [string]$column, [int]$row) {
    $v = $sheet.Range("$column$row").Value2
    if ($null -eq $v) { return "" }
    return [string]$v
}
function PutValue($sheet, [string]$column, [int]$row, $value) {
    # Merged cells have the writable value in the top-left cell.
    $target = $sheet.Range("$column$row")
    if ($target.MergeCells) { $target = $target.MergeArea.Cells.Item(1,1) }
    $target.Value2 = $value
}
function NormalizeName([string]$name) { return $name.Trim().Replace([char]0x3000, ' ') }
function PickWorkbook($excel, $config) {
    $pat = [string]$config.workbook.filename_pattern
    if (-not $pat) { $pat = '成形スケジュール*.xlsx' }
    $matches = @()
    foreach ($book in $excel.Workbooks) {
        if ($book.Name -like $pat) { $matches += $book }
    }
    if ($matches.Count -eq 0) { throw "Không có workbook đang mở khớp $pat. Chỉ cập nhật Excel đang mở." }
    $selected = [string]$config.workbook.selected_workbook
    if ($selected) {
        $found = @($matches | Where-Object { $_.Name -eq $selected })
        if ($found.Count -eq 1) { return $found[0] }
        throw "Workbook đã chọn '$selected' chưa mở. Không ghi sang workbook khác."
    }
    if ($matches.Count -eq 1) { return $matches[0] }
    if ($config.workbook.choose_if_multiple -eq $false) {
        throw "Có nhiều workbook phù hợp; hãy đặt selected_workbook trong settings.json."
    }
    $lines = for ($i=0; $i -lt $matches.Count; $i++) {
        "$($i+1). $($matches[$i].Name)  [$($matches[$i].FullName)]"
    }
    $message = "Chọn file Excel đang mở để cập nhật (nhập số):" + [Environment]::NewLine +
        ($lines -join [Environment]::NewLine)
    $answer = [Microsoft.VisualBasic.Interaction]::InputBox($message, "Excel Daily - chọn workbook", "1")
    $choice = 0
    if (-not [int]::TryParse($answer, [ref]$choice) -or $choice -lt 1 -or $choice -gt $matches.Count) {
        throw "Đã hủy chọn workbook; không thay đổi Excel."
    }
    return $matches[$choice-1]
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
    throw "Không tìm thấy sheet '$wanted'. Kiểm tra tên sheet trong settings.json."
}
function ShiftDate([string]$shift) {
    $now=[DateTime]::Now
    # Night shift from 19:00 to 07:30 belongs to the day it started.
    if ($shift -eq 'night' -and $now.TimeOfDay -lt [TimeSpan]::FromHours(7.5)) {
        return $now.AddDays(-1).ToString('yyyy-MM-dd')
    }
    return $now.ToString('yyyy-MM-dd')
}
function FetchEvents($cfg, [string]$baseDir, [string]$shift) {
    $url = [string]$cfg.data_source.api_url
    if ($url) {
        $headers = @{}
        if ($cfg.data_source.api_token) { $headers.Authorization = "Bearer $($cfg.data_source.api_token)" }
        # API must return {"events":[...]} with the documented fields.
        if (!$cfg.data_source.api_token) { throw 'Chưa cấu hình data_source.api_token. Cần token Excel Daily riêng trên Cloudflare.' }
        if ($url -notmatch '^https://') { throw 'API phải sử dụng HTTPS.' }
        $date = ShiftDate $shift
        $separator = if ($url.Contains('?')) { '&' } else { '?' }
        $requestUrl = $url + $separator + 'shift=' + [uri]::EscapeDataString($shift) + '&date=' + [uri]::EscapeDataString($date)
        try {
            $response = Invoke-RestMethod -Uri $requestUrl -Headers $headers -Method Get -TimeoutSec 15
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            if ($statusCode -eq 401) { throw 'API 401 Unauthorized: Excel Daily token không khớp EXCEL_DAILY_TOKEN của Worker.' }
            if ($statusCode -eq 503) { throw 'API 503: Worker chưa cấu hình EXCEL_DAILY_TOKEN hoặc D1 chưa sẵn sàng.' }
            throw
        }
        if ($null -eq $response.events) { throw "API không có trường events; không sửa Excel." }
        return @($response.events)
    }
    $path = [string]$cfg.data_source.events_file
    if (-not $path) { throw "Cần khai báo data_source.events_file hoặc api_url." }
    if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path $baseDir $path }
    if (-not (Test-Path -LiteralPath $path)) { throw "Không thấy nguồn lịch '$path'; không sửa Excel." }
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $json.events) { throw "events_file cần có danh sách events." }
    return @($json.events)
}
function EventTime($event) {
    $t = [string]$event.time
    if ($t -notmatch '^\d{1,2}:\d{2}$') { throw "Giờ không hợp lệ ($($event.id)): $t" }
    return $t
}
$cfg = Get-Content -LiteralPath $Settings -Raw -Encoding UTF8 | ConvertFrom-Json
$baseDir = Split-Path -Parent $Settings
$events = @(FetchEvents $cfg $baseDir $Shift)
# Validate entire input before touching Excel.
$usable = @()
foreach ($e in $events) {
    if ([string]$e.shift -ne $Shift) { continue }
    # A carried-over event must explicitly target this shift. It retains its
    # stable event ID and original timestamp, despite the previous shift date.
    if ([string]$e.kind -notin @("mold","color")) { throw "Loại lịch không hợp lệ: $($e.kind)" }
    if (-not $e.id -or -not $e.machine -or -not $e.scheduled_at) {
        throw "Mỗi event cần id, machine, scheduled_at."
    }
    $null = EventTime $e
    $usable += $e
}
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
if ($null -eq $excel) { throw "Excel chưa được mở." }
$book = PickWorkbook $excel $cfg
$sheetName = if ($Shift -eq "day") { [string]$cfg.sheets.day } else { [string]$cfg.sheets.night }
$sheet = PickSheet $book $sheetName
$start = if ($Shift -eq "day") { [int]$cfg.rows.day_start } else { [int]$cfg.rows.night_start }
$step = [int]$cfg.rows.step
$count = [int]$cfg.rows.max_entries
if ($step -lt 1 -or $count -lt 1 -or $count -gt 100) { throw "Cấu hình rows không hợp lệ." }
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
    if ($previous) {
        $row = [int]$previous.row
        $previousMachine = CellValue $sheet $machineCol $row
        if ($previousMachine -eq $machine) {
            if ($past -or ([bool]$cfg.update_policy.preserve_past -and (-not $carryover) -and
                $previous.scheduled_at -and ([DateTime]::Parse([string]$previous.scheduled_at) -le $now))) {
                $kept++
                continue
            }
            $newTime = EventTime $e
            if ((CellValue $sheet $timeCol $row) -ne $newTime) {
                PutValue $sheet $timeCol $row $newTime
                $changed++
            }
            $state[$key] = @{row=$row; scheduled_at=[string]$e.scheduled_at}
            continue
        }
        # The user edited/deleted a previously tracked row. Do not overwrite it.
        $state.Remove($key)
    }
    # Never append historical/completed work, except an explicitly active
    # handover item with status 100/410/440/490 targeting this shift.
    if ($past) { $kept++; continue }
    $emptyRow = 0
    for ($i=0; $i -lt $count; $i++) {
        $row = $start + ($i*$step)
        if ((CellValue $sheet $machineCol $row) -eq "" -and (CellValue $sheet $timeCol $row) -eq "") {
            $emptyRow = $row
            break
        }
    }
    if ($emptyRow -eq 0) {
        Write-Output "Không có dòng trống: $kind / máy $machine"
        continue
    }
    PutValue $sheet $machineCol $emptyRow $machine
    PutValue $sheet $timeCol $emptyRow (EventTime $e)
    if ($kind -eq "color") {
        if ($e.color_before) { PutValue $sheet ([string]$cfg.columns.color_before) $emptyRow ([string]$e.color_before) }
        if ($e.color_after) { PutValue $sheet ([string]$cfg.columns.color_after) $emptyRow ([string]$e.color_after) }
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
Write-Output "Workbook=$($book.Name); sheet=$($sheet.Name); added=$added; updated_times=$changed; preserved=$kept; excel_saved=$($cfg.update_policy.save_automatically)"
