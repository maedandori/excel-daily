# Preserves manual entries and finds the Japanese drying table by its headers.
# Excel cell helpers, MachineId, EventTime and ExcelTimeText are defined by update_excel.ps1.
function DryingTimeFromConfig($cfg,[object[]]$materials) {
 $rules=$null
 if($cfg.drying -and $cfg.drying.time_by_material){$rules=$cfg.drying.time_by_material}
 $values=@()
 foreach($material in @($materials)){
  $name=[string]$material
  if(!$rules -or !$rules.PSObject.Properties[$name]){return '要確認'}
  $value=[string]$rules.PSObject.Properties[$name].Value
  if([string]::IsNullOrWhiteSpace($value)){return '要確認'}
  $values+= $value.Trim()
 }
 $unique=@($values|Select-Object -Unique)
 if($unique.Count -ne 1){return '要確認'}
 return [string]$unique[0]
}
function DryingRowSpan($cell){
 if($cell.MergeCells){return [Math]::Max(1,[int]$cell.MergeArea.Rows.Count)}
 return 1
}
function FindDryingRows($sheet,$cfg){
 $c=$cfg.drying.columns
 $m=if($c -and $c.machine){[string]$c.machine}else{'AY'}
 $s=if($c -and $c.production_start){[string]$c.production_start}else{'AZ'}
 $d=if($c -and $c.drying_time){[string]$c.drying_time}else{'BA'}
 $title=if($cfg.drying -and $cfg.drying.header){[string]$cfg.drying.header}else{'乾燥原料必要'}
 $max=5
 if($cfg.drying -and $cfg.drying.max_entries){$max=[int]$cfg.drying.max_entries}
 if($max -lt 1 -or $max -gt 20){throw '乾燥原料の最大件数は1～20に設定してください。'}
 $range=$sheet.Range(($m+'1:'+$d+'250'))
 $found=$range.Find($title,[Type]::Missing,-4163,2)
 if($null -eq $found){return $null}
 $headerRow=[int]$found.Row+(DryingRowSpan $found)
 if((CellValue $sheet $m $headerRow) -notmatch '号機' -or
    (CellValue $sheet $s $headerRow) -notmatch '量産開始' -or
    (CellValue $sheet $d $headerRow) -notmatch '乾燥時間'){return $null}
 $row=$headerRow+(DryingRowSpan (ExcelCell $sheet $m $headerRow))
 $rows=@()
 for($i=0;$i -lt $max;$i++){
  if($row -gt 250){break}
  $rows+=$row
  $row+=(DryingRowSpan (ExcelCell $sheet $m $row))
 }
 return @{rows=@($rows);machine=$m;start=$s;duration=$d}
}
function UpdateDryingTable {
 param($sheet,$cfg,$book,$state,[datetime]$now,[bool]$timeOnly,[object[]]$entries)
 $stats=@{added=0;updated=0;preserved=0}
 if($cfg.drying -and $cfg.drying.enabled -eq $false){return $stats}
 if(!$entries -or $entries.Count -eq 0){return $stats}
 $usable=@()
 foreach($e in $entries){
  if([string]$e.shift -ne [string]$Shift){continue}
  if(!$e.id -or !$e.machine -or !$e.scheduled_at -or !$e.time -or
     !$e.materials -or @($e.materials).Count -eq 0){
   throw '乾燥予定のデータ形式が正しくありません。Excelは変更しません。'
  }
  $scheduled=[datetime]::Parse([string]$e.scheduled_at,[Globalization.CultureInfo]::InvariantCulture)
  $null=EventTime $e
  $usable+=@{data=$e;scheduled=$scheduled}
 }
 if(!$usable.Count){return $stats}
 $layout=FindDryingRows $sheet $cfg
 if($null -eq $layout){
  Write-Output '乾燥原料必要の表を確認できないため、乾燥予定は変更しません。'
  return $stats
 }
 $slots=@($layout.rows)
 $machineCol=$layout.machine;$startCol=$layout.start;$durationCol=$layout.duration
 $prefix="$($book.FullName)|$($sheet.Name)|drying|"
 $occupied=@{}
 foreach($mapping in $state.GetEnumerator()){
  if(([string]$mapping.Key).StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -and
      $mapping.Value.row){$occupied[[int]$mapping.Value.row]=[string]$mapping.Key}
 }
 foreach($item in @($usable|Sort-Object {$_.scheduled})){
  $e=$item.data;$scheduled=$item.scheduled
  $key=$prefix+[string]$e.id
  $machine=[string]$e.machine
  $targetTime=EventTime $e
  $carryover=($e.carryover -eq $true) -and ($e.completed -ne $true) -and
     (@('100','410','440','490') -contains [string]$e.status_code)
  $past=($scheduled -le $now -and !$carryover) -or $e.completed -eq $true
  $previous=$state[$key]
  if($previous -and $slots -notcontains [int]$previous.row){
   if($timeOnly){$stats.preserved++;continue}
   $state.Remove($key);$previous=$null
  }
  if($previous){
   $row=[int]$previous.row
   if((MachineId (CellValue $sheet $machineCol $row)) -ne (MachineId $machine)){
    if($timeOnly){$stats.preserved++;continue}
    $state.Remove($key);$previous=$null
   }
  }
  if($previous){
   $row=[int]$previous.row
   if($past -or ($cfg.update_policy.preserve_past -eq $true -and !$carryover -and
      $previous.scheduled_at -and
      [datetime]::Parse([string]$previous.scheduled_at,[Globalization.CultureInfo]::InvariantCulture) -le $now)){
    $stats.preserved++;continue
   }
   if((ExcelTimeText $sheet $startCol $row) -ne $targetTime){
    PutValue $sheet $startCol $row $targetTime
    $stats.updated++
   }
   $state[$key]=@{row=$row;scheduled_at=[string]$e.scheduled_at}
   continue
  }
  if($timeOnly -or $past){$stats.preserved++;continue}
  $same=@()
  foreach($r in $slots){
   if($occupied.ContainsKey([int]$r)){continue}
   if((MachineId (CellValue $sheet $machineCol $r)) -eq (MachineId $machine) -and
      (ExcelTimeText $sheet $startCol $r) -eq $targetTime){$same+=[int]$r}
  }
  if($same.Count -eq 1){
   $row=[int]$same[0]
   $occupied[$row]=$key
   $state[$key]=@{row=$row;scheduled_at=[string]$e.scheduled_at}
   $stats.preserved++;continue
  }
  if($same.Count -gt 1){
   Write-Output "乾燥予定の行が重複しています: $machine 号機 / $targetTime"
   $stats.preserved++;continue
  }
  $row=0
  foreach($r in $slots){
   if($occupied.ContainsKey([int]$r)){continue}
   if((CellValue $sheet $machineCol $r) -eq '' -and
      (CellValue $sheet $startCol $r) -eq '' -and
      (CellValue $sheet $durationCol $r) -eq ''){
    $row=[int]$r;break
   }
  }
  if(!$row){
   Write-Output "乾燥原料の空き行がありません: $machine 号機"
   $stats.preserved++;continue
  }
  PutValue $sheet $machineCol $row $machine
  PutValue $sheet $startCol $row $targetTime
  # Resin drying duration is grade-specific: never fabricate hours.
  PutValue $sheet $durationCol $row (DryingTimeFromConfig $cfg @($e.materials))
  $occupied[$row]=$key
  $state[$key]=@{row=$row;scheduled_at=[string]$e.scheduled_at}
  $stats.added++
 }
 return $stats
}

# Join the two feeds by the exact encoded machine/order/sequence/tool/mold identity.
# A machine number or rounded hour alone is never sufficient.
function GetDryingMaterialsForEvent($event,[object[]]$entries){
 $parts=([string]$event.id).Split(':')
 if($parts.Count -ne 6 -or $parts[1] -notin @('mold','color')){return}
 foreach($entry in $entries){
  if([string]$entry.shift -ne [string]$event.shift -or [string]$entry.machine -ne [string]$event.machine){continue}
  $d=([string]$entry.id).Split(':')
  if($d.Count -ne 7 -or $d[1] -cne 'drying'){continue}
  if($parts[0] -cne $d[0] -or $parts[2] -cne $d[2] -or $parts[3] -cne $d[3] -or
     $parts[4] -cne $d[5] -or $parts[5] -cne $d[6]){continue}
  if([string]$entry.scheduled_at -cne [string]$event.scheduled_at){continue}
  foreach($material in @($entry.materials)){
   if([string]$material -in @('ABS','PS-TH21','PC','ナイロン','AS')){[string]$material}
  }
 }
}
