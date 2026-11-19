# Excel Daily: ceil to ten-minute boundary, keep exact boundaries.
function RoundEventTimeUp([string]$time) {
 if($time -notmatch '^([01]?\d|2[0-3]):([0-5]\d)$'){throw "Invalid HH:mm time: $time"}
 $hour=[int]$Matches[1];$minute=[int]$Matches[2]
 $rounded=([int]([Math]::Ceiling(($hour*60+$minute)/10.0)*10))%1440
 return ('{0}:{1:D2}' -f ([int][Math]::Floor($rounded/60)),($rounded%60))
}
