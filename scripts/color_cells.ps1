# Excel uses VBA/OLE RGB order: red + green*256 + blue*65536.
function ExcelDailyColorStyle([string]$label){
 if([string]::IsNullOrWhiteSpace($label)){return $null}
 $n=$label.Normalize([Text.NormalizationForm]::FormKC).ToUpperInvariant() -replace '[\s　・･]',''
 $bg=$null;$fg=0x222222
 if($n -match 'シールズBK|TH-BK|ブラック|BLACK|黒'){$bg=0x252525;$fg=0xFFFFFF}
 elseif($n -match 'ダークブラウン|ショコラブラウン|DARKBROWN'){$bg=0x394C70;$fg=0xFFFFFF}
 elseif($n -match 'ブラウン|BROWN|茶色'){$bg=0x4576A8;$fg=0xFFFFFF}
 elseif($n -match 'チャコール|CHARCOAL'){$bg=0x666666;$fg=0xFFFFFF}
 elseif($n -match 'グレー|GRAY|GREY|GY$'){$bg=0xB5B5B5;$fg=0x101010}
 elseif($n -match 'ホワイト|イオンW|Wイオン|メッシュW|WHITE|^\(W\)$'){$bg=0xFFFFFF}
 elseif($n -match 'クリア|CLEAR|透明'){$bg=0xFFF0DF}
 elseif($n -match 'イエロー|YELLOW|黄色'){$bg=0x7BEBFF}
 elseif($n -match 'ベージュ|BEIGE|STCBE'){$bg=0xBFE4F4}
 elseif($n -match 'ブルー|BLUE|青色'){$bg=0xE5A34D;$fg=0xFFFFFF}
 elseif($n -match 'グリーン|GREEN|緑色'){$bg=0x78B373}
 elseif($n -match 'レッド|RED|赤色'){$bg=0x4646E8;$fg=0xFFFFFF}
 if($null -eq $bg){return $null}
 return [pscustomobject]@{Background=[int]$bg;Foreground=[int]$fg}
}
