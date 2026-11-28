param(
 [Parameter(Mandatory=$true)][string]$Settings,
 [Parameter(Mandatory=$true)][string]$Version,
 [long]$CommandId=0,
 [ValidateSet('','staged','failed')][string]$CommandStatus=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$OutputEncoding=[Console]::OutputEncoding
[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
function Test-AllowedApiUri([string]$raw){
 try{
  $u=[System.Uri]$raw
  if(-not $u.IsAbsoluteUri -or $u.Scheme -ne 'https' -or $u.AbsolutePath -ne '/api/excel-daily/events' -or $u.UserInfo -or $u.Query -or $u.Fragment){return $false}
  $h=$u.DnsSafeHost.ToLowerInvariant()
  if(!$h -or $h -eq 'localhost' -or $h.EndsWith('.localhost') -or $h.EndsWith('.local') -or $h.EndsWith('.internal') -or $h.Contains(':')){return $false}
  $ip=$null
  if([System.Net.IPAddress]::TryParse($h,[ref]$ip)){
   $b=$ip.GetAddressBytes()
   if($b.Length -ne 4 -or $b[0] -eq 0 -or $b[0] -eq 10 -or $b[0] -eq 127 -or ($b[0] -eq 169 -and $b[1] -eq 254) -or ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) -or ($b[0] -eq 192 -and $b[1] -eq 168) -or ($b[0] -eq 255 -and $b[1] -eq 255)){return $false}
  }
  return $true
 }catch{return $false}
}
function Resolve-ExcelDailyToken([string]$raw){
 if(!$raw){return ''}
 $t=$raw.Trim()
 if($t.StartsWith('dpapi:')){
  Add-Type -AssemblyName System.Security
  $enc=[Convert]::FromBase64String($t.Substring(6))
  $plain=[System.Security.Cryptography.ProtectedData]::Unprotect($enc,$null,[System.Security.Cryptography.DataProtectionScope]::CurrentUser)
  return [Text.Encoding]::UTF8.GetString($plain)
 }
 return $t
}
try {
 $cfg=Get-Content -LiteralPath $Settings -Raw -Encoding UTF8|ConvertFrom-Json
 $url=[string]$cfg.data_source.api_url
 $token=Resolve-ExcelDailyToken ([string]$cfg.data_source.api_token)
 if(!$token -or -not (Test-AllowedApiUri $url)){exit 0}
 $idPath=Join-Path (Split-Path -Parent $Settings) 'excel-daily-agent-id.txt'
 if(Test-Path -LiteralPath $idPath){$id=[IO.File]::ReadAllText($idPath).Trim()}
 else{$id=[guid]::NewGuid().ToString('N');[IO.File]::WriteAllText($idPath,$id,(New-Object Text.UTF8Encoding($false)))}
 if($id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid Excel Daily agent id'}
 $name=[string]$env:COMPUTERNAME
 if($cfg.agent -and $cfg.agent.display_name){$name=[string]$cfg.agent.display_name}
 if(!$name){$name='Excel Daily PC'}
 $name=$name.Trim()
 if($name.Length -gt 64){$name=$name.Substring(0,64)}
 $data=@{agent_id=$id;display_name=$name;version=$Version}
 if($CommandId -gt 0 -and $CommandStatus){
  $data.command_id=$CommandId;$data.command_status=$CommandStatus
 }
 $json=ConvertTo-Json -InputObject $data -Compress
 $headers=@{Authorization="Bearer $token"}
 $res=Invoke-RestMethod -Method Post -Uri ($url -replace '/events$','/agent/presence') -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 12 -MaximumRedirection 0
 $res|ConvertTo-Json -Compress -Depth 5
}catch{Write-Error ("Excel Daily presence failed: "+$_.Exception.Message);exit 1}
