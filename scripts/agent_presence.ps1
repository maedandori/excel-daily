param(
 [Parameter(Mandatory=$true)][string]$Settings,
 [Parameter(Mandatory=$true)][string]$Version,
 [long]$CommandId=0,
 [ValidateSet('','staged','failed')][string]$CommandStatus=''
)
$ErrorActionPreference='Stop'
try {
 $cfg=Get-Content -LiteralPath $Settings -Raw -Encoding UTF8|ConvertFrom-Json
 $url=[string]$cfg.data_source.api_url
 $token=[string]$cfg.data_source.api_token
 if(!$token -or $url -notmatch '^https://[^/]+/api/excel-daily/events$'){exit 0}
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
 $res=Invoke-RestMethod -Method Post -Uri ($url -replace '/events$','/agent/presence') -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 12
 $res|ConvertTo-Json -Compress -Depth 5
}catch{Write-Error ("Excel Daily presence failed: "+$_.Exception.Message);exit 1}
