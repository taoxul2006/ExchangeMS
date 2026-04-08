param(
  [string]$BaseDir = "",
  [int]$Port = 3080
)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false)
$BaseDir = if([string]::IsNullOrWhiteSpace($BaseDir)) { $PSScriptRoot } else { $BaseDir }
$script:BaseDir=$BaseDir
$script:Web=Join-Path $BaseDir 'web-ui'
$script:UsersFile=Join-Path $BaseDir 'leaverlist.ps1'
$script:ConfigFile=Join-Path $BaseDir 'exchange-web-config.json'
$script:LogFile=Join-Path $PSScriptRoot 'exchange-web-server.log'
$script:Session=$null
$script:Verified=$null

function ReadText([string]$p){ if(Test-Path -LiteralPath $p){ [IO.File]::ReadAllText($p,[Text.Encoding]::UTF8) } else { '' } }
function Defaults {
  if(Test-Path -LiteralPath $script:ConfigFile){
    $config = Get-Content -Raw $script:ConfigFile | ConvertFrom-Json
    return [ordered]@{
      username=if($config.username){[string]$config.username}else{''}
      connectionUri=if($config.connectionUri){[string]$config.connectionUri}else{'http://ExchangeFQDN/PowerShell/'}
      authentication=if($config.authentication){[string]$config.authentication}else{'Kerberos'}
      targetFolder=if($config.targetFolder){[string]$config.targetFolder}else{'Deleted Items'}
    }
  }

  [ordered]@{
    username=''
    connectionUri='http://ExchangeFQDN/PowerShell/'
    authentication='Kerberos'
    targetFolder='Deleted Items'
  }
}
function NormUsers([string]$t){ @((($t -split "\r?\n")|%{$_.Trim()}|?{$_}|%{$_ -replace '@brbiotech\.com$',''})) }
function LoadUsers { (ReadText $script:UsersFile).Replace("`r`n","`n").Trim() }
function SaveUsers([string]$t){ $l=NormUsers $t; [IO.File]::WriteAllText($script:UsersFile,($(if($l.Count){($l -join "`r`n")+"`r`n"}else{''})),[Text.UTF8Encoding]::new($false)); $l }
function JsonBytes($o){ [Text.Encoding]::UTF8.GetBytes(($o|ConvertTo-Json -Depth 8 -Compress)) }
function StatusText([int]$c){ switch($c){200{'OK'}400{'Bad Request'}404{'Not Found'}default{'OK'}} }
function LogLine([string]$m){ Add-Content -LiteralPath $script:LogFile -Value ('[{0}] {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $m) -Encoding UTF8 }
function Resp($s,[int]$c,[string]$ct,[byte[]]$b){ if($null -eq $b){$b=[byte[]]::new(0)}; $h=@("HTTP/1.1 $c $(StatusText $c)","Content-Type: $ct","Content-Length: $($b.Length)",'Cache-Control: no-store','Connection: close','','') -join "`r`n"; $hb=[Text.Encoding]::ASCII.GetBytes($h); $s.Write($hb,0,$hb.Length); if($b.Length){$s.Write($b,0,$b.Length)}; $s.Flush() }
function JResp($s,[int]$c,$o){ Resp $s $c 'application/json; charset=utf-8' (JsonBytes $o) }
function Req($s){
  $buf=New-Object byte[] 4096; $all=New-Object 'System.Collections.Generic.List[byte]'; $end=-1
  while($end -lt 0){ $r=$s.Read($buf,0,$buf.Length); if($r -le 0){break}; for($i=0;$i -lt $r;$i++){$all.Add($buf[$i])}; $arr=$all.ToArray(); for($i=[Math]::Max(0,$arr.Length-$r-4);$i -le $arr.Length-4;$i++){ if($arr[$i]-eq 13 -and $arr[$i+1]-eq 10 -and $arr[$i+2]-eq 13 -and $arr[$i+3]-eq 10){$end=$i+4; break} } }
  if($end -lt 0){ throw 'Invalid HTTP request.' }
  $arr=$all.ToArray(); $head=[Text.Encoding]::ASCII.GetString($arr,0,$end); $ls=$head -split "`r`n"; $rp=$ls[0].Split(' '); if($rp.Count -lt 2){ throw 'Invalid request line.' }
  $hs=@{}; if($ls.Count -gt 1){ foreach($line in $ls[1..($ls.Count-1)]){ if(-not $line){continue}; $x=$line.IndexOf(':'); if($x -lt 0){continue}; $hs[$line.Substring(0,$x).Trim().ToLowerInvariant()]=$line.Substring($x+1).Trim() } }
  $len=0; if($hs.ContainsKey('content-length')){ $len=[int]$hs['content-length'] }
  $body=New-Object byte[] $len; $already=$arr.Length-$end; if($already -gt 0){ [Array]::Copy($arr,$end,$body,0,[Math]::Min($already,$len)) }
  $off=[Math]::Min($already,$len); while($off -lt $len){ $r=$s.Read($body,$off,$len-$off); if($r -le 0){break}; $off+=$r }
  $raw=$rp[1]; [pscustomobject]@{ Method=$rp[0].ToUpperInvariant(); Path=[Uri]::UnescapeDataString((($raw -split '\?')[0])); Body=$(if($off){[Text.Encoding]::UTF8.GetString($body,0,$off)}else{''}) }
}
function Json([string]$b){ if([string]::IsNullOrWhiteSpace($b)){ [pscustomobject]@{} } else { $b=$b.Replace([string][char]0,'').Trim().Trim([char]0xFEFF); try{ $b|ConvertFrom-Json } catch{ throw ('Request body is not valid JSON: ' + $b) } } }
function Finger($o){ ([ordered]@{startTime=[string]$o.startTime;endTime=[string]$o.endTime;from=[string]$o.from;subject=[string]$o.subject;keyword=[string]$o.keyword;users=(NormUsers ([string]$o.users))}|ConvertTo-Json -Depth 5 -Compress) }
function NeedConn { if($null -eq $script:Session){ throw 'Please connect to Exchange first.' } }
function NeedInput($b){ $u=NormUsers ([string]$b.users); if(-not $u.Count){ throw 'At least one target account is required.' }; if([string]::IsNullOrWhiteSpace([string]$b.startTime) -or [string]::IsNullOrWhiteSpace([string]$b.endTime)){ throw 'Start time and end time are required.' }; $u }
function NeedDelete($b){ if((@([string]$b.from,[string]$b.subject,[string]$b.keyword)|?{ -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0){ throw 'Delete requires from, subject, or keyword.' }; if(-not [bool]$b.reviewConfirmed){ throw 'Please confirm that the search results have been reviewed.' }; if($null -eq $script:Verified){ throw 'Run a matching search before delete.' }; if((Finger $b) -ne $script:Verified.fingerprint){ throw 'Delete filters do not match the last verified search.' } }
function Cred($c){ $sp=ConvertTo-SecureString -String ([string]$c.password) -AsPlainText -Force; New-Object pscredential ([string]$c.username),$sp }
function OpenEx($c){ $a=[string]$c.authentication; if([string]::IsNullOrWhiteSpace($a)){$a='Kerberos'}; $ses=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ([string]$c.connectionUri) -Authentication $a -Credential (Cred $c) -ErrorAction Stop; Import-PSSession $ses -DisableNameChecking -AllowClobber|Out-Null; $ses }
function Dt([string]$v){ if([string]::IsNullOrWhiteSpace($v)){$null}else{ [datetime]::Parse($v) } }
function UQ([datetime]$d){ if($null -eq $d){''}else{ $u=$d.ToUniversalTime(); "{0}/{1}/{2} {3}:{4}:{5}" -f $u.Month,$u.Day,$u.Year,$u.Hour,$u.Minute,$u.Second } }
function Esc([string]$v){ if([string]::IsNullOrWhiteSpace($v)){''}else{ $v.Replace('"','`"') } }
function Query($f){ $p=@(); $st=Dt ([string]$f.startTime); $et=Dt ([string]$f.endTime); if($null -ne $st -or $null -ne $et){ $p+=('Sent:"{0} ... {1}"' -f (UQ $st),(UQ $et)) }; if(-not [string]::IsNullOrWhiteSpace([string]$f.from)){ $p+=('From:"{0}"' -f (Esc ([string]$f.from))) }; if(-not [string]::IsNullOrWhiteSpace([string]$f.subject)){ $p+=('Subject:"{0}"' -f (Esc ([string]$f.subject))) }; if(-not [string]::IsNullOrWhiteSpace([string]$f.keyword)){ $p+=('"{0}"' -f (Esc ([string]$f.keyword))) }; if(-not $p.Count){ throw 'At least one search filter is required.' }; $p -join ' AND ' }
function Count($r){ if($null -eq $r){0}else{ $x=[string](($r|select -First 1).ResultItemsCount); if([string]::IsNullOrWhiteSpace($x)){0}else{ $d=$x -replace '[^\d]',''; if([string]::IsNullOrWhiteSpace($d)){0}else{ [int]$d } } } }
function DoConnect($c){ $ses=$null; try{ $ses=OpenEx $c; $cmd=Get-Command Search-Mailbox -ErrorAction Stop; [ordered]@{ ok=$true; message='Exchange connection succeeded and Search-Mailbox is available.'; server=$ses.ComputerName; commandName=$cmd.Name } } finally { if($ses){ Remove-PSSession -Session $ses -ErrorAction SilentlyContinue } } }
function DoAction([string]$mode,$c,$f,[string[]]$users,[string]$tm,[string]$tf){
  $q=Query $f; $ses=$null
  try{
    $ses=OpenEx $c; $rows=@()
    foreach($u in $users){
      $row=[ordered]@{ user=$u; estimatedCount=0; actionTaken=$false; status='success'; message='' }
      try{
        $est=Search-Mailbox -SearchQuery $q -Identity $u -EstimateResultOnly -ErrorAction Stop
        $n=Count $est; $row.estimatedCount=$n
        if($mode -eq 'search' -and $n -gt 0){ Search-Mailbox -SearchQuery $q -Identity $u -TargetMailbox $tm -TargetFolder $tf -ErrorAction Stop|Out-Null; $row.actionTaken=$true; $row.message='Copied to target mailbox.' }
        if($mode -eq 'delete' -and $n -gt 0){ Search-Mailbox -SearchQuery $q -Identity $u -DeleteContent -Force -ErrorAction Stop|Out-Null; $row.actionTaken=$true; $row.message='Delete executed.' }
        if($n -eq 0){ $row.message='No matching mail found.' }
      } catch { $row.status='error'; $row.message=$_.Exception.Message }
      $rows+=[pscustomobject]$row
    }
    $ok=@($rows|?{$_.status -eq 'success'}).Count
    [ordered]@{ ok=$true; message=$(if($mode -eq 'search'){"Search completed. Accounts processed: $($users.Count). Success: $ok."}else{"Delete completed. Accounts processed: $($users.Count). Success: $ok."}); query=$q; results=$rows }
  } finally { if($ses){ Remove-PSSession -Session $ses -ErrorAction SilentlyContinue } }
}
function CType([string]$p){ switch(([IO.Path]::GetExtension($p)).ToLowerInvariant()){ '.html'{'text/html; charset=utf-8'} '.css'{'text/css; charset=utf-8'} '.js'{'application/javascript; charset=utf-8'} default{'application/octet-stream'} } }
function SendFile($s,[string]$p){ if(-not (Test-Path -LiteralPath $p)){ JResp $s 404 @{ok=$false;message='File not found.'}; return }; Resp $s 200 (CType $p) ([IO.File]::ReadAllBytes($p)) }
function Route($s,$r){
  if($r.Method -eq 'GET' -and ($r.Path -eq '/' -or $r.Path -eq '/index.html')){ SendFile $s (Join-Path $script:Web 'index.html'); return }
  if($r.Method -eq 'GET' -and $r.Path -eq '/styles.css'){ SendFile $s (Join-Path $script:Web 'styles.css'); return }
  if($r.Method -eq 'GET' -and $r.Path -eq '/app.js'){ SendFile $s (Join-Path $script:Web 'app.js'); return }
  if($r.Method -eq 'GET' -and $r.Path -eq '/api/state'){
    $sess=if($script:Session){ [ordered]@{ username=$script:Session.username; connectionUri=$script:Session.connectionUri; authentication=$script:Session.authentication; connectedAt=$script:Session.connectedAt } }else{$null}
    $ver=if($script:Verified){ [ordered]@{ verifiedAt=$script:Verified.verifiedAt; usersCount=$script:Verified.usersCount } }else{$null}
    JResp $s 200 @{ ok=$true; defaults=(Defaults); connected=($null -ne $script:Session); session=$sess; lastVerifiedSearch=$ver }
    return
  }
  if($r.Method -eq 'GET' -and $r.Path -eq '/api/users'){ JResp $s 200 @{ ok=$true; usersText=(LoadUsers) }; return }
  if($r.Method -eq 'POST' -and $r.Path -eq '/api/users'){ $b=Json $r.Body; $l=SaveUsers ([string]$b.users); JResp $s 200 @{ ok=$true; message="Saved $($l.Count) accounts to leaverlist.ps1."; usersText=($l -join "`n") }; return }
  if($r.Method -eq 'POST' -and $r.Path -eq '/api/connect'){
    $b=Json $r.Body
    if([string]::IsNullOrWhiteSpace([string]$b.connectionUri) -or [string]::IsNullOrWhiteSpace([string]$b.username) -or [string]::IsNullOrWhiteSpace([string]$b.password)){ throw 'Connection URI, username, and password are required.' }
    $x=DoConnect $b
    $script:Session=[ordered]@{ connectionUri=([string]$b.connectionUri).Trim(); username=([string]$b.username).Trim(); password=[string]$b.password; authentication=$(if([string]::IsNullOrWhiteSpace([string]$b.authentication)){'Kerberos'}else{([string]$b.authentication).Trim()}); connectedAt=(Get-Date).ToString('o') }
    $script:Verified=$null
    JResp $s 200 @{ ok=$true; message=$x.message; session=([ordered]@{ username=$script:Session.username; connectionUri=$script:Session.connectionUri; authentication=$script:Session.authentication; connectedAt=$script:Session.connectedAt }) }
    return
  }
  if($r.Method -eq 'POST' -and $r.Path -eq '/api/disconnect'){ $script:Session=$null; $script:Verified=$null; JResp $s 200 @{ ok=$true; message='Local Exchange session state cleared.' }; return }
  if($r.Method -eq 'POST' -and $r.Path -eq '/api/search'){
    NeedConn; $b=Json $r.Body; $u=NeedInput $b; if([string]::IsNullOrWhiteSpace([string]$b.targetMailbox)){ throw 'Search requires a target mailbox.' }
    $x=DoAction 'search' $script:Session $b $u ([string]$b.targetMailbox) ([string]$b.targetFolder)
    $script:Verified=[ordered]@{ fingerprint=(Finger $b); verifiedAt=(Get-Date).ToString('o'); usersCount=$u.Count }
    JResp $s 200 @{ ok=$true; message=$x.message; query=$x.query; results=$x.results; verifiedAt=$script:Verified.verifiedAt }
    return
  }
  if($r.Method -eq 'POST' -and $r.Path -eq '/api/delete'){
    NeedConn; $b=Json $r.Body; $u=NeedInput $b; NeedDelete $b
    $x=DoAction 'delete' $script:Session $b $u '' ''
    JResp $s 200 @{ ok=$true; message=$x.message; query=$x.query; results=$x.results }
    return
  }
  JResp $s 404 @{ ok=$false; message='Route not found.' }
}
function Client([System.Net.Sockets.TcpClient]$c){
  $c.ReceiveTimeout=10000
  $c.SendTimeout=10000
  $s=$c.GetStream()
  try{
    $r=Req $s
    Route $s $r
  } catch {
    $m=$_.Exception.Message
    if($m -match '0x80090311' -or $m -match "domain isn't available" -or ($m -match 'Kerberos' -and $m -match 'domain')){
      $m='Kerberos authentication failed. This machine is not using a domain sign-in, or it cannot reach the domain controller / Exchange host. Use a domain-joined machine on the corporate network or VPN, or switch Authentication to Negotiate or Basic if your Exchange allows it. Also prefer a FQDN or HTTPS endpoint instead of only a short host name.'
    }
    LogLine ("Request failed: " + $m)
    try { JResp $s 400 @{ ok=$false; message=$m } } catch { LogLine ("Response write failed: " + $_.Exception.Message) }
  } finally {
    try { $s.Dispose() } catch {}
    try { $c.Close() } catch {}
  }
}
if(-not (Test-Path -LiteralPath $script:Web)){ throw "Web UI directory not found: $script:Web" }
$listener=[System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse('127.0.0.1'),$Port)
$listener.Start()
Write-Host "Exchange PowerShell web server started at http://127.0.0.1:$Port" -ForegroundColor Green
LogLine ("Server started on port " + $Port)
try{
  while($true){
    try {
      $client = $listener.AcceptTcpClient()
      Client $client
    } catch {
      LogLine ("Listener loop error: " + $_.Exception.Message)
      Start-Sleep -Milliseconds 200
    }
  }
} finally {
  LogLine "Server stopping"
  $listener.Stop()
}




