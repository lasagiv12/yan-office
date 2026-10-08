# Microsoft 365 connector for the consulting-office plugin: an MCP server over stdio.
# Mail (search, read, attachments, drafts, send when allowed), calendar (read, create, update,
# cancel) and OneDrive (search, list) through Microsoft Graph, signed in as the consultant himself
# (device code flow, delegated permissions only, no client secret).
#
# Runs on Windows PowerShell 5.1, which ships with Windows, so nothing has to be installed.
# Config and the refresh token live in %APPDATA%\ConsultingOffice; the token is encrypted with
# Windows DPAPI for the current user. stdout carries JSON-RPC only; diagnostics go to stderr.
# Every write action that reaches other people (send, invitations, cancellations) requires
# confirmed=true, which the skills set only after the user approved that exact action.

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $Utf8
[Console]::OutputEncoding = $Utf8

$ServerVersion = '0.4.0'
$TimeZone = 'Israel Standard Time'
$Graph = 'https://graph.microsoft.com/v1.0'
$DataDir = Join-Path $env:APPDATA 'ConsultingOffice'
$ConfigPath = Join-Path $DataDir 'microsoft.json'
$TokenPath = Join-Path $DataDir 'microsoft-token.dat'
$PendingPath = Join-Path $DataDir 'microsoft-pending.json'
$script:AccessToken = $null
$script:AccessExpires = [DateTime]::MinValue

function Log($text) { [Console]::Error.WriteLine("[microsoft] $text") }

# ---------- config and secure storage ----------

function Get-Config {
  if (-not (Test-Path $ConfigPath)) { return $null }
  return Get-Content -Raw -Encoding UTF8 $ConfigPath | ConvertFrom-Json
}

function Save-Config($tenant, $client, $allowSend) {
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  $cfg = [ordered]@{ tenant_id = $tenant; client_id = $client; allow_send = [bool]$allowSend }
  [IO.File]::WriteAllText($ConfigPath, ($cfg | ConvertTo-Json), $Utf8)
}

function Get-Scopes($cfg) {
  $scopes = 'offline_access User.Read Mail.ReadWrite Calendars.ReadWrite Files.Read.All'
  if ($cfg -and $cfg.allow_send) { $scopes += ' Mail.Send' }
  return $scopes
}

function Save-RefreshToken($token) {
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  $enc = ConvertTo-SecureString -String $token -AsPlainText -Force | ConvertFrom-SecureString
  [IO.File]::WriteAllText($TokenPath, $enc, $Utf8)
}

function Get-RefreshToken {
  if (-not (Test-Path $TokenPath)) { return $null }
  $secure = (Get-Content -Raw $TokenPath).Trim() | ConvertTo-SecureString
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}

# ---------- HTTP ----------

function Read-ErrorBody($err) {
  # PowerShell puts the response body in ErrorDetails; Windows PowerShell 5.1 can also expose the raw stream.
  if ($err.ErrorDetails -and $err.ErrorDetails.Message) { return $err.ErrorDetails.Message }
  try {
    $resp = $err.Exception.Response
    if ($resp -and ($resp | Get-Member -Name GetResponseStream)) {
      $reader = New-Object IO.StreamReader($resp.GetResponseStream(), $Utf8)
      return $reader.ReadToEnd()
    }
  } catch {}
  return $err.Exception.Message
}

function Invoke-Form($url, $fields) {
  $pairs = foreach ($k in $fields.Keys) { [Uri]::EscapeDataString($k) + '=' + [Uri]::EscapeDataString([string]$fields[$k]) }
  $bytes = $Utf8.GetBytes(($pairs -join '&'))
  try {
    return Invoke-RestMethod -Method Post -Uri $url -Body $bytes -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
  } catch {
    $body = Read-ErrorBody $_
    try { return ($body | ConvertFrom-Json) } catch { throw "HTTP error: $body" }
  }
}

function Get-AccessToken {
  if ($script:AccessToken -and [DateTime]::UtcNow -lt $script:AccessExpires) { return $script:AccessToken }
  $cfg = Get-Config
  if (-not $cfg) { throw 'NOT_CONFIGURED: run ms_configure with tenant_id and client_id first.' }
  $refresh = Get-RefreshToken
  if (-not $refresh) { throw 'NOT_SIGNED_IN: run ms_sign_in, then ms_sign_in_complete.' }
  $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
    client_id = $cfg.client_id; grant_type = 'refresh_token'; refresh_token = $refresh; scope = (Get-Scopes $cfg)
  }
  if (-not ($r.PSObject.Properties.Name -contains 'access_token')) {
    throw "SIGN_IN_EXPIRED: $($r.error_description) Run ms_sign_in again."
  }
  if ($r.PSObject.Properties.Name -contains 'refresh_token') { Save-RefreshToken $r.refresh_token }
  $script:AccessToken = $r.access_token
  $script:AccessExpires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in - 120)
  return $script:AccessToken
}

function Invoke-Graph($method, $path, $body = $null, $prefer = $null) {
  $headers = @{ Authorization = "Bearer $(Get-AccessToken)" }
  if ($prefer) { $headers['Prefer'] = $prefer }
  $url = if ($path.StartsWith('https://')) { $path } else { "$Graph$path" }
  $req = @{ Method = $method; Uri = $url; Headers = $headers; UseBasicParsing = $true }
  if ($null -ne $body) {
    $req['Body'] = $Utf8.GetBytes((ConvertTo-Json -InputObject $body -Depth 20))
    $req['ContentType'] = 'application/json; charset=utf-8'
  }
  try { return Invoke-RestMethod @req }
  catch { throw "Graph $method $path failed: $(Read-ErrorBody $_)" }
}

function Q($text) { return [Uri]::EscapeDataString($text) }

function Get-Prop($obj, $name, $default = $null) {
  if ($null -ne $obj -and ($obj.PSObject.Properties.Name -contains $name) -and $null -ne $obj.$name) { return $obj.$name }
  return $default
}

function Require-Confirmed($a, $what) {
  if (-not (Get-Prop $a 'confirmed' $false)) {
    throw "CONFIRMATION_REQUIRED: $what affects other people. Show the user exactly what will happen, get explicit approval, then call again with confirmed=true."
  }
}

function Format-Address($r) {
  $addr = Get-Prop $r 'emailAddress' $null
  if ($null -eq $addr) { return '' }
  return "$(Get-Prop $addr 'name' '') <$(Get-Prop $addr 'address' '')>"
}

function Format-Message($m) {
  return [ordered]@{
    id = $m.id; subject = (Get-Prop $m 'subject' ''); from = (Format-Address (Get-Prop $m 'from' $null))
    to = @((Get-Prop $m 'toRecipients' @()) | ForEach-Object { Format-Address $_ })
    received = $m.receivedDateTime; is_read = $m.isRead; has_attachments = $m.hasAttachments
    preview = (Get-Prop $m 'bodyPreview' ''); web_link = (Get-Prop $m 'webLink' '')
  }
}

function Format-Event($e) {
  return [ordered]@{
    id = $e.id; subject = $e.subject; start = $e.start.dateTime; end = $e.end.dateTime
    location = (Get-Prop (Get-Prop $e 'location' $null) 'displayName' ''); organizer = (Format-Address (Get-Prop $e 'organizer' $null))
    attendees = @((Get-Prop $e 'attendees' @()) | ForEach-Object { Format-Address $_ })
    is_online = (Get-Prop $e 'isOnlineMeeting' $false); preview = (Get-Prop $e 'bodyPreview' ''); web_link = (Get-Prop $e 'webLink' '')
  }
}

function To-Recipients($list) {
  return @(@($list) | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = [string]$_ } } })
}

# ---------- tools ----------

$Tools = @(
  @{ name = 'ms_status'; title = 'מצב החיבור ל-Microsoft'; readOnly = $true
     description = 'Shows whether the Microsoft 365 connection is configured and signed in, and as whom. Call this first when mail, calendar or OneDrive is needed.'
     props = @{}; required = @() },
  @{ name = 'ms_configure'; title = 'הגדרת החיבור'; readOnly = $false
     description = 'One-time setup: saves the Entra tenant ID and app (client) ID from the app registration. allow_send enables sending mail (default false: drafts only).'
     props = @{ tenant_id = @{ type = 'string' }; client_id = @{ type = 'string' }; allow_send = @{ type = 'boolean' } }; required = @('tenant_id', 'client_id') },
  @{ name = 'ms_sign_in'; title = 'כניסה לחשבון Microsoft'; readOnly = $false
     description = 'Starts sign-in. Returns a short code and a web address; tell the user to open the address, enter the code and sign in with his own Microsoft 365 account. Then call ms_sign_in_complete.'
     props = @{}; required = @() },
  @{ name = 'ms_sign_in_complete'; title = 'סיום כניסה'; readOnly = $false
     description = 'Waits (up to about 45 seconds) for the user to finish the sign-in started by ms_sign_in and stores the credentials securely on this computer.'
     props = @{}; required = @() },
  @{ name = 'mail_search'; title = 'חיפוש מיילים'; readOnly = $true
     description = 'Searches the signed-in mailbox. Use from (address or domain, e.g. a client domain from the client index), query (free text), since/until (YYYY-MM-DD). Returns newest first.'
     props = @{ query = @{ type = 'string' }; from = @{ type = 'string' }; since = @{ type = 'string' }; until = @{ type = 'string' }; top = @{ type = 'integer' } }; required = @() },
  @{ name = 'mail_read'; title = 'קריאת מייל'; readOnly = $true
     description = 'Returns the full plain-text body and recipients of one message.'
     props = @{ message_id = @{ type = 'string' } }; required = @('message_id') },
  @{ name = 'mail_attachments'; title = 'קבצים מצורפים'; readOnly = $true
     description = 'Lists the attachments of a message (id, name, size, type).'
     props = @{ message_id = @{ type = 'string' } }; required = @('message_id') },
  @{ name = 'mail_save_attachment'; title = 'שמירת קובץ מצורף'; readOnly = $false
     description = 'Saves one attachment to a local folder (for example the client''s 07 לתיוק folder). Never overwrites: adds a number if the name exists.'
     props = @{ message_id = @{ type = 'string' }; attachment_id = @{ type = 'string' }; folder = @{ type = 'string' } }; required = @('message_id', 'attachment_id', 'folder') },
  @{ name = 'mail_create_draft'; title = 'טיוטת מייל'; readOnly = $false
     description = 'Creates a draft in the user''s Outlook Drafts folder (nothing is sent). For a reply pass reply_to_message_id; the text is placed above the quoted message. Body is plain text; Hebrew is fine.'
     props = @{ to = @{ type = 'array'; items = @{ type = 'string' } }; cc = @{ type = 'array'; items = @{ type = 'string' } }; subject = @{ type = 'string' }; body = @{ type = 'string' }; reply_to_message_id = @{ type = 'string' } }; required = @('body') },
  @{ name = 'mail_send_draft'; title = 'שליחת טיוטה'; readOnly = $false; destructive = $true
     description = 'Sends an existing draft. Works only if sending was enabled in ms_configure (allow_send) and only with confirmed=true after the user approved sending this exact draft.'
     props = @{ draft_id = @{ type = 'string' }; confirmed = @{ type = 'boolean' } }; required = @('draft_id', 'confirmed') },
  @{ name = 'calendar_events'; title = 'אירועים ביומן'; readOnly = $true
     description = 'Lists calendar events between start and end (YYYY-MM-DD or YYYY-MM-DDTHH:MM, Israel time). Optional query filters by text in the subject. Use it to answer what is scheduled and when the user is free.'
     props = @{ start = @{ type = 'string' }; end = @{ type = 'string' }; query = @{ type = 'string' } }; required = @('start', 'end') },
  @{ name = 'calendar_event_read'; title = 'קריאת אירוע'; readOnly = $true
     description = 'Returns one calendar event with its full plain-text notes (the consultant writes meeting notes inside the event). Use it to bring past meeting notes into the client file.'
     props = @{ event_id = @{ type = 'string' } }; required = @('event_id') },
  @{ name = 'calendar_create_event'; title = 'קביעת אירוע'; readOnly = $false; destructive = $true
     description = 'Creates an event in the user''s own calendar (Israel time). With attendees, Outlook sends them invitations, so attendees require confirmed=true after the user approved the exact event.'
     props = @{ subject = @{ type = 'string' }; start = @{ type = 'string' }; end = @{ type = 'string' }; location = @{ type = 'string' }; body = @{ type = 'string' }; attendees = @{ type = 'array'; items = @{ type = 'string' } }; online_meeting = @{ type = 'boolean' }; confirmed = @{ type = 'boolean' } }; required = @('subject', 'start', 'end') },
  @{ name = 'calendar_update_event'; title = 'עדכון אירוע'; readOnly = $false; destructive = $true
     description = 'Changes an event (subject, start, end, location, body). If the event has attendees they are notified, so confirmed=true is required after user approval.'
     props = @{ event_id = @{ type = 'string' }; subject = @{ type = 'string' }; start = @{ type = 'string' }; end = @{ type = 'string' }; location = @{ type = 'string' }; body = @{ type = 'string' }; confirmed = @{ type = 'boolean' } }; required = @('event_id') },
  @{ name = 'calendar_cancel_event'; title = 'ביטול אירוע'; readOnly = $false; destructive = $true
     description = 'Cancels an event. Attendees receive a cancellation with the optional comment. Always requires confirmed=true after user approval.'
     props = @{ event_id = @{ type = 'string' }; comment = @{ type = 'string' }; confirmed = @{ type = 'boolean' } }; required = @('event_id', 'confirmed') },
  @{ name = 'onedrive_search'; title = 'חיפוש ב-OneDrive'; readOnly = $true
     description = 'Searches the user''s OneDrive in the cloud by file name and content. For reading and writing client files on this computer, use the synced folder instead.'
     props = @{ query = @{ type = 'string' }; top = @{ type = 'integer' } }; required = @('query') },
  @{ name = 'onedrive_list'; title = 'רשימת קבצים ב-OneDrive'; readOnly = $true
     description = 'Lists a OneDrive folder in the cloud (path relative to the OneDrive root, empty for the root).'
     props = @{ path = @{ type = 'string' } }; required = @() }
)

function To-DateTime($text, $endOfDay) {
  if ($text -match '^\d{4}-\d{2}-\d{2}$') { return $text + $(if ($endOfDay) { 'T23:59:59' } else { 'T00:00:00' }) }
  if ($text -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$') { return "$($text):00" }
  return $text
}

function Tool-Status($a) {
  $cfg = Get-Config
  if (-not $cfg) { return 'לא מוגדר. צריך ms_configure עם tenant_id ו-client_id מרישום האפליקציה ב-Entra.' }
  if (-not (Test-Path $TokenPath)) { return "מוגדר (tenant $($cfg.tenant_id)), אבל עוד לא נכנסו לחשבון. הצעד הבא: ms_sign_in." }
  $me = Invoke-Graph 'GET' '/me?$select=displayName,mail,userPrincipalName'
  $send = if ($cfg.allow_send) { 'מופעלת (רק עם אישור לכל מייל)' } else { 'כבויה: טיוטות בלבד' }
  return "מחובר בשם $($me.displayName) ($($me.mail)). שליחת מייל: $send."
}

function Tool-Configure($a) {
  if ($a.tenant_id -notmatch '^[0-9a-fA-F-]{36}$' -or $a.client_id -notmatch '^[0-9a-fA-F-]{36}$') {
    throw 'tenant_id ו-client_id צריכים להיות מזהים בפורמט GUID (36 תווים) מעמוד האפליקציה ב-Entra.'
  }
  Save-Config $a.tenant_id $a.client_id (Get-Prop $a 'allow_send' $false)
  if (Test-Path $TokenPath) { Remove-Item $TokenPath -Force }
  $script:AccessToken = $null
  return 'נשמר. הצעד הבא: ms_sign_in.'
}

function Tool-SignIn($a) {
  $cfg = Get-Config
  if (-not $cfg) { throw 'NOT_CONFIGURED: run ms_configure first.' }
  $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/devicecode" @{ client_id = $cfg.client_id; scope = (Get-Scopes $cfg) }
  if (-not ($r.PSObject.Properties.Name -contains 'device_code')) { throw "Sign-in could not start: $($r.error_description)" }
  $pending = [ordered]@{ device_code = $r.device_code; interval = [int]$r.interval; expires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in).ToString('o') }
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  [IO.File]::WriteAllText($PendingPath, ($pending | ConvertTo-Json), $Utf8)
  return "פתח בדפדפן: $($r.verification_uri)`nהזן את הקוד: $($r.user_code)`nוהתחבר עם החשבון של Microsoft 365. אחרי שזה הסתיים, הפעל ms_sign_in_complete."
}

function Tool-SignInComplete($a) {
  $cfg = Get-Config
  if (-not (Test-Path $PendingPath)) { throw 'אין כניסה פתוחה. הפעל ms_sign_in קודם.' }
  $p = Get-Content -Raw -Encoding UTF8 $PendingPath | ConvertFrom-Json
  # Stay under the host's per-tool timeout; the skill simply calls again if sign-in is still pending.
  $deadline = [DateTime]::UtcNow.AddSeconds(45)
  $interval = [Math]::Max(5, [int]$p.interval)
  while ([DateTime]::UtcNow -lt $deadline) {
    $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
      client_id = $cfg.client_id; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $p.device_code
    }
    if ($r.PSObject.Properties.Name -contains 'access_token') {
      Save-RefreshToken $r.refresh_token
      Remove-Item $PendingPath -Force
      $script:AccessToken = $r.access_token
      $script:AccessExpires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in - 120)
      return (Tool-Status $a)
    }
    switch ($r.error) {
      'authorization_pending' { Start-Sleep -Seconds $interval }
      'slow_down' { $interval += 5; Start-Sleep -Seconds $interval }
      default { Remove-Item $PendingPath -Force -ErrorAction SilentlyContinue; throw "הכניסה נכשלה: $($r.error) $($r.error_description)" }
    }
  }
  return 'הכניסה עוד לא הושלמה. אחרי שמסיימים בדפדפן, הפעל ms_sign_in_complete שוב.'
}

function Tool-MailSearch($a) {
  $top = [Math]::Min([int](Get-Prop $a 'top' 15), 50)
  $sel = '$select=id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink'
  $query = Get-Prop $a 'query' ''
  $from = Get-Prop $a 'from' ''
  $since = Get-Prop $a 'since' ''
  $until = Get-Prop $a 'until' ''
  if ($query -or ($from -and -not $from.Contains('@'))) {
    # $search cannot be combined with $filter/$orderby, so sender, domain and dates go into the KQL query.
    $kql = $query
    if ($from) { $kql = ("$kql from:" + $from.TrimStart('@')).Trim() }
    if ($since) { $kql += " received>=$since" }
    if ($until) { $kql += " received<=$until" }
    $r = Invoke-Graph 'GET' "/me/messages?`$search=$(Q ('"' + $kql + '"'))&`$top=$top&$sel"
  } else {
    $filters = @()
    if ($since) { $filters += "receivedDateTime ge $(To-DateTime $since $false)Z" }
    if ($until) { $filters += "receivedDateTime le $(To-DateTime $until $true)Z" }
    if ($from -and $from.Contains('@')) { $filters += "from/emailAddress/address eq '$($from.Replace("'", "''"))'" }
    $url = "/me/messages?`$top=$top&$sel&`$orderby=receivedDateTime desc"
    if ($filters.Count) { $url += "&`$filter=$(Q ($filters -join ' and '))" }
    $r = Invoke-Graph 'GET' $url
  }
  return @($r.value | ForEach-Object { Format-Message $_ })
}

function Tool-MailRead($a) {
  $m = Invoke-Graph 'GET' "/me/messages/$($a.message_id)?`$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,webLink" 'outlook.body-content-type="text"'
  $out = Format-Message $m
  $out['cc'] = @($m.ccRecipients | ForEach-Object { Format-Address $_ })
  $out['body'] = $m.body.content
  return $out
}

function Tool-MailAttachments($a) {
  $r = Invoke-Graph 'GET' "/me/messages/$($a.message_id)/attachments?`$select=id,name,size,contentType"
  return @($r.value | ForEach-Object { [ordered]@{ id = $_.id; name = $_.name; size = $_.size; type = $_.contentType } })
}

function Tool-MailSaveAttachment($a) {
  if (-not (Test-Path -LiteralPath $a.folder -PathType Container)) { throw "התיקייה לא קיימת: $($a.folder)" }
  $att = Invoke-Graph 'GET' "/me/messages/$($a.message_id)/attachments/$($a.attachment_id)"
  if (-not ($att.PSObject.Properties.Name -contains 'contentBytes')) { throw 'זה לא קובץ מצורף רגיל (למשל פריט Outlook מקושר) ולא ניתן לשמור אותו.' }
  $name = ($att.name -replace '[\\/:*?"<>|]', '_')
  $target = Join-Path $a.folder $name
  $i = 2
  while (Test-Path -LiteralPath $target) {
    $target = Join-Path $a.folder ("{0} ({1}){2}" -f [IO.Path]::GetFileNameWithoutExtension($name), $i, [IO.Path]::GetExtension($name)); $i++
  }
  [IO.File]::WriteAllBytes($target, [Convert]::FromBase64String($att.contentBytes))
  return "נשמר: $target"
}

function Tool-MailCreateDraft($a) {
  $body = [string]$a.body
  $replyTo = Get-Prop $a 'reply_to_message_id' ''
  if ($replyTo) {
    $draft = Invoke-Graph 'POST' "/me/messages/$replyTo/createReply" @{}
    $html = '<div dir="rtl" style="font-family:Arial">' + [Net.WebUtility]::HtmlEncode($body).Replace("`n", '<br>') + '</div><br>' + $draft.body.content
    $patch = @{ body = @{ contentType = 'HTML'; content = $html } }
    $to = Get-Prop $a 'to' $null
    if ($to) { $patch['toRecipients'] = To-Recipients $to }
    $cc = Get-Prop $a 'cc' $null
    if ($cc) { $patch['ccRecipients'] = To-Recipients $cc }
    $draft = Invoke-Graph 'PATCH' "/me/messages/$($draft.id)" $patch
  } else {
    $html = '<div dir="rtl" style="font-family:Arial">' + [Net.WebUtility]::HtmlEncode($body).Replace("`n", '<br>') + '</div>'
    $msg = @{ subject = [string](Get-Prop $a 'subject' ''); body = @{ contentType = 'HTML'; content = $html }; toRecipients = (To-Recipients (Get-Prop $a 'to' @())) }
    $cc = Get-Prop $a 'cc' $null
    if ($cc) { $msg['ccRecipients'] = To-Recipients $cc }
    $draft = Invoke-Graph 'POST' '/me/messages' $msg
  }
  return [ordered]@{ draft_id = $draft.id; subject = $draft.subject; web_link = (Get-Prop $draft 'webLink' ''); note = 'הטיוטה נשמרה בתיקיית הטיוטות ב-Outlook. שום דבר לא נשלח.' }
}

function Tool-MailSendDraft($a) {
  $cfg = Get-Config
  if (-not $cfg.allow_send) { throw 'שליחת מייל כבויה בהגדרות (טיוטות בלבד). הטיוטה מחכה ב-Outlook לשליחה ידנית.' }
  Require-Confirmed $a 'Sending this email'
  Invoke-Graph 'POST' "/me/messages/$($a.draft_id)/send" | Out-Null
  return 'המייל נשלח.'
}

function Tool-CalendarEvents($a) {
  $start = To-DateTime $a.start $false
  $end = To-DateTime $a.end $true
  $url = "/me/calendarView?startDateTime=$(Q $start)&endDateTime=$(Q $end)&`$top=100&`$orderby=start/dateTime&`$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,bodyPreview,webLink"
  $r = Invoke-Graph 'GET' $url "outlook.timezone=`"$TimeZone`""
  $events = @($r.value)
  $query = Get-Prop $a 'query' ''
  if ($query) { $events = @($events | Where-Object { $_.subject -like "*$query*" }) }
  return @($events | ForEach-Object { Format-Event $_ })
}

function Tool-CalendarEventRead($a) {
  $e = Invoke-Graph 'GET' "/me/events/$($a.event_id)?`$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,bodyPreview,webLink,body" "outlook.timezone=`"$TimeZone`", outlook.body-content-type=`"text`""
  $out = Format-Event $e
  $out['notes'] = (Get-Prop (Get-Prop $e 'body' $null) 'content' '')
  return $out
}

function Tool-CalendarCreate($a) {
  $attendees = @(Get-Prop $a 'attendees' @())
  if ($attendees.Count) { Require-Confirmed $a 'Inviting attendees' }
  $event = @{
    subject = [string]$a.subject
    start = @{ dateTime = (To-DateTime $a.start $false); timeZone = $TimeZone }
    end = @{ dateTime = (To-DateTime $a.end $false); timeZone = $TimeZone }
  }
  $loc = Get-Prop $a 'location' ''
  if ($loc) { $event['location'] = @{ displayName = $loc } }
  $text = Get-Prop $a 'body' ''
  if ($text) { $event['body'] = @{ contentType = 'Text'; content = $text } }
  if ($attendees.Count) { $event['attendees'] = @($attendees | ForEach-Object { @{ emailAddress = @{ address = [string]$_ }; type = 'required' } }) }
  if (Get-Prop $a 'online_meeting' $false) { $event['isOnlineMeeting'] = $true; $event['onlineMeetingProvider'] = 'teamsForBusiness' }
  $e = Invoke-Graph 'POST' '/me/events' $event "outlook.timezone=`"$TimeZone`""
  return (Format-Event $e)
}

function Tool-CalendarUpdate($a) {
  $current = Invoke-Graph 'GET' "/me/events/$($a.event_id)?`$select=attendees"
  if (@(Get-Prop $current 'attendees' @()).Count) { Require-Confirmed $a 'Changing an event with attendees' }
  $patch = @{}
  if (Get-Prop $a 'subject' '') { $patch['subject'] = $a.subject }
  if (Get-Prop $a 'start' '') { $patch['start'] = @{ dateTime = (To-DateTime $a.start $false); timeZone = $TimeZone } }
  if (Get-Prop $a 'end' '') { $patch['end'] = @{ dateTime = (To-DateTime $a.end $false); timeZone = $TimeZone } }
  if (Get-Prop $a 'location' '') { $patch['location'] = @{ displayName = $a.location } }
  if (Get-Prop $a 'body' '') { $patch['body'] = @{ contentType = 'Text'; content = $a.body } }
  if ($patch.Count -eq 0) { throw 'לא נמסר שום שדה לעדכון.' }
  $e = Invoke-Graph 'PATCH' "/me/events/$($a.event_id)" $patch "outlook.timezone=`"$TimeZone`""
  return (Format-Event $e)
}

function Tool-CalendarCancel($a) {
  Require-Confirmed $a 'Cancelling an event'
  $e = Invoke-Graph 'GET' "/me/events/$($a.event_id)?`$select=isOrganizer,attendees,subject"
  if ($e.isOrganizer -and @(Get-Prop $e 'attendees' @()).Count) {
    Invoke-Graph 'POST' "/me/events/$($a.event_id)/cancel" @{ comment = [string](Get-Prop $a 'comment' '') } | Out-Null
  } else {
    Invoke-Graph 'DELETE' "/me/events/$($a.event_id)" | Out-Null
  }
  return "האירוע בוטל: $($e.subject)"
}

function Tool-OneDriveSearch($a) {
  $top = [Math]::Min([int](Get-Prop $a 'top' 20), 50)
  $q = $a.query.Replace("'", "''")
  $r = Invoke-Graph 'GET' "/me/drive/root/search(q='$(Q $q)')?`$top=$top&`$select=id,name,webUrl,lastModifiedDateTime,size,parentReference,file,folder"
  return @($r.value | ForEach-Object {
    [ordered]@{ name = $_.name; path = (Get-Prop $_.parentReference 'path' ''); modified = $_.lastModifiedDateTime; size = $_.size; web_url = $_.webUrl; is_folder = [bool](Get-Prop $_ 'folder' $null) }
  })
}

function Tool-OneDriveList($a) {
  $path = ([string](Get-Prop $a 'path' '')).Trim('/')
  $url = if ($path) { "/me/drive/root:/$((($path -split '/') | ForEach-Object { Q $_ }) -join '/'):/children" } else { '/me/drive/root/children' }
  $r = Invoke-Graph 'GET' "$url?`$top=200&`$select=name,webUrl,lastModifiedDateTime,size,file,folder"
  return @($r.value | ForEach-Object { [ordered]@{ name = $_.name; modified = $_.lastModifiedDateTime; size = $_.size; is_folder = [bool](Get-Prop $_ 'folder' $null); web_url = $_.webUrl } })
}

$Handlers = @{
  ms_status = ${function:Tool-Status}; ms_configure = ${function:Tool-Configure}
  ms_sign_in = ${function:Tool-SignIn}; ms_sign_in_complete = ${function:Tool-SignInComplete}
  mail_search = ${function:Tool-MailSearch}; mail_read = ${function:Tool-MailRead}
  mail_attachments = ${function:Tool-MailAttachments}; mail_save_attachment = ${function:Tool-MailSaveAttachment}
  mail_create_draft = ${function:Tool-MailCreateDraft}; mail_send_draft = ${function:Tool-MailSendDraft}
  calendar_events = ${function:Tool-CalendarEvents}; calendar_event_read = ${function:Tool-CalendarEventRead}
  calendar_create_event = ${function:Tool-CalendarCreate}
  calendar_update_event = ${function:Tool-CalendarUpdate}; calendar_cancel_event = ${function:Tool-CalendarCancel}
  onedrive_search = ${function:Tool-OneDriveSearch}; onedrive_list = ${function:Tool-OneDriveList}
}

# ---------- MCP over stdio ----------

function Send-Message($obj) {
  [Console]::Out.WriteLine((ConvertTo-Json -InputObject $obj -Depth 30 -Compress))
  [Console]::Out.Flush()
}

function Tool-List {
  return @($Tools | ForEach-Object {
    $t = $_
    $destructive = $false
    if ($t.ContainsKey('destructive')) { $destructive = $t.destructive }
    [ordered]@{
      name = $t.name; title = $t.title; description = $t.description
      inputSchema = [ordered]@{ type = 'object'; properties = $t.props; required = $t.required }
      annotations = [ordered]@{ readOnlyHint = $t.readOnly; destructiveHint = $destructive; openWorldHint = $true }
    }
  })
}

Log "starting $ServerVersion"
while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  if (-not $line.Trim()) { continue }
  try { $msg = $line | ConvertFrom-Json } catch { Log "bad json: $line"; continue }
  $hasId = $msg.PSObject.Properties.Name -contains 'id'
  try {
    switch ($msg.method) {
      'initialize' {
        $requested = Get-Prop $msg.params 'protocolVersion' '2025-06-18'
        Send-Message ([ordered]@{ jsonrpc = '2.0'; id = $msg.id; result = [ordered]@{
          protocolVersion = $requested
          capabilities = @{ tools = @{ listChanged = $false } }
          serverInfo = [ordered]@{ name = 'consulting-office-microsoft'; version = $ServerVersion }
        } })
      }
      'ping' { Send-Message @{ jsonrpc = '2.0'; id = $msg.id; result = @{} } }
      'tools/list' { Send-Message @{ jsonrpc = '2.0'; id = $msg.id; result = @{ tools = @(Tool-List) } } }
      'tools/call' {
        $name = $msg.params.name
        $toolArgs = Get-Prop $msg.params 'arguments' ([pscustomobject]@{})
        if (-not $Handlers.ContainsKey($name)) { throw "Unknown tool: $name" }
        try {
          $items = @(& $Handlers[$name] $toolArgs)
          $text = if ($items.Count -eq 1 -and $items[0] -is [string]) { $items[0] } else { ConvertTo-Json -InputObject $items -Depth 20 }
          Send-Message @{ jsonrpc = '2.0'; id = $msg.id; result = @{ content = @(@{ type = 'text'; text = $text }); isError = $false } }
        } catch {
          Send-Message @{ jsonrpc = '2.0'; id = $msg.id; result = @{ content = @(@{ type = 'text'; text = "$($_.Exception.Message)" }); isError = $true } }
        }
      }
      default {
        if ($hasId) { Send-Message @{ jsonrpc = '2.0'; id = $msg.id; error = @{ code = -32601; message = "Method not found: $($msg.method)" } } }
      }
    }
  } catch {
    Log "error: $($_.Exception.Message)"
    if ($hasId) { Send-Message @{ jsonrpc = '2.0'; id = $msg.id; error = @{ code = -32603; message = "$($_.Exception.Message)" } } }
  }
}
