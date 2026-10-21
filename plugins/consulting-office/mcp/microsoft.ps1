# Microsoft 365 connector for the consulting-office plugin: an MCP server over stdio.
# Mail (search, read, attachments, drafts, send when allowed), calendar (read, create, update,
# cancel), contacts (search, create, update) and OneDrive (search, list) through Microsoft Graph, signed in as the consultant himself
# (device code or browser sign-in with PKCE, delegated permissions only, no client secret).
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
# Read and write the raw stdio streams as UTF-8. Setting [Console]::*Encoding throws in Windows PowerShell
# when the host starts us without a console window, which is how desktop apps launch MCP servers.
$StdIn = New-Object IO.StreamReader([Console]::OpenStandardInput(), $Utf8)
$StdOut = New-Object IO.StreamWriter([Console]::OpenStandardOutput(), $Utf8)
$StdOut.AutoFlush = $true
# Windows PowerShell 5.1 can serialize arrays as {"value":[...],"Count":n} because of this type data.
Remove-TypeData -TypeName System.Array -ErrorAction SilentlyContinue

$ServerVersion = '0.7.2'
$TimeZone = 'Israel Standard Time'
$Graph = 'https://graph.microsoft.com/v1.0'
$DataDir = Join-Path $env:APPDATA 'ConsultingOffice'
$ConfigPath = Join-Path $DataDir 'microsoft.json'
$TokenPath = Join-Path $DataDir 'microsoft-token.dat'
$PendingPath = Join-Path $DataDir 'microsoft-pending.json'
$LogPath = Join-Path $DataDir 'microsoft-log.txt'
$script:AccessToken = $null
$script:AccessExpires = [DateTime]::MinValue
$script:Browser = $null
# False when the stored sign-in predates the contacts permission: mail and calendar keep working, contacts need one new sign-in.
$script:ContactsAllowed = $true

function Log($text) { [Console]::Error.WriteLine("[microsoft] $text") }

# Sign-in history (never tokens), so ms_check can show why and when the connection asked to sign in again.
function Write-AuthLog($text) {
  try {
    New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
    $lines = @()
    if (Test-Path $LogPath) { $lines = @(Get-Content -Encoding UTF8 $LogPath | Select-Object -Last 39) }
    $text = "$text"; if ($text.Length -gt 300) { $text = $text.Substring(0, 300) }
    $lines += "$((Get-Date).ToString('dd/MM/yyyy HH:mm')) $text"
    [IO.File]::WriteAllLines($LogPath, [string[]]$lines, $Utf8)
  } catch {}
}

# ---------- config and secure storage ----------

# Without a registration of its own, sign-in uses Microsoft's public "Microsoft Graph Command Line Tools"
# client (the one Graph PowerShell uses) against any work account, so no IDs are needed. ms_configure can
# still point at an app registered in the consultant's own tenant if his IT admin prefers that.
$DefaultTenant = 'organizations'
$DefaultClient = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

function Get-Config {
  if (-not (Test-Path $ConfigPath)) {
    return [pscustomobject]@{ tenant_id = $DefaultTenant; client_id = $DefaultClient; allow_send = $false; default = $true }
  }
  return Get-Content -Raw -Encoding UTF8 $ConfigPath | ConvertFrom-Json
}

function Get-AdminConsentUrl {
  $cfg = Get-Config
  # Always include Mail.Send, so one approval also covers turning sending on later. Drafts-only stays the default.
  $scopes = (@('User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite', 'Contacts.ReadWrite', 'Files.Read.All') | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' '
  return "https://login.microsoftonline.com/$($cfg.tenant_id)/v2.0/adminconsent?client_id=$($cfg.client_id)&scope=$([uri]::EscapeDataString($scopes))&redirect_uri=$([uri]::EscapeDataString('https://login.microsoftonline.com/common/oauth2/nativeclient'))"
}

function Explain-AuthError($code, $text) {
  # Wrong IDs first: Microsoft's text for AADSTS700016 also contains the word "consented".
  if ("$code $text" -match 'AADSTS700016|AADSTS90002|AADSTS900023') {
    return "במחשב שמורים מזהי אפליקציה או ארגון שלא קיימים אצל Microsoft (כנראה מזהים שהוזנו בניסיון קודם). לא צריך מזהים: מחזירים לברירת המחדל עם ms_configure ו-reset=true, ואז ms_sign_in. ($code)"
  }
  # Expired sign-ins, before the consent and Conditional Access checks: their texts mention both.
  if ("$code $text" -match 'AADSTS70043|AADSTS70044|AADSTS50078|AADSTS50076|AADSTS50079') {
    return "מדיניות האבטחה של הארגון מחייבת להיכנס מחדש כל פרק זמן קבוע (Sign-in frequency או אימות דו-שלבי חוזר). זה יחזור על עצמו עד שאיש המחשבים יאריך את התקופה או יחריג את האפליקציה Microsoft Graph Command Line Tools. ($code)"
  }
  if ("$code $text" -match 'AADSTS50173|AADSTS50133|AADSTS50089') {
    return "הסיסמה של החשבון הוחלפה, או שהחשבון נותק מכל המכשירים. צריך להיכנס שוב פעם אחת, והחיבור יישאר. ($code)"
  }
  if ("$code $text" -match 'AADSTS700082|AADSTS700081') {
    return "לא השתמשו בחיבור יותר מ-90 יום, ולכן הכניסה פגה. צריך להיכנס שוב פעם אחת. ($code)"
  }
  if ("$code $text" -match 'AADSTS65001|AADSTS90094|AADSTS90008|consent|admin approval|admin_consent') {
    return "צריך אישור חד-פעמי של מנהל Microsoft 365 בארגון (לרוב איש המחשבים). הוא פותח את הקישור הזה, נכנס עם חשבון המנהל שלו ומאשר (Accept): $(Get-AdminConsentUrl) . אחר כך מפעילים ms_sign_in שוב. ($code)"
  }
  if ("$code $text" -match 'AADSTS50020|AADSTS500200|AADSTS50034|personal Microsoft account|MSA') {
    return "נכנסת עם חשבון שלא שייך לארגון (למשל חשבון אישי של Microsoft עם אותה כתובת). להיכנס עם הכתובת שמופיעה ב-Outlook בתיבת הדואר, ואם נשאלים 'Work or school account' או 'Personal account', לבחור Work or school. ($code)"
  }
  if ("$code $text" -match 'AADSTS53003|Conditional Access') {
    return "מדיניות האבטחה של הארגון (Conditional Access) חוסמת את הכניסה. רק מנהל המערכת יכול לאפשר. ($code)"
  }
  return "$code $text"
}

function Save-Config($tenant, $client, $allowSend) {
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  $cfg = [ordered]@{ tenant_id = $tenant; client_id = $client; allow_send = [bool]$allowSend }
  [IO.File]::WriteAllText($ConfigPath, ($cfg | ConvertTo-Json), $Utf8)
}

function Get-Scopes($cfg, [switch]$WithoutContacts) {
  $scopes = 'offline_access User.Read Mail.ReadWrite Calendars.ReadWrite Files.Read.All'
  if (-not $WithoutContacts) { $scopes += ' Contacts.ReadWrite' }
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
  $refresh = Get-RefreshToken
  if (-not $refresh) { throw 'NOT_SIGNED_IN: run ms_sign_in, then ms_sign_in_complete.' }
  $temporary = 'TEMPORARY: אין כרגע תשובה תקינה משרתי הכניסה של Microsoft (רשת או תקלה זמנית). הכניסה השמורה בסדר, לא צריך להתחבר שוב. לנסות שוב בעוד דקה.'
  try {
    $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
      client_id = $cfg.client_id; grant_type = 'refresh_token'; refresh_token = $refresh; scope = (Get-Scopes $cfg)
    }
    $script:ContactsAllowed = $true
    # A sign-in from before contacts were added has no consent for Contacts.ReadWrite (AADSTS65001).
    # Renew without it, so mail and calendar keep working; the contacts tools then ask for one new sign-in.
    if ($null -ne $r -and $null -eq $r.PSObject.Properties['access_token'] -and "$(Get-Prop $r 'error_description' '')" -match 'AADSTS65001') {
      $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
        client_id = $cfg.client_id; grant_type = 'refresh_token'; refresh_token = $refresh; scope = (Get-Scopes $cfg -WithoutContacts)
      }
      $script:ContactsAllowed = $false
    }
  } catch {
    # No answer from Microsoft (network, proxy, outage): the stored sign-in is still valid.
    Write-AuthLog "תקלה זמנית בחידוש הכניסה: $($_.Exception.Message)"
    throw $temporary
  }
  if ($null -eq $r -or $null -eq $r.PSObject.Properties['access_token']) {
    $code = Get-Prop $r 'error' ''
    $desc = Get-Prop $r 'error_description' ''
    if ($code -in @('temporarily_unavailable', 'server_error', '')) { Write-AuthLog "תקלה זמנית בחידוש הכניסה: $code"; throw $temporary }
    $aad = if ("$desc" -match 'AADSTS\d+') { $Matches[0] } else { $code }
    Write-AuthLog "הכניסה פגה ($aad)"
    throw ("SIGN_IN_EXPIRED: " + (Explain-AuthError $code $desc) + " Run ms_sign_in again.")
  }
  if ($null -ne $r.PSObject.Properties['refresh_token']) { Save-RefreshToken $r.refresh_token }
  $script:AccessToken = $r.access_token
  $script:AccessExpires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in - 120)
  return $script:AccessToken
}

function Invoke-Graph($method, $path, $body = $null, $Prefer = $null, [switch]$Retried) {
  # Pass the Prefer header by name (-Prefer): positionally it lands in $body, and Windows PowerShell refuses a body on GET.
  if ($null -ne $body -and $method -in @('GET', 'DELETE')) { throw "Invoke-Graph: $method $path cannot carry a body" }
  $headers = @{ Authorization = "Bearer $(Get-AccessToken)" }
  if ($Prefer) { $headers['Prefer'] = $Prefer }
  $url = if ($path.StartsWith('https://')) { $path } else { "$Graph$path" }
  $req = @{ Method = $method; Uri = $url; Headers = $headers; UseBasicParsing = $true }
  if ($null -ne $body) {
    $req['Body'] = $Utf8.GetBytes((ConvertTo-Json -InputObject $body -Depth 20))
    $req['ContentType'] = 'application/json; charset=utf-8'
  }
  try { return Invoke-RestMethod @req }
  catch {
    $err = $_
    $status = 0
    try { $status = [int]$err.Exception.Response.StatusCode } catch {}
    # Graph rejected the access token before its expiry (revoked or policy change): renew it once.
    if ($status -eq 401 -and -not $Retried) {
      $script:AccessToken = $null
      return (Invoke-Graph $method $path $body $Prefer -Retried)
    }
    throw "Graph $method $path failed: $(Read-ErrorBody $err)"
  }
}

function Q($text) { return [Uri]::EscapeDataString($text) }

function Get-Prop($obj, $name, $default = $null) {
  # Properties[$name] instead of Properties.Name: on an object with no properties, .Name throws under strict mode.
  if ($null -eq $obj) { return $default }
  if ($obj -is [System.Collections.IDictionary]) { if ($obj.Contains($name) -and $null -ne $obj[$name]) { return $obj[$name] }; return $default }
  $prop = $obj.PSObject.Properties[$name]
  if ($null -ne $prop -and $null -ne $prop.Value) { return $prop.Value }
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
     description = 'Optional settings. Not needed to connect: sign-in works without any IDs. allow_send enables sending mail (default false: drafts only). reset=true deletes saved IDs and sign-in and returns to the default (no IDs); use it when sign-in fails with AADSTS700016 or after IDs were entered by mistake. tenant_id and client_id only if the IT admin registered a dedicated app in Entra and asked for it.'
     props = @{ tenant_id = @{ type = 'string' }; client_id = @{ type = 'string' }; allow_send = @{ type = 'boolean' }; reset = @{ type = 'boolean' } }; required = @() },
  @{ name = 'ms_check'; title = 'בדיקת החיבור ל-Microsoft'; readOnly = $true
     description = 'Diagnoses this computer before or after sign-in: PowerShell version and language mode, execution policy, secure storage, access to Microsoft sign-in and Graph, and the sign-in state. Run it first when setting up or when anything fails.'
     props = @{}; required = @() },
  @{ name = 'ms_admin_link'; title = 'קישור אישור למנהל'; readOnly = $true
     description = 'Returns the one-time admin consent link for the IT administrator, with the app name and permissions, ready to forward. Use when the user asks for the link for his IT person or when sign-in needs admin approval.'
     props = @{}; required = @() },
  @{ name = 'ms_sign_out'; title = 'ניתוק מחשבון Microsoft'; readOnly = $false
     description = 'Disconnects: deletes the stored sign-in from this computer. Use it immediately if ms_status shows someone other than the consultant (for example the IT admin who signed in to approve), then sign in again with the consultant own account.'
     props = @{}; required = @() },
  @{ name = 'ms_sign_in'; title = 'כניסה לחשבון Microsoft'; readOnly = $false
     description = 'Starts sign-in. method "code" (default): returns a web address and a short code; the user opens it, enters the code and signs in with his own work account. method "browser": opens the sign-in page in the default browser and returns the address in case it did not open; use it when the code method is blocked or fails. Either way, then call ms_sign_in_complete.'
     props = @{ method = @{ type = 'string'; enum = @('code', 'browser') } }; required = @() },
  @{ name = 'ms_sign_in_complete'; title = 'סיום כניסה'; readOnly = $false
     description = 'Waits (up to about 45 seconds) for the user to finish the sign-in started by ms_sign_in and stores the credentials securely on this computer.'
     props = @{}; required = @() },
  @{ name = 'mail_search'; title = 'חיפוש מיילים'; readOnly = $true
     description = 'Searches the signed-in mailbox (all folders, including sent mail). Use from (sender address or domain, e.g. a client domain from the client index), to (recipient address or domain, for mail the user sent), query (free text: words from the subject, body or attachment names), since/until (YYYY-MM-DD). Returns newest first.'
     props = @{ query = @{ type = 'string' }; from = @{ type = 'string' }; to = @{ type = 'string' }; since = @{ type = 'string' }; until = @{ type = 'string' }; top = @{ type = 'integer' } }; required = @() },
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
     description = 'Creates a draft in the user''s Outlook Drafts folder (nothing is sent). This is the default for any request to write an email. For a reply pass reply_to_message_id; the text is placed above the quoted message. Body is plain text; Hebrew is fine. To change it afterwards use mail_update_draft, not a second draft.'
     props = @{ to = @{ type = 'array'; items = @{ type = 'string' } }; cc = @{ type = 'array'; items = @{ type = 'string' } }; subject = @{ type = 'string' }; body = @{ type = 'string' }; reply_to_message_id = @{ type = 'string' } }; required = @('body') },
  @{ name = 'mail_update_draft'; title = 'עדכון טיוטה'; readOnly = $false
     description = 'Changes a draft created earlier (subject, recipients, or the text; a reply keeps its quoted message). Use it for corrections instead of creating another draft. Nothing is sent.'
     props = @{ draft_id = @{ type = 'string' }; to = @{ type = 'array'; items = @{ type = 'string' } }; cc = @{ type = 'array'; items = @{ type = 'string' } }; subject = @{ type = 'string' }; body = @{ type = 'string' } }; required = @('draft_id') },
  @{ name = 'mail_send_draft'; title = 'שליחת טיוטה'; readOnly = $false; destructive = $true
     description = 'Sends an existing draft. Works only if sending was enabled in ms_configure (allow_send) and only with confirmed=true after the user approved sending this exact draft.'
     props = @{ draft_id = @{ type = 'string' }; confirmed = @{ type = 'boolean' } }; required = @('draft_id', 'confirmed') },
  @{ name = 'calendar_events'; title = 'אירועים ביומן'; readOnly = $true
     description = 'Lists calendar events between start and end (YYYY-MM-DD or YYYY-MM-DDTHH:MM, Israel time). Optional query keeps events whose subject, location, organizer or attendee names and addresses contain the text (a client name, a person, or a client email domain). Use it to answer what is scheduled, when the next or last meeting with someone is, and when the user is free.'
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
  @{ name = 'contacts_search'; title = 'חיפוש איש קשר'; readOnly = $true
     description = 'Searches the user''s Outlook contacts (default contacts folder) by any part of a name, email address, phone number or company. Use it before contacts_create to avoid duplicates, and to find a contact_id for contacts_update.'
     props = @{ query = @{ type = 'string' }; top = @{ type = 'integer' } }; required = @('query') },
  @{ name = 'contacts_create'; title = 'איש קשר חדש'; readOnly = $false
     description = 'Creates a contact in the user''s Outlook contacts. Nothing is sent to anyone. Returns the saved contact with its contact_id; only this result proves the contact exists. display_name defaults to given_name and surname.'
     props = @{ given_name = @{ type = 'string' }; surname = @{ type = 'string' }; display_name = @{ type = 'string' }; emails = @{ type = 'array'; items = @{ type = 'string' } }; mobile_phone = @{ type = 'string' }; business_phone = @{ type = 'string' }; company = @{ type = 'string' }; job_title = @{ type = 'string' }; notes = @{ type = 'string' } }; required = @() },
  @{ name = 'contacts_update'; title = 'עדכון איש קשר'; readOnly = $false
     description = 'Changes an existing Outlook contact. Only the fields passed change. emails are added to the existing addresses (none are removed); mobile_phone and business_phone replace the current number; notes replaces the current notes.'
     props = @{ contact_id = @{ type = 'string' }; given_name = @{ type = 'string' }; surname = @{ type = 'string' }; display_name = @{ type = 'string' }; emails = @{ type = 'array'; items = @{ type = 'string' } }; mobile_phone = @{ type = 'string' }; business_phone = @{ type = 'string' }; company = @{ type = 'string' }; job_title = @{ type = 'string' }; notes = @{ type = 'string' } }; required = @('contact_id') },
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
  $custom = ($cfg.client_id -ne $DefaultClient) -or ($cfg.tenant_id -ne $DefaultTenant)
  $note = if ($custom) { " שים לב: שמורים במחשב מזהים (ארגון $($cfg.tenant_id), אפליקציה $($cfg.client_id)). אם איש המחשבים לא ביקש אותם במפורש, להחזיר לברירת המחדל: ms_configure עם reset=true." } else { '' }
  if (-not (Test-Path $TokenPath)) { return "עוד לא נכנסו לחשבון Microsoft. לא צריך מזהים: הצעד הבא הוא ms_sign_in.$note" }
  $me = Invoke-Graph 'GET' '/me?$select=displayName,mail,userPrincipalName'
  $send = if ($cfg.allow_send) { 'מופעלת (רק עם אישור לכל מייל)' } else { 'כבויה: טיוטות בלבד' }
  $contacts = if ($script:ContactsAllowed) { 'מאושרים' } else { 'עוד לא אושרו: צריך להיכנס שוב פעם אחת (ms_sign_in) כדי לעבוד עם אנשי קשר. מייל ויומן עובדים גם בלי זה' }
  return "מחובר בשם $($me.displayName) ($($me.mail), כניסה: $($me.userPrincipalName)). שליחת מייל: $send. אנשי קשר: $contacts. אם זה לא החשבון של בעל התיבה, להפעיל ms_sign_out ולהתחבר מחדש."
}

function Tool-Configure($a) {
  if (Get-Prop $a 'reset' $false) {
    $allowKeep = [bool](Get-Config).allow_send
    Tool-SignOut $a | Out-Null
    Remove-Item $ConfigPath -Force -ErrorAction SilentlyContinue
    if ($allowKeep) { Save-Config $DefaultTenant $DefaultClient $true }
    return 'חזר לברירת המחדל: אפליקציה ציבורית של Microsoft, בלי מזהים. הצעד הבא: ms_sign_in.'
  }
  $cfg = Get-Config
  $tenant = Get-Prop $a 'tenant_id' $cfg.tenant_id
  $client = Get-Prop $a 'client_id' $cfg.client_id
  if (-not $tenant) { $tenant = $DefaultTenant }
  if (-not $client) { $client = $DefaultClient }
  if ($tenant -notmatch '^([0-9a-fA-F-]{36}|organizations|[A-Za-z0-9.-]+\.[A-Za-z]{2,})$' -or $client -notmatch '^[0-9a-fA-F-]{36}$') {
    throw 'tenant_id: מזהה GUID או דומיין של הארגון. client_id: מזהה GUID. אפשר גם בלי שניהם.'
  }
  $allow = Get-Prop $a 'allow_send' $cfg.allow_send
  $changed = ($tenant -ne $cfg.tenant_id) -or ($client -ne $cfg.client_id) -or ([bool]$allow -ne [bool]$cfg.allow_send)
  Save-Config $tenant $client $allow
  if ($changed -and (Test-Path $TokenPath)) { Remove-Item $TokenPath -Force; $script:AccessToken = $null; Write-AuthLog 'ההגדרות שונו, נדרשת כניסה מחדש'; return 'נשמר. צריך להיכנס שוב: ms_sign_in.' }
  if (Test-Path $TokenPath) { return 'נשמר.' }
  return 'נשמר. הצעד הבא: ms_sign_in.'
}

function New-UrlSafeRandom([int]$bytes) {
  $b = New-Object byte[] $bytes
  $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
  $rng.GetBytes($b)
  return [Convert]::ToBase64String($b).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Stop-BrowserSignIn {
  if ($script:Browser) { foreach ($l in $script:Browser.listeners) { try { $l.Stop() } catch {} } }
  $script:Browser = $null
}

function Start-BrowserSignIn($cfg) {
  Stop-BrowserSignIn
  # Loopback redirect on 127.0.0.1 and ::1 only, so Windows Firewall never asks and nothing is reachable from outside.
  $first = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
  $first.Start()
  $port = ([Net.IPEndPoint]$first.LocalEndpoint).Port
  $listeners = @($first)
  try { $v6 = New-Object Net.Sockets.TcpListener([Net.IPAddress]::IPv6Loopback, $port); $v6.Start(); $listeners += $v6 } catch {}
  $verifier = New-UrlSafeRandom 48
  $sha = [Security.Cryptography.SHA256]::Create()
  $challenge = [Convert]::ToBase64String($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))).TrimEnd('=').Replace('+', '-').Replace('/', '_')
  $state = New-UrlSafeRandom 16
  $redirect = "http://localhost:$port"
  $url = "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/authorize?client_id=$($cfg.client_id)&response_type=code&redirect_uri=$(Q $redirect)&response_mode=query&scope=$(Q (Get-Scopes $cfg))&code_challenge=$challenge&code_challenge_method=S256&state=$state&prompt=select_account"
  $script:Browser = @{ listeners = $listeners; verifier = $verifier; state = $state; redirect = $redirect; expires = [DateTime]::UtcNow.AddMinutes(10) }
  $opened = $true
  if ($env:CONSULTING_OFFICE_NO_BROWSER) { $opened = $false } else { try { Start-Process $url } catch { $opened = $false } }
  $lead = if ($opened) { 'נפתח דף כניסה בדפדפן.' } else { 'הדפדפן לא נפתח לבד.' }
  return "$lead אם הוא לא מופיע, פותחים את הכתובת הזו:`n$url`nנכנסים עם החשבון של Microsoft 365, ומאשרים אם מופיע מסך הרשאות. אחר כך מפעילים ms_sign_in_complete."
}

function Read-BrowserReply($client) {
  $stream = $client.GetStream()
  $stream.ReadTimeout = 5000
  $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::ASCII)
  $first = $reader.ReadLine()
  $query = @{}
  if ($first -match '^GET\s+(\S+)') {
    $target = $Matches[1]
    $qs = if ($target.Contains('?')) { $target.Substring($target.IndexOf('?') + 1) } else { '' }
    foreach ($pair in ($qs -split '&')) {
      if (-not $pair) { continue }
      $kv = $pair -split '=', 2
      $val = if ($kv.Count -gt 1) { [Uri]::UnescapeDataString($kv[1].Replace('+', ' ')) } else { '' }
      $query[[Uri]::UnescapeDataString($kv[0])] = $val
    }
  }
  $isAnswer = $query.ContainsKey('code') -or $query.ContainsKey('error')
  $html = if ($query.ContainsKey('code')) { 'הכניסה הצליחה. אפשר לסגור את החלון ולחזור לעוזר.' } elseif ($query.ContainsKey('error')) { 'הכניסה לא הושלמה. אפשר לסגור את החלון ולחזור לעוזר.' } else { '' }
  $body = $Utf8.GetBytes("<!doctype html><html dir=`"rtl`" lang=`"he`"><meta charset=`"utf-8`"><body style=`"font-family:sans-serif;padding:40px;font-size:20px`">$html</body></html>")
  $status = if ($isAnswer) { '200 OK' } else { '404 Not Found' }
  $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $status`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
  try { $stream.Write($head, 0, $head.Length); $stream.Write($body, 0, $body.Length); $stream.Flush() } catch {}
  $client.Close()
  if ($isAnswer) { return $query }
  return $null
}

function Complete-BrowserSignIn($cfg) {
  $b = $script:Browser
  if ([DateTime]::UtcNow -gt $b.expires) { Stop-BrowserSignIn; throw 'זמן הכניסה עבר. הפעל ms_sign_in שוב.' }
  $deadline = [DateTime]::UtcNow.AddSeconds(45)
  while ([DateTime]::UtcNow -lt $deadline) {
    foreach ($l in $b.listeners) {
      if (-not $l.Pending()) { continue }
      $reply = $null
      try { $reply = Read-BrowserReply ($l.AcceptTcpClient()) } catch { Log "browser reply: $($_.Exception.Message)" }
      if ($null -eq $reply) { continue }
      if ($reply.ContainsKey('error')) { Stop-BrowserSignIn; throw ("הכניסה נכשלה: " + (Explain-AuthError $reply['error'] $reply['error_description'])) }
      if ($reply['state'] -ne $b.state) { Stop-BrowserSignIn; throw 'הכניסה נכשלה: תשובה לא תואמת מהדפדפן. הפעל ms_sign_in שוב.' }
      $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
        client_id = $cfg.client_id; grant_type = 'authorization_code'; code = $reply['code']; redirect_uri = $b.redirect
        code_verifier = $b.verifier; scope = (Get-Scopes $cfg)
      }
      Stop-BrowserSignIn
      if (-not ($null -ne $r.PSObject.Properties['access_token'])) { throw ("הכניסה נכשלה: " + (Explain-AuthError (Get-Prop $r 'error' '') (Get-Prop $r 'error_description' ''))) }
      return $r
    }
    Start-Sleep -Milliseconds 500
  }
  return $null
}

function Save-SignIn($r) {
  Save-RefreshToken $r.refresh_token
  Write-AuthLog 'נכנס לחשבון'
  $script:ContactsAllowed = $true
  $script:AccessToken = $r.access_token
  $script:AccessExpires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in - 120)
}

function Tool-SignIn($a) {
  $cfg = Get-Config
  Remove-Item $PendingPath -Force -ErrorAction SilentlyContinue
  if ((Get-Prop $a 'method' 'code') -eq 'browser') { return (Start-BrowserSignIn $cfg) }
  Stop-BrowserSignIn
  $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/devicecode" @{ client_id = $cfg.client_id; scope = (Get-Scopes $cfg) }
  if (-not ($null -ne $r.PSObject.Properties['device_code'])) { throw ("הכניסה לא התחילה: " + (Explain-AuthError (Get-Prop $r 'error' '') (Get-Prop $r 'error_description' '')) + " אפשר לנסות ms_sign_in עם method=browser.") }
  $pending = [ordered]@{ device_code = $r.device_code; interval = [int]$r.interval; expires = [DateTime]::UtcNow.AddSeconds([int]$r.expires_in).ToString('o') }
  New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
  [IO.File]::WriteAllText($PendingPath, ($pending | ConvertTo-Json), $Utf8)
  return "פתח בדפדפן: $($r.verification_uri)`nהזן את הקוד: $($r.user_code)`nוהתחבר עם החשבון של Microsoft 365. אחרי שזה הסתיים, הפעל ms_sign_in_complete."
}

function Tool-SignInComplete($a) {
  $cfg = Get-Config
  if ($script:Browser) {
    $r = Complete-BrowserSignIn $cfg
    if ($null -eq $r) { return 'הכניסה עוד לא הושלמה. אחרי שמסיימים בדפדפן, הפעל ms_sign_in_complete שוב.' }
    Save-SignIn $r
    return (Tool-Status $a)
  }
  if (-not (Test-Path $PendingPath)) { throw 'אין כניסה פתוחה. הפעל ms_sign_in קודם.' }
  $p = Get-Content -Raw -Encoding UTF8 $PendingPath | ConvertFrom-Json
  # Stay under the host's per-tool timeout; the skill simply calls again if sign-in is still pending.
  $deadline = [DateTime]::UtcNow.AddSeconds(45)
  $interval = [Math]::Max(5, [int]$p.interval)
  while ([DateTime]::UtcNow -lt $deadline) {
    $r = Invoke-Form "https://login.microsoftonline.com/$($cfg.tenant_id)/oauth2/v2.0/token" @{
      client_id = $cfg.client_id; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $p.device_code
    }
    if ($null -ne $r.PSObject.Properties['access_token']) {
      Remove-Item $PendingPath -Force
      Save-SignIn $r
      return (Tool-Status $a)
    }
    switch (Get-Prop $r 'error' '') {
      'authorization_pending' { Start-Sleep -Seconds $interval }
      'slow_down' { $interval += 5; Start-Sleep -Seconds $interval }
      default { Remove-Item $PendingPath -Force -ErrorAction SilentlyContinue; throw ("הכניסה נכשלה: " + (Explain-AuthError (Get-Prop $r 'error' '') (Get-Prop $r 'error_description' '')) + " אפשר לנסות ms_sign_in עם method=browser.") }
    }
  }
  return 'הכניסה עוד לא הושלמה. אחרי שמסיימים בדפדפן, הפעל ms_sign_in_complete שוב.'
}

function Tool-SignOut($a) {
  Stop-BrowserSignIn
  if (Test-Path $TokenPath) { Write-AuthLog 'נותק (ms_sign_out)' }
  Remove-Item $TokenPath -Force -ErrorAction SilentlyContinue
  Remove-Item $PendingPath -Force -ErrorAction SilentlyContinue
  $script:AccessToken = $null
  $script:AccessExpires = [DateTime]::MinValue
  return 'נותק. הכניסה השמורה נמחקה מהמחשב. להתחברות מחדש: ms_sign_in.'
}

function Tool-AdminLink($a) {
  $perUser = @'
Connect-MgGraph -Scopes "Application.ReadWrite.All","DelegatedPermissionGrant.ReadWrite.All","User.Read.All"
$app = Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'"
if (-not $app) { $app = New-MgServicePrincipal -AppId '14d82eec-204b-4c2f-b7e8-296a70dab67e' }
$graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$user = Get-MgUser -UserId '<the address the user signs in to Outlook with>'
New-MgOauth2PermissionGrant -BodyParameter @{ clientId = $app.Id; consentType = 'Principal'; principalId = $user.Id; resourceId = $graph.Id; scope = 'offline_access User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite Files.Read.All' }
'@
  return "האפליקציה: Microsoft Graph Command Line Tools ($DefaultClient)`nההרשאות (delegated): User.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Contacts.ReadWrite, Files.Read.All`nהחשבון: זה שהמשתמש נכנס איתו ל-Outlook (כתובת הכניסה). כתובת השולח העסקית היא בדרך כלל כתובת של אותה תיבה.`n`nאפשרות א, לכל הארגון: המנהל פותח את הקישור, נכנס עם חשבון מנהל ולוחץ Accept (דף ריק אחרי האישור תקין):`n$(Get-AdminConsentUrl)`n`nאפשרות ב, רק לחשבון אחד (Microsoft Graph PowerShell, מנהל):`n" + $perUser
}

function Tool-Check($a) {
  $lines = @()
  $lines += "PowerShell $($PSVersionTable.PSVersion) · מצב שפה: $($ExecutionContext.SessionState.LanguageMode)"
  try { $pol = (Get-ExecutionPolicy -List | Where-Object { $_.ExecutionPolicy -ne 'Undefined' } | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ', '; $lines += "מדיניות הרצה: $pol" } catch { $lines += 'מדיניות הרצה: לא ידוע' }
  try {
    New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
    $probe = ConvertTo-SecureString -String 'probe' -AsPlainText -Force | ConvertFrom-SecureString
    $back = $probe | ConvertTo-SecureString
    $lines += "שמירה מוצפנת: תקין ($DataDir)"
  } catch { $lines += "שמירה מוצפנת: נכשל: $($_.Exception.Message)" }
  foreach ($t in @(@('כניסה של Microsoft', 'https://login.microsoftonline.com/organizations/v2.0/.well-known/openid-configuration'), @('Microsoft Graph', 'https://graph.microsoft.com/v1.0/$metadata'))) {
    try { Invoke-WebRequest -Uri $t[1] -UseBasicParsing -TimeoutSec 15 -Method Head | Out-Null; $lines += "$($t[0]): נגיש" }
    catch {
      $code = $null
      try { $code = [int]$_.Exception.Response.StatusCode } catch {}
      if ($code) { $lines += "$($t[0]): נגיש (HTTP $code)" } else { $lines += "$($t[0]): לא נגיש: $($_.Exception.Message)" }
    }
  }
  $cfg = Get-Config
  $mode = if ((Get-Prop $cfg 'default' $false) -or $cfg.client_id -eq $DefaultClient) { 'אפליקציה ציבורית של Microsoft, בלי מזהים' } else { "אפליקציה ייעודית ($($cfg.client_id))" }
  $lines += "חיבור: $mode · שליחה: $(if ($cfg.allow_send) { 'מופעלת' } else { 'טיוטות בלבד' })"
  try { $lines += "מצב: $(Tool-Status $a)" } catch { $lines += "מצב: $($_.Exception.Message)" }
  if (Test-Path $LogPath) { $lines += 'יומן חיבור (אחרונים):'; $lines += @(Get-Content -Encoding UTF8 $LogPath | Select-Object -Last 10) }
  return ($lines -join "`n")
}

function Tool-MailSearch($a) {
  $top = [Math]::Min([int](Get-Prop $a 'top' 15), 50)
  $sel = '$select=id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink'
  $query = Get-Prop $a 'query' ''
  $from = Get-Prop $a 'from' ''
  $to = Get-Prop $a 'to' ''
  $since = Get-Prop $a 'since' ''
  $until = Get-Prop $a 'until' ''
  if ($query -or $to -or ($from -and -not $from.Contains('@'))) {
    # $search cannot be combined with $filter/$orderby, so sender, recipient, domain and dates go into the KQL query.
    $kql = $query
    if ($from) { $kql = ("$kql from:" + $from.TrimStart('@')).Trim() }
    if ($to) { $kql = ("$kql to:" + $to.TrimStart('@')).Trim() }
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
  $m = Invoke-Graph 'GET' "/me/messages/$($a.message_id)?`$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,webLink" -Prefer 'outlook.body-content-type="text"'
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
  if (-not ($null -ne $att.PSObject.Properties['contentBytes'])) { throw 'זה לא קובץ מצורף רגיל (למשל פריט Outlook מקושר) ולא ניתן לשמור אותו.' }
  $name = ($att.name -replace '[\\/:*?"<>|]', '_')
  $target = Join-Path $a.folder $name
  $i = 2
  while (Test-Path -LiteralPath $target) {
    $target = Join-Path $a.folder ("{0} ({1}){2}" -f [IO.Path]::GetFileNameWithoutExtension($name), $i, [IO.Path]::GetExtension($name)); $i++
  }
  [IO.File]::WriteAllBytes($target, [Convert]::FromBase64String($att.contentBytes))
  return "נשמר: $target"
}

function Format-DraftHtml($text) {
  return '<div dir="rtl" style="font-family:Arial">' + [Net.WebUtility]::HtmlEncode([string]$text).Replace("`n", '<br>') + '</div>'
}

function Tool-MailCreateDraft($a) {
  $body = [string]$a.body
  $replyTo = Get-Prop $a 'reply_to_message_id' ''
  if ($replyTo) {
    $draft = Invoke-Graph 'POST' "/me/messages/$replyTo/createReply" @{}
    $html = (Format-DraftHtml $body) + '<br>' + $draft.body.content
    $patch = @{ body = @{ contentType = 'HTML'; content = $html } }
    $to = Get-Prop $a 'to' $null
    if ($to) { $patch['toRecipients'] = To-Recipients $to }
    $cc = Get-Prop $a 'cc' $null
    if ($cc) { $patch['ccRecipients'] = To-Recipients $cc }
    $draft = Invoke-Graph 'PATCH' "/me/messages/$($draft.id)" $patch
  } else {
    $html = Format-DraftHtml $body
    $msg = @{ subject = [string](Get-Prop $a 'subject' ''); body = @{ contentType = 'HTML'; content = $html }; toRecipients = (To-Recipients (Get-Prop $a 'to' @())) }
    $cc = Get-Prop $a 'cc' $null
    if ($cc) { $msg['ccRecipients'] = To-Recipients $cc }
    $draft = Invoke-Graph 'POST' '/me/messages' $msg
  }
  return [ordered]@{ draft_id = $draft.id; subject = $draft.subject; web_link = (Get-Prop $draft 'webLink' ''); note = 'הטיוטה נשמרה בתיקיית הטיוטות ב-Outlook. שום דבר לא נשלח.' }
}

function Tool-MailUpdateDraft($a) {
  $id = [string]$a.draft_id
  $draft = Invoke-Graph 'GET' ("/me/messages/$id" + '?$select=isDraft,subject,body')
  if (-not $draft.isDraft) { throw 'זה לא טיוטה: המייל כבר נשלח, ואי אפשר לשנות אותו.' }
  $patch = @{}
  $subject = Get-Prop $a 'subject' $null
  if ($null -ne $subject) { $patch['subject'] = [string]$subject }
  $to = Get-Prop $a 'to' $null
  if ($to) { $patch['toRecipients'] = To-Recipients $to }
  $cc = Get-Prop $a 'cc' $null
  if ($cc) { $patch['ccRecipients'] = To-Recipients $cc }
  $body = Get-Prop $a 'body' $null
  if ($null -ne $body) {
    # A reply keeps the quoted message, which Outlook places after a horizontal line (divRplyFwdMsg).
    $html = Format-DraftHtml $body
    $old = [string]$draft.body.content
    $quote = [regex]::Match($old, '(?is)<hr[^>]*>|<div[^>]*divRplyFwdMsg')
    if ($quote.Success) { $html += '<br>' + $old.Substring($quote.Index) }
    $patch['body'] = @{ contentType = 'HTML'; content = $html }
  }
  if ($patch.Count -eq 0) { throw 'לא נמסר מה לשנות בטיוטה.' }
  $draft = Invoke-Graph 'PATCH' "/me/messages/$id" $patch
  return [ordered]@{ draft_id = $draft.id; subject = $draft.subject; web_link = (Get-Prop $draft 'webLink' ''); note = 'הטיוטה עודכנה ב-Outlook. שום דבר לא נשלח.' }
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
  $r = Invoke-Graph 'GET' $url -Prefer "outlook.timezone=`"$TimeZone`""
  $events = @($r.value)
  $query = Get-Prop $a 'query' ''
  if ($query) {
    # Not only the subject: a client meeting is often titled with a person's name, so attendees and location count too.
    $events = @($events | Where-Object {
      $people = @((Get-Prop $_ 'organizer' $null)) + @(Get-Prop $_ 'attendees' @())
      $text = @((Get-Prop $_ 'subject' ''), (Get-Prop (Get-Prop $_ 'location' $null) 'displayName' '')) + @($people | ForEach-Object {
        $ea = Get-Prop $_ 'emailAddress' $null
        (Get-Prop $ea 'name' ''); (Get-Prop $ea 'address' '')
      })
      ($text -join ' ').IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0
    })
  }
  return @($events | ForEach-Object { Format-Event $_ })
}

function Tool-CalendarEventRead($a) {
  $e = Invoke-Graph 'GET' "/me/events/$($a.event_id)?`$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,bodyPreview,webLink,body" -Prefer "outlook.timezone=`"$TimeZone`", outlook.body-content-type=`"text`""
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
  $e = Invoke-Graph 'POST' '/me/events' $event -Prefer "outlook.timezone=`"$TimeZone`""
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
  $e = Invoke-Graph 'PATCH' "/me/events/$($a.event_id)" $patch -Prefer "outlook.timezone=`"$TimeZone`""
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

$ContactFields = 'id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,homePhones,companyName,jobTitle,personalNotes'
$ContactsNotAllowed = 'CONTACTS_NOT_ALLOWED: החיבור נוצר לפני שנוספו אנשי קשר, ולכן עוד לא אושרה גישה אליהם. צריך להיכנס שוב פעם אחת (ms_sign_in ואז ms_sign_in_complete) ולאשר את ההרשאה החדשה. מייל ויומן ממשיכים לעבוד בינתיים. אם מופיע "צריך אישור מנהל": ms_admin_link.'

function Invoke-Contacts($method, $path, $body = $null) {
  Get-AccessToken | Out-Null
  if (-not $script:ContactsAllowed) { throw $ContactsNotAllowed }
  try { return (Invoke-Graph $method $path $body) }
  catch {
    if ("$($_.Exception.Message)" -match 'ErrorAccessDenied|Access is denied|403') { throw $ContactsNotAllowed }
    throw
  }
}

function Format-Contact($c) {
  $phones = @()
  $mobile = Get-Prop $c 'mobilePhone' ''
  if ($mobile) { $phones += "נייד: $mobile" }
  foreach ($n in @(Get-Prop $c 'businessPhones' @())) { if ($n) { $phones += "עבודה: $n" } }
  foreach ($n in @(Get-Prop $c 'homePhones' @())) { if ($n) { $phones += "בית: $n" } }
  return [ordered]@{
    contact_id = $c.id; display_name = (Get-Prop $c 'displayName' ''); given_name = (Get-Prop $c 'givenName' ''); surname = (Get-Prop $c 'surname' '')
    emails = @(@(Get-Prop $c 'emailAddresses' @()) | ForEach-Object { Get-Prop $_ 'address' '' } | Where-Object { $_ })
    phones = $phones; company = (Get-Prop $c 'companyName' ''); job_title = (Get-Prop $c 'jobTitle' ''); notes = (Get-Prop $c 'personalNotes' '')
  }
}

function To-ContactEmails($list) {
  return @(@($list) | Where-Object { $_ } | ForEach-Object { $addr = ([string]$_).Trim(); @{ address = $addr; name = $addr } })
}

function Get-ContactPatch($a) {
  $patch = @{}
  $map = [ordered]@{ given_name = 'givenName'; surname = 'surname'; display_name = 'displayName'; mobile_phone = 'mobilePhone'; company = 'companyName'; job_title = 'jobTitle'; notes = 'personalNotes' }
  foreach ($k in $map.Keys) { $v = Get-Prop $a $k $null; if ($null -ne $v) { $patch[$map[$k]] = ([string]$v).Trim() } }
  $business = Get-Prop $a 'business_phone' $null
  if ($null -ne $business) { $patch['businessPhones'] = @(@(([string]$business).Trim()) | Where-Object { $_ }) }
  return $patch
}

function Tool-ContactsSearch($a) {
  $query = ([string]$a.query).Trim()
  if (-not $query) { throw 'לא נמסר מה לחפש.' }
  $top = [Math]::Min([int](Get-Prop $a 'top' 10), 50)
  $digits = $query -replace '\D', ''
  # Graph has no substring search on contacts, so read the folder (up to 5,000) and match here: any part of a name, address, phone or company.
  $found = @()
  $url = "/me/contacts?`$top=500&`$select=$ContactFields"
  $pages = 0
  while ($url -and $pages -lt 10) {
    $r = Invoke-Contacts 'GET' $url
    foreach ($c in @($r.value)) {
      $text = (@((Get-Prop $c 'displayName' ''), (Get-Prop $c 'givenName' ''), (Get-Prop $c 'surname' ''), (Get-Prop $c 'companyName' '')) + @(@(Get-Prop $c 'emailAddresses' @()) | ForEach-Object { Get-Prop $_ 'address' '' })) -join ' '
      $hit = $text.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0
      if (-not $hit -and $digits.Length -ge 6) {
        $numbers = @((Get-Prop $c 'mobilePhone' '')) + @(Get-Prop $c 'businessPhones' @()) + @(Get-Prop $c 'homePhones' @())
        foreach ($n in $numbers) { if ($n -and (([string]$n -replace '\D', '').Contains($digits.TrimStart('0')))) { $hit = $true } }
      }
      if ($hit) { $found += (Format-Contact $c) }
    }
    $url = Get-Prop $r '@odata.nextLink' $null
    $pages++
  }
  if (-not $found.Count) { return "לא נמצא איש קשר שמתאים ל-'$query'." }
  return @($found | Select-Object -First $top)
}

function Tool-ContactsCreate($a) {
  $contact = Get-ContactPatch $a
  # @() around the call: a function returning an empty or one-item array hands back $null or the item itself.
  $emails = @(To-ContactEmails (Get-Prop $a 'emails' @()))
  if ($emails.Count) { $contact['emailAddresses'] = $emails }
  if (-not (Get-Prop $contact 'displayName' '')) { $contact['displayName'] = ("$(Get-Prop $contact 'givenName' '') $(Get-Prop $contact 'surname' '')").Trim() }
  if (-not $contact['displayName'] -and $emails.Count) { $contact['displayName'] = $emails[0].address }
  if (-not $contact['displayName']) { throw 'חסר שם או כתובת מייל לאיש הקשר.' }
  $c = Invoke-Contacts 'POST' '/me/contacts' $contact
  $out = Format-Contact $c
  $out['note'] = 'איש הקשר נשמר באנשי הקשר ב-Outlook.'
  return $out
}

function Tool-ContactsUpdate($a) {
  $id = [string]$a.contact_id
  $patch = Get-ContactPatch $a
  $add = @(Get-Prop $a 'emails' @())
  if ($add.Count) {
    $current = Invoke-Contacts 'GET' ("/me/contacts/$id" + '?$select=emailAddresses')
    $existing = @(@(Get-Prop $current 'emailAddresses' @()) | ForEach-Object { Get-Prop $_ 'address' '' } | Where-Object { $_ })
    $merged = @($existing)
    foreach ($e in $add) { $e = ([string]$e).Trim(); if ($e -and -not (@($merged | Where-Object { $_ -ieq $e }).Count)) { $merged += $e } }
    $patch['emailAddresses'] = @(To-ContactEmails $merged)
  }
  if ($patch.Count -eq 0) { throw 'לא נמסר מה לשנות באיש הקשר.' }
  $c = Invoke-Contacts 'PATCH' "/me/contacts/$id" $patch
  $out = Format-Contact $c
  $out['note'] = 'איש הקשר עודכן ב-Outlook.'
  return $out
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
  $r = Invoke-Graph 'GET' "${url}?`$top=200&`$select=name,webUrl,lastModifiedDateTime,size,file,folder"
  return @($r.value | ForEach-Object { [ordered]@{ name = $_.name; modified = $_.lastModifiedDateTime; size = $_.size; is_folder = [bool](Get-Prop $_ 'folder' $null); web_url = $_.webUrl } })
}

$Handlers = @{
  ms_status = ${function:Tool-Status}; ms_configure = ${function:Tool-Configure}; ms_check = ${function:Tool-Check}; ms_admin_link = ${function:Tool-AdminLink}; ms_sign_out = ${function:Tool-SignOut}
  ms_sign_in = ${function:Tool-SignIn}; ms_sign_in_complete = ${function:Tool-SignInComplete}
  mail_search = ${function:Tool-MailSearch}; mail_read = ${function:Tool-MailRead}
  mail_attachments = ${function:Tool-MailAttachments}; mail_save_attachment = ${function:Tool-MailSaveAttachment}
  mail_create_draft = ${function:Tool-MailCreateDraft}; mail_update_draft = ${function:Tool-MailUpdateDraft}; mail_send_draft = ${function:Tool-MailSendDraft}
  calendar_events = ${function:Tool-CalendarEvents}; calendar_event_read = ${function:Tool-CalendarEventRead}
  calendar_create_event = ${function:Tool-CalendarCreate}
  calendar_update_event = ${function:Tool-CalendarUpdate}; calendar_cancel_event = ${function:Tool-CalendarCancel}
  contacts_search = ${function:Tool-ContactsSearch}; contacts_create = ${function:Tool-ContactsCreate}; contacts_update = ${function:Tool-ContactsUpdate}
  onedrive_search = ${function:Tool-OneDriveSearch}; onedrive_list = ${function:Tool-OneDriveList}
}

# ---------- MCP over stdio ----------

function Send-Message($obj) {
  $StdOut.WriteLine((ConvertTo-Json -InputObject $obj -Depth 30 -Compress))
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
  $line = $StdIn.ReadLine()
  if ($null -eq $line) { break }
  if (-not $line.Trim()) { continue }
  try { $msg = $line | ConvertFrom-Json } catch { Log "bad json: $line"; continue }
  $hasId = $null -ne $msg.PSObject.Properties['id']
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
