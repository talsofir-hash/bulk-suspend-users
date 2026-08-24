#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk-suspend (or restore) Atlassian cloud organization users listed in a CSV.

.DESCRIPTION
    Reads a CSV that contains one column of email addresses, resolves each email
    to an Atlassian accountId + directoryId through the Organizations REST API,
    then suspends (or restores) each user's access in the organization.

    Suspending keeps the account, its product roles and its group memberships,
    so the action is reversible with -Mode Restore.

    Endpoints used (https://developer.atlassian.com/cloud/admin/organization/):
      GET  /v2/orgs/{orgId}/directories
      POST /v2/orgs/{orgId}/directories/{directoryId}/users/search
      POST /v2/orgs/{orgId}/directories/{directoryId}/users/{accountId}/suspend
      POST /v2/orgs/{orgId}/directories/{directoryId}/users/{accountId}/restore

    SAFETY: without -Apply the script only reads. It resolves everything, prints
    exactly who would be actioned, writes a plan CSV and changes nothing.

.EXAMPLE
    .\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv
    Dry run - shows who would be suspended.

.EXAMPLE
    .\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv -Apply
    Suspends them (asks you to type the count first).

.EXAMPLE
    .\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv -Mode Restore -Apply
    Undo - restores access for the same list.

.NOTES
    Windows PowerShell 5.1 and PowerShell 7+. ASCII only.
    stdout carries the result objects; all logging goes to stderr.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Suspend', 'Restore')][string]$Mode = 'Suspend',
    [switch]$Apply,                 # without this, the run is a dry run
    [switch]$Yes,                   # skip the typed confirmation (unattended runs)
    [string]$CsvPath,               # overrides $Config.CsvPath
    [string]$OrgId,                 # overrides $Config.OrgId
    [string]$ApiKey                 # overrides $Config.ApiKey / the env var
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ===========================================================================
#  CONFIGURATION - everything you may want to change lives in this one block
# ===========================================================================
$Config = @{

    # --- Identity -----------------------------------------------------------
    OrgId               = ''  # your organization id; or set the env var below
    OrgIdEnvVar         = 'ATLASSIAN_ORG_ID'          # env var checked when OrgId is empty
    ApiKey              = ''  # inline org API key - prefer the env var, never commit a key
    ApiKeyEnvVar        = 'ATLASSIAN_ADMIN_API_KEY'   # env var checked when ApiKey is empty
    BaseUrl             = 'https://api.atlassian.com/admin'

    # --- Input --------------------------------------------------------------
    CsvPath             = '.\users.csv'               # CSV with one column of emails
    EmailColumn         = 'email'                     # column name; '' = auto-detect (headerless files work too)

    # --- Scope and guard rails ---------------------------------------------
    DirectoryIds        = @()                         # empty = every directory in the org
    ExcludeEmails       = @(                          # never actioned, whatever the CSV says
                              'you@example.com'
                          )
    SkipAlreadyInState  = $true                       # skip users already suspended (or already active on restore)
    SuspendedPattern    = 'suspend|disabled|deactivat' # regex on the directory 'status' field = already suspended
    RequireTypedConfirm = $true                       # with -Apply, make the operator type the user count

    # --- Throughput and resilience -----------------------------------------
    SearchBatchSize     = 50                          # emails per search request (also the page limit)
    ThrottleMs          = 250                         # pause between write calls, to stay under the rate limit
    MaxRetries          = 5                           # retries for 429 / 5xx / transport errors
    RetryBaseSeconds    = 2                           # exponential back-off base when there is no Retry-After
    MaxPagesPerRequest  = 200                         # runaway-pagination backstop
    SendEmptyJsonBody   = $true                       # send {} + Content-Type: application/json - a bodyless POST gets HTTP 415

    # --- Output -------------------------------------------------------------
    OutputDir           = '.\logs'                    # plan / result CSVs land here
}
# ===========================================================================
#  END OF CONFIGURATION - no need to edit below this line
# ===========================================================================


# ---------------------------------------------------------------------------
#  Logging - always stderr so stdout stays machine-readable
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO'
    )
    $stamp = (Get-Date).ToString('HH:mm:ss')
    [Console]::Error.WriteLine("[$stamp] [$Level] $Message")
}

# ---------------------------------------------------------------------------
#  Small helpers
# ---------------------------------------------------------------------------
function Get-Prop {
    # StrictMode-safe property/key read that returns $null when absent
    param($Obj, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [hashtable]) {
        if ($Obj.ContainsKey($Name)) { return $Obj[$Name] }
        return $null
    }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-PageItems {
    # Admin API pages carry their rows under 'data'; tolerate 'values' / 'results'
    param($Page)
    foreach ($name in @('data', 'values', 'results')) {
        $v = Get-Prop $Page $name
        if ($null -ne $v) { return @($v) }
    }
    return @()
}

function Get-NextLink {
    param($Page)
    $links = Get-Prop $Page 'links'
    $next = Get-Prop $links 'next'
    if ($next) { return [string]$next }
    return $null
}

function Get-NextCursor {
    # Cursor for the next page of a POST search: links.next is either a full URL
    # carrying ?cursor=... or a bare cursor token.
    param($Page)
    $next = Get-NextLink -Page $Page
    if (-not $next) { return $null }
    if ($next -match '[?&]cursor=([^&]+)') { return [uri]::UnescapeDataString($Matches[1]) }
    if ($next -match '^https?://') { return $null }
    return $next
}

function Get-NextUri {
    # Next page URI for a GET: absolute URL as-is, bare cursor grafted on
    param($Page, [Parameter(Mandatory)][string]$CurrentUri)
    $next = Get-NextLink -Page $Page
    if (-not $next) { return $null }
    if ($next -match '^https?://') { return $next }
    $base = ($CurrentUri -replace '([?&])cursor=[^&]*', '$1').TrimEnd('?', '&')
    $sep = if ($base -match '\?') { '&' } else { '?' }
    return ($base + $sep + 'cursor=' + [uri]::EscapeDataString($next))
}

function Get-ErrorDetail {
    param([Parameter(Mandatory)]$ErrorRecord)
    $text = ''
    try { if ($ErrorRecord.ErrorDetails) { $text = [string]$ErrorRecord.ErrorDetails.Message } } catch { }
    if (-not $text) {
        try {
            $stream = $ErrorRecord.Exception.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $text = $reader.ReadToEnd()
                $reader.Dispose()
            }
        } catch { }
    }
    if ($text) {
        try {
            $obj = $text | ConvertFrom-Json
            $parts = @()
            foreach ($n in @('detail', 'message', 'title', 'error', 'error_description')) {
                $v = Get-Prop $obj $n
                if ($v) { $parts += [string]$v }
            }
            $errs = Get-Prop $obj 'errors'
            if ($errs) {
                foreach ($e in @($errs)) {
                    foreach ($n in @('detail', 'message', 'title')) {
                        $v = Get-Prop $e $n
                        if ($v) { $parts += [string]$v }
                    }
                }
            }
            if ($parts.Count -gt 0) { $text = (($parts | Select-Object -Unique) -join '; ') }
        } catch { }
    }
    if (-not $text) { $text = [string]$ErrorRecord.Exception.Message }
    return ($text -replace '\s+', ' ').Trim()
}

function Get-RetryAfterSeconds {
    param([Parameter(Mandatory)]$ErrorRecord)
    try {
        $h = $ErrorRecord.Exception.Response.Headers
        if ($null -eq $h) { return 0 }
        $raw = $null
        try { $raw = $h['Retry-After'] } catch { }
        if (-not $raw) {
            try { if ($h.RetryAfter -and $h.RetryAfter.Delta) { $raw = $h.RetryAfter.Delta.TotalSeconds } } catch { }
        }
        if ($raw) {
            $n = 0.0
            if ([double]::TryParse([string]$raw, [ref]$n) -and $n -gt 0) { return [math]::Min($n, 120) }
        }
    } catch { }
    return 0
}

function Protect-Secret {
    # Keep the API key out of every log line and error message
    param([string]$Text)
    if (-not $Text) { return $Text }
    if (-not $script:ApiKeyResolved) { return $Text }
    return $Text.Replace($script:ApiKeyResolved, '<api-key>')
}

# ---------------------------------------------------------------------------
#  Core REST call - retries 429 / 5xx / transport errors, honours Retry-After
# ---------------------------------------------------------------------------
function Invoke-AdminApi {
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        $Body,
        [switch]$ThrowOnFailure      # fatal for setup calls; off for per-user writes
    )

    $params = @{
        Method          = $Method
        Uri             = $Uri
        Headers         = @{ Authorization = "Bearer $script:ApiKeyResolved"; Accept = 'application/json' }
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($null -ne $Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
        $params['ContentType'] = 'application/json'
    }

    for ($attempt = 0; $attempt -le $Config.MaxRetries; $attempt++) {
        try {
            $resp = Invoke-WebRequest @params
            $data = $null
            if ($resp.Content) { try { $data = $resp.Content | ConvertFrom-Json } catch { $data = $null } }
            return [pscustomobject]@{ Ok = $true; StatusCode = [int]$resp.StatusCode; Data = $data; Message = '' }
        } catch {
            $status = 0
            try { if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode } } catch { $status = 0 }
            $retryable = ($status -eq 429) -or ($status -ge 500) -or ($status -eq 0)

            if ($retryable -and $attempt -lt $Config.MaxRetries) {
                $delay = $Config.RetryBaseSeconds * [math]::Pow(2, $attempt)
                $ra = Get-RetryAfterSeconds -ErrorRecord $_
                if ($ra -gt 0) { $delay = $ra }
                Write-Log ("HTTP {0} on {1} {2} - retry {3}/{4} in {5}s" -f $status, $Method, (Protect-Secret $Uri), ($attempt + 1), $Config.MaxRetries, $delay) 'WARN'
                Start-Sleep -Seconds $delay
                continue
            }

            $msg = Protect-Secret (Get-ErrorDetail -ErrorRecord $_)
            if ($ThrowOnFailure) {
                throw ("{0} {1} failed (HTTP {2}): {3}" -f $Method, (Protect-Secret $Uri), $status, $msg)
            }
            return [pscustomobject]@{ Ok = $false; StatusCode = $status; Data = $null; Message = $msg }
        }
    }
}

# ---------------------------------------------------------------------------
#  Directories
# ---------------------------------------------------------------------------
function Get-DirectoryList {
    $dirs = @()
    $uri = "{0}/v2/orgs/{1}/directories?limit=100" -f $script:BaseUrlTrimmed, $script:OrgIdResolved
    $page = 0
    while ($uri -and $page -lt $Config.MaxPagesPerRequest) {
        $page++
        $r = Invoke-AdminApi -Method GET -Uri $uri -ThrowOnFailure
        foreach ($d in (Get-PageItems -Page $r.Data)) { $dirs += $d }
        $nextUri = Get-NextUri -Page $r.Data -CurrentUri $uri
        if ($nextUri -eq $uri) { break }
        $uri = $nextUri
    }
    return $dirs
}

# ---------------------------------------------------------------------------
#  Email -> accountId resolution, per directory, in batches
# ---------------------------------------------------------------------------
function Resolve-UsersByEmail {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Emails,
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Directories
    )

    $map = @{}   # lowercased email -> array of match records
    if ($Emails.Count -eq 0) { return $map }

    foreach ($dir in $Directories) {
        $dirId = [string](Get-Prop $dir 'directoryId')
        if (-not $dirId) { continue }
        $dirName = [string](Get-Prop $dir 'name')
        if (-not $dirName) { $dirName = $dirId }

        $searchUri = "{0}/v2/orgs/{1}/directories/{2}/users/search" -f $script:BaseUrlTrimmed, $script:OrgIdResolved, [uri]::EscapeDataString($dirId)
        $found = 0

        for ($i = 0; $i -lt $Emails.Count; $i += $Config.SearchBatchSize) {
            $last = [math]::Min($i + $Config.SearchBatchSize, $Emails.Count) - 1
            $chunk = @($Emails[$i..$last])

            $cursor = $null
            $pages = 0
            do {
                $pages++
                $body = @{ emails = $chunk; limit = $Config.SearchBatchSize }
                if ($cursor) { $body['cursor'] = $cursor }

                $r = Invoke-AdminApi -Method POST -Uri $searchUri -Body $body -ThrowOnFailure

                foreach ($u in (Get-PageItems -Page $r.Data)) {
                    $mail = [string](Get-Prop $u 'email')
                    $acct = [string](Get-Prop $u 'accountId')
                    if (-not $mail -or -not $acct) { continue }
                    $key = $mail.ToLowerInvariant()
                    if (-not $map.ContainsKey($key)) { $map[$key] = @() }
                    # one record per (email, directory)
                    if (-not (@($map[$key]) | Where-Object { $_.DirectoryId -eq $dirId })) {
                        $map[$key] += [pscustomobject]@{
                            Email            = $mail
                            AccountId        = $acct
                            DirectoryId      = $dirId
                            DirectoryName    = $dirName
                            Name             = [string](Get-Prop $u 'name')
                            AccountType      = [string](Get-Prop $u 'accountType')
                            Status           = [string](Get-Prop $u 'status')
                            AccountStatus    = [string](Get-Prop $u 'accountStatus')
                            ClaimStatus      = [string](Get-Prop $u 'claimStatus')
                            MembershipStatus = [string](Get-Prop $u 'membershipStatus')
                        }
                        $found++
                    }
                }

                $nextCursor = Get-NextCursor -Page $r.Data
                if ($nextCursor -eq $cursor) { $nextCursor = $null }
                $cursor = $nextCursor
            } while ($cursor -and $pages -lt $Config.MaxPagesPerRequest)
        }

        Write-Log ("Directory '{0}': matched {1} of {2} email(s)" -f $dirName, $found, $Emails.Count)
    }

    return $map
}

function Test-IsSuspended {
    # Only the directory 'status' field decides this. accountStatus describes the
    # Atlassian account itself, which for an external user is none of our business
    # - reading it here would wrongly skip external users we can still suspend.
    param([Parameter(Mandatory)]$Record)
    $s = ([string]$Record.Status).ToLowerInvariant()
    if (-not $s) { return $false }   # unknown state -> attempt it; suspend is idempotent
    return ($s -match [string]$Config.SuspendedPattern)
}

function New-PlanRow {
    # One shape for every plan row - Export-Csv takes its columns from the first
    # object, so they all have to match.
    param(
        [Parameter(Mandatory)][string]$Email,
        [string]$Name = '', [string]$AccountId = '', [string]$AccountType = '',
        [string]$DirectoryId = '', [string]$DirectoryName = '',
        [string]$Status = '', [string]$ClaimStatus = '',
        [Parameter(Mandatory)][string]$Classification,
        [Parameter(Mandatory)][string]$Action,
        [string]$StatusCode = '', [string]$Message = ''
    )
    return [pscustomobject]@{
        Email          = $Email
        Name           = $Name
        AccountId      = $AccountId
        AccountType    = $AccountType
        DirectoryId    = $DirectoryId
        DirectoryName  = $DirectoryName
        Status         = $Status
        ClaimStatus    = $ClaimStatus
        Classification = $Classification
        Action         = $Action
        StatusCode     = $StatusCode
        Message        = $Message
    }
}

# ---------------------------------------------------------------------------
#  CSV input
# ---------------------------------------------------------------------------
function Read-EmailList {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "CSV not found: $Path" }

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1
    if ($null -eq $firstLine -or -not ([string]$firstLine).Trim()) { throw "CSV is empty: $Path" }

    $column = [string]$Config.EmailColumn
    if ($firstLine -match '@') {
        # no header row - the first line is already an address
        Write-Log "First CSV line looks like an address - reading the file as headerless" 'WARN'
        $rows = @(Import-Csv -LiteralPath $Path -Header 'email')
        $column = 'email'
    } else {
        $rows = @(Import-Csv -LiteralPath $Path)
        if ($rows.Count -eq 0) { throw "CSV has a header but no data rows: $Path" }
        $cols = @($rows[0].PSObject.Properties.Name)
        if (-not $column -or -not ($cols -contains $column)) {
            $guess = @($cols | Where-Object { $_ -match '(?i)mail' })
            if ($guess.Count -ge 1) {
                $column = $guess[0]
                Write-Log ("Using column '{0}' for emails" -f $column) 'WARN'
            } elseif ($cols.Count -eq 1) {
                $column = $cols[0]
                Write-Log ("Using the only column '{0}' for emails" -f $column) 'WARN'
            } else {
                throw ("No email column in {0}. Columns found: {1}. Set EmailColumn in the config block." -f $Path, ($cols -join ', '))
            }
        }
    }

    $valid = New-Object System.Collections.Generic.List[string]
    $malformed = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $blank = 0
    $dupes = 0

    foreach ($row in $rows) {
        $v = ([string](Get-Prop $row $column)).Trim().Trim('"').Trim()
        if (-not $v) { $blank++; continue }
        $key = $v.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { $dupes++; continue }
        $seen[$key] = $true
        if ($key -match '^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$') { $valid.Add($key) } else { $malformed.Add($v) }
    }

    return [pscustomobject]@{
        Column         = $column
        Valid          = @($valid)
        Malformed      = @($malformed)
        BlankCount     = $blank
        DuplicateCount = $dupes
        RowCount       = $rows.Count
    }
}

# ===========================================================================
#  MAIN
# ===========================================================================
$exitCode = 0
try {
    $script:BaseUrlTrimmed = ([string]$Config.BaseUrl).TrimEnd('/')

    # --- credentials -------------------------------------------------------
    $script:OrgIdResolved = $OrgId
    if (-not $script:OrgIdResolved) { $script:OrgIdResolved = [string]$Config.OrgId }
    if (-not $script:OrgIdResolved) { $script:OrgIdResolved = [Environment]::GetEnvironmentVariable($Config.OrgIdEnvVar) }
    if (-not $script:OrgIdResolved) {
        throw ("No organization id. Pass -OrgId, set OrgId in the config block, or set the env var {0}. Find it at admin.atlassian.com -> Settings -> API keys." -f $Config.OrgIdEnvVar)
    }

    $script:ApiKeyResolved = $ApiKey
    if (-not $script:ApiKeyResolved) { $script:ApiKeyResolved = [string]$Config.ApiKey }
    if (-not $script:ApiKeyResolved) { $script:ApiKeyResolved = [Environment]::GetEnvironmentVariable($Config.ApiKeyEnvVar) }
    if (-not $script:ApiKeyResolved) {
        throw ("No API key. Set the env var {0} (recommended) or pass -ApiKey. Create one at admin.atlassian.com -> Settings -> API keys." -f $Config.ApiKeyEnvVar)
    }

    # --- paths -------------------------------------------------------------
    $csv = if ($CsvPath) { $CsvPath } else { [string]$Config.CsvPath }
    $outDir = [string]$Config.OutputDir
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

    $runKind = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    Write-Log ("Mode: {0} | Org: {1} | CSV: {2} | {3}" -f $Mode, $script:OrgIdResolved, $csv, $runKind) 'OK'

    # --- read the CSV ------------------------------------------------------
    $csvData = Read-EmailList -Path $csv
    Write-Log ("CSV: {0} row(s), column '{1}' -> {2} unique address(es), {3} duplicate(s), {4} blank(s), {5} malformed" -f $csvData.RowCount, $csvData.Column, $csvData.Valid.Count, $csvData.DuplicateCount, $csvData.BlankCount, $csvData.Malformed.Count)
    foreach ($bad in $csvData.Malformed) { Write-Log ("Malformed address skipped: '{0}'" -f $bad) 'WARN' }

    $excluded = @($Config.ExcludeEmails | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ })
    $skippedByExclude = @($csvData.Valid | Where-Object { $excluded -contains $_ })
    $targets = @($csvData.Valid | Where-Object { $excluded -notcontains $_ })
    foreach ($e in $skippedByExclude) { Write-Log ("Excluded by config, will not be touched: {0}" -f $e) 'WARN' }
    if ($targets.Count -eq 0) { throw "Nothing to do: no usable addresses left after validation and exclusions." }

    # --- directories -------------------------------------------------------
    $allDirs = @(Get-DirectoryList)
    if ($allDirs.Count -eq 0) { throw ("No directories returned for org {0}. Check the org id and the API key's permissions." -f $script:OrgIdResolved) }
    $dirs = $allDirs
    if (@($Config.DirectoryIds).Count -gt 0) {
        $dirs = @($allDirs | Where-Object { @($Config.DirectoryIds) -contains [string](Get-Prop $_ 'directoryId') })
        if ($dirs.Count -eq 0) { throw "None of the configured DirectoryIds exist in this org." }
    }
    Write-Log ("Directories in scope: {0}" -f (($dirs | ForEach-Object { "{0} ({1})" -f (Get-Prop $_ 'name'), (Get-Prop $_ 'directoryId') }) -join ', '))

    # --- resolve emails ---------------------------------------------------
    $map = Resolve-UsersByEmail -Emails $targets -Directories $dirs

    # --- build the action plan --------------------------------------------
    $wantSuspended = ($Mode -eq 'Suspend')
    $plan = New-Object System.Collections.Generic.List[object]

    foreach ($mail in $csvData.Malformed) {
        $plan.Add((New-PlanRow -Email $mail -Classification 'Malformed' -Action 'Skip' -Message 'Not a valid email address'))
    }
    foreach ($mail in $skippedByExclude) {
        $plan.Add((New-PlanRow -Email $mail -Classification 'Excluded' -Action 'Skip' -Message 'Listed in ExcludeEmails'))
    }
    foreach ($mail in $targets) {
        $hits = @()
        if ($map.ContainsKey($mail)) { $hits = @($map[$mail]) }

        if ($hits.Count -eq 0) {
            $plan.Add((New-PlanRow -Email $mail -Classification 'NotFound' -Action 'Skip' `
                -Message 'No account with this email in the directories searched'))
            continue
        }

        foreach ($m in $hits) {
            $isSuspended = Test-IsSuspended -Record $m
            $note = if ($hits.Count -gt 1) { "Present in {0} directories" -f $hits.Count } else { '' }

            $common = @{
                Email       = $m.Email
                Name        = $m.Name
                AccountId   = $m.AccountId
                AccountType = $m.AccountType
                DirectoryId = $m.DirectoryId
                DirectoryName = $m.DirectoryName
                Status      = $m.Status
                ClaimStatus = $m.ClaimStatus
            }

            if ($Config.SkipAlreadyInState -and ($isSuspended -eq $wantSuspended)) {
                $cls = if ($wantSuspended) { 'AlreadySuspended' } else { 'AlreadyActive' }
                $plan.Add((New-PlanRow @common -Classification $cls -Action 'Skip' `
                    -Message (("accountStatus={0} membershipStatus={1} {2}" -f $m.AccountStatus, $m.MembershipStatus, $note).Trim())))
            } else {
                $plan.Add((New-PlanRow @common -Classification 'Resolved' -Action $Mode -Message $note))
            }
        }
    }

    # --- summary -----------------------------------------------------------
    $toAction = @($plan | Where-Object { $_.Classification -eq 'Resolved' })
    Write-Log '--- plan ---------------------------------------------------'
    foreach ($g in ($plan | Group-Object Classification | Sort-Object Name)) {
        Write-Log ("{0,-18} {1}" -f $g.Name, $g.Count)
    }
    Write-Log ("{0,-18} {1}" -f 'TO BE ACTIONED', $toAction.Count) 'OK'
    Write-Log '------------------------------------------------------------'

    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')

    # --- dry run ends here -------------------------------------------------
    if (-not $Apply) {
        $planFile = Join-Path $outDir ("{0}-plan-{1}.csv" -f $Mode.ToLowerInvariant(), $stamp)
        $plan | Export-Csv -LiteralPath $planFile -NoTypeInformation -Encoding UTF8
        Write-Log ("Plan written to {0}" -f $planFile)
        Write-Log ("DRY RUN - nothing was changed. Re-run with -Apply to {0} {1} user(s)." -f $Mode.ToLowerInvariant(), $toAction.Count) 'OK'
        $plan
        exit 0
    }

    if ($toAction.Count -eq 0) {
        Write-Log "Nothing to action - every address was skipped." 'WARN'
        $plan
        exit 0
    }

    # --- confirmation ------------------------------------------------------
    if ($Config.RequireTypedConfirm -and -not $Yes) {
        [Console]::Error.WriteLine('')
        [Console]::Error.Write(("About to {0} {1} user(s) in org {2}. Type the number to confirm: " -f $Mode.ToUpperInvariant(), $toAction.Count, $script:OrgIdResolved))
        $answer = ([string](Read-Host)).Trim()
        if ($answer -ne [string]$toAction.Count) {
            Write-Log ("Aborted - expected '{0}', got '{1}'. Nothing was changed." -f $toAction.Count, $answer) 'ERROR'
            exit 2
        }
    }

    # --- apply -------------------------------------------------------------
    $verb = $Mode.ToLowerInvariant()   # 'suspend' | 'restore'
    $doneLabel = if ($Mode -eq 'Suspend') { 'Suspended' } else { 'Restored' }
    $ok = 0
    $failed = 0
    $i = 0

    foreach ($row in $toAction) {
        $i++
        $uri = "{0}/v2/orgs/{1}/directories/{2}/users/{3}/{4}" -f $script:BaseUrlTrimmed, $script:OrgIdResolved, [uri]::EscapeDataString($row.DirectoryId), [uri]::EscapeDataString($row.AccountId), $verb

        if (-not $PSCmdlet.ShouldProcess($row.Email, $Mode)) {
            $row.Action = 'WhatIf'
            $row.Message = 'Skipped by -WhatIf'
            continue
        }

        $body = $null
        if ($Config.SendEmptyJsonBody) { $body = @{} }
        $r = Invoke-AdminApi -Method POST -Uri $uri -Body $body

        $row.StatusCode = $r.StatusCode
        if ($r.Ok) {
            $ok++
            $row.Classification = $doneLabel
            $row.Message = 'OK'
            Write-Log ("[{0}/{1}] {2} {3} -> HTTP {4}" -f $i, $toAction.Count, $verb, $row.Email, $r.StatusCode) 'OK'
        } else {
            $failed++
            $row.Classification = 'Failed'
            $row.Message = $r.Message
            Write-Log ("[{0}/{1}] {2} {3} -> HTTP {4}: {5}" -f $i, $toAction.Count, $verb, $row.Email, $r.StatusCode, $r.Message) 'ERROR'
        }

        if ($Config.ThrottleMs -gt 0 -and $i -lt $toAction.Count) { Start-Sleep -Milliseconds $Config.ThrottleMs }
    }

    # --- results -----------------------------------------------------------
    $resultFile = Join-Path $outDir ("{0}-result-{1}.csv" -f $verb, $stamp)
    $plan | Export-Csv -LiteralPath $resultFile -NoTypeInformation -Encoding UTF8
    Write-Log ("Results written to {0}" -f $resultFile)

    $skipped = @($plan | Where-Object { $_.Action -eq 'Skip' }).Count
    $level = if ($failed -gt 0) { 'WARN' } else { 'OK' }
    Write-Log ("Done: {0} {1}ed, {2} skipped, {3} failed" -f $ok, $verb, $skipped, $failed) $level
    if ($failed -gt 0) { $exitCode = 1 }

    $plan
}
catch {
    Write-Log ([string]$_.Exception.Message) 'ERROR'
    $exitCode = 2
}

exit $exitCode
