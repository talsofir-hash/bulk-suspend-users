# Atlassian org admin - bulk suspend users from a CSV

`Suspend-OrgUsers.ps1` takes a CSV with one column of email addresses and
suspends every one of those users' access in your Atlassian cloud organization.
The same script restores them with `-Mode Restore`.

Suspending keeps the account, its product roles and its group memberships, so
nothing is destroyed and the operation is reversible.

## One-time setup

1. Go to <https://admin.atlassian.com> -> **Settings** -> **API keys** ->
   *Create API key*. Copy **both** the API key and the **Organization ID** shown
   on that page - the key is never displayed again.
2. Put them in the current PowerShell session (nothing is written to disk):

   ```powershell
   $env:ATLASSIAN_ORG_ID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
   $env:ATLASSIAN_ADMIN_API_KEY = 'the-api-key'
   ```

   To keep them across sessions use `[Environment]::SetEnvironmentVariable(...,'User')`
   instead. You can also pass `-OrgId` / `-ApiKey`, or fill in `OrgId` /`ApiKey`
   in the config block - but do not commit a key.

This suspends **organization access**, which works for external (unmanaged)
users too - it is the same action as the "Suspend access" item under
*Organization* in the admin UI's `...` menu. Do not confuse it with the
*Atlassian account* actions in that menu (deactivate the account, reset
password); those need a managed account on a domain your org has verified, and
this script does not touch them.

## The CSV

One column of emails. A header row is optional - the script detects a headerless
file, and if the header is not `email` it picks the first column whose name
contains "mail", or the only column if there is just one.

```
email
leaver.one@example.com
leaver.two@example.com
```

Blank rows, duplicates and case differences are handled. Malformed values are
reported and skipped, never guessed at.

## Running it

```powershell
cd C:\path\to\atlassian-org-admin

# 1. DRY RUN (default) - reads only, changes nothing
.\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv

# 2. read logs\suspend-plan-<timestamp>.csv, check the list is right

# 3. do it - you are asked to type the user count to confirm
.\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv -Apply

# undo the same list
.\Suspend-OrgUsers.ps1 -CsvPath .\leavers.csv -Mode Restore -Apply
```

Useful switches:

| Switch | Effect |
| --- | --- |
| *(none)* | dry run: resolve, report, write the plan CSV, change nothing |
| `-Apply` | actually suspend / restore |
| `-Mode Restore` | restore access instead of suspending |
| `-Yes` | skip the typed confirmation (for unattended runs) |
| `-WhatIf` | per-user no-op even with `-Apply` |
| `-CsvPath` / `-OrgId` / `-ApiKey` | override the config block |

Exit codes: `0` all good, `1` finished with per-user failures, `2` fatal
(bad config, auth, or CSV) - nothing was changed.

## Output

Every run writes a CSV to `logs\`:
`suspend-plan-<stamp>.csv` for a dry run, `suspend-result-<stamp>.csv` (or
`restore-result-...`) for a real one. Columns:

| Column | Meaning |
| --- | --- |
| `Email`, `Name`, `AccountId` | the resolved user |
| `AccountType`, `Status`, `ClaimStatus` | what the directory says about them - `ClaimStatus` shows whether the account is external/unclaimed |
| `DirectoryId`, `DirectoryName` | which directory the action targeted |
| `Classification` | `Resolved` (will be actioned), `Suspended` / `Restored` (done), `Failed`, `NotFound`, `AlreadySuspended`, `AlreadyActive`, `Excluded`, `Malformed` |
| `Action` | `Suspend` / `Restore` / `Skip` / `WhatIf` |
| `StatusCode` | HTTP status of the write call (`204` = success) |
| `Message` | API error text, or why the row was skipped |

The result CSV of a suspend run is also the input you need for the restore -
it lists exactly who was changed.

## Configuration

Everything tunable is in the single `$Config` block at the top of the script:
credential sources, CSV path and column, directory scope, `ExcludeEmails`
(a hard never-touch list, pre-seeded with your own address so you cannot lock
yourself out), `SkipAlreadyInState`, `RequireTypedConfirm`, batch size, throttle
and retry settings, and the output directory.

`SendEmptyJsonBody` is `$true` and should stay that way. The docs say the
suspend/restore endpoints take no body, but a POST with no body carries no
`Content-Type` and Atlassian answers **HTTP 415 Unsupported Media Type**.
Sending `{}` with `Content-Type: application/json` is what actually works.

## API used

Atlassian Organizations REST API, `https://api.atlassian.com/admin`, with the
org API key as `Authorization: Bearer <key>`:

- `GET  /v2/orgs/{orgId}/directories` - discover directories
- `POST /v2/orgs/{orgId}/directories/{directoryId}/users/search` - exact-match
  email lookup (batched, `{ "emails": [...] }`) to get each `accountId`
- `POST /v2/orgs/{orgId}/directories/{directoryId}/users/{accountId}/suspend`
- `POST /v2/orgs/{orgId}/directories/{directoryId}/users/{accountId}/restore`

Docs: <https://developer.atlassian.com/cloud/admin/organization/suspend-user/>

429 and 5xx responses are retried with back-off (honouring `Retry-After`); a
failure on one user is recorded and the batch continues.
