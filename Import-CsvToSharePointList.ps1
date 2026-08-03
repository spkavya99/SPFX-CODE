<#
=====================================================================
 Import-CsvToSharePointList.ps1
 CSV  ->  SharePoint Online list   (Connect-PnPOnline -UseWebLogin)
 Internal names (field_1, field_2 ...) resolved from DISPLAY names.
 Title is left BLANK - every CSV column maps to its own list column.
 TEST MODE = first 5 rows.  Set $TestMode = $false for the full 25k.
 Run in Windows PowerShell 5.1.
=====================================================================
#>

# ------------------------- CONFIG ------------------------------------
$SiteUrl     = "https://spc.cranenxt.com/kaizen"
$ListName    = "Kaizen Participation"
$CsvPath     = "C:\Temp\kaizenparticipationtoday.csv"

$TestMode    = $true
$TestCount   = 5
$CommitEvery = 100               # rows per server round-trip

$ConvertDatesToUtc = $true       # CSV dates are local time

# CSV column shown in progress/error messages
$KeyColumn   = 'Participant Name'

# Only needed when a CSV header differs from the SharePoint display name:
#   'CSV header' = 'SharePoint display name'
$HeaderAliases = @{
    # 'Emp ID' = 'Employee ID'
}

# CSV columns to ignore entirely (Title stays blank)
$SkipColumns = @('Title','Modified','Created','Created By','Modified By','ID','Item Type','Path')

$DateFormats = @('M/d/yyyy H:mm','M/d/yyyy','yyyy-MM-dd HH:mm:ss','yyyy-MM-dd',
                 'M/d/yyyy h:mm:ss tt','MM/dd/yyyy HH:mm:ss')
# ---------------------------------------------------------------------

Write-Host "`nCSV -> SharePoint List Import`n" -ForegroundColor Green
Write-Host "Site: $SiteUrl"
Write-Host "List: $ListName"
Write-Host "File: $CsvPath`n"

function ConvertTo-SpValue {
    param($Raw)
    if ($null -eq $Raw) { return $null }
    $v = ([string]$Raw).Trim()
    if ($v -eq '' -or $v -eq 'NULL' -or $v -eq '#N/A') { return $null }
    return $v
}

function ConvertTo-SpDate {
    param([string]$Raw)
    $v = ConvertTo-SpValue $Raw
    if ($null -eq $v) { return $null }
    if ($v -match '^00:00(:00)?(\.0)?$') { return $null }    # the 00:00.0 junk
    $dt = [datetime]::MinValue
    $ok = [datetime]::TryParseExact($v, $DateFormats, [cultureinfo]::InvariantCulture,
              [System.Globalization.DateTimeStyles]::None, [ref]$dt)
    if (-not $ok) { $ok = [datetime]::TryParse($v, [ref]$dt) }
    if (-not $ok) { return 'BAD' }
    if ($ConvertDatesToUtc) { return $dt.ToUniversalTime() }
    return $dt
}

# ---- Connect to SharePoint -------------------------------------------
Write-Host "Connecting to SharePoint..."
try {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "Connected!`n" -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# ---- Read the CSV -----------------------------------------------------
Write-Host "Reading CSV..."
try {
    if (-not (Test-Path $CsvPath)) { throw "File not found: $CsvPath" }
    $allRows = @(Import-Csv -Path $CsvPath)
    if ($allRows.Count -eq 0) { throw "CSV has no data rows" }

    $headers = $allRows[0].PSObject.Properties.Name
    $rows    = if ($TestMode) { $allRows | Select-Object -First $TestCount } else { $allRows }

    Write-Host "Found $($allRows.Count) rows, importing $(@($rows).Count)`n" -ForegroundColor Green
}
catch {
    Write-Host "CSV ERROR: $_" -ForegroundColor Red
    Disconnect-PnPOnline
    Read-Host "Press Enter to exit"
    exit
}

# ---- Resolve SharePoint internal names --------------------------------
Write-Host "Resolving list columns..."
try {
    $sysFields = @('ContentType','Attachments','Edit','DocIcon','LinkTitle','LinkTitleNoMenu',
                   'ItemChildCount','FolderChildCount','_UIVersionString','_ComplianceFlags',
                   '_ComplianceTag','_ComplianceTagWrittenTime','_ComplianceTagUserId',
                   '_IsRecord','AppAuthor','AppEditor','Author','Editor','Created','Modified')

    $allListFields = Get-PnPField -List $ListName
    $listFields = $allListFields |
        Where-Object { -not $_.Hidden -and -not $_.ReadOnlyField -and $_.InternalName -notin $sysFields }

    $byDisplay = @{}
    foreach ($f in $listFields) {
        $k = $f.Title.Trim().ToLower()
        if (-not $byDisplay.ContainsKey($k)) { $byDisplay[$k] = $f }
    }
}
catch {
    Write-Host "ERROR reading list: $_" -ForegroundColor Red
    Write-Host "Make sure list name is '$ListName' (case-sensitive)" -ForegroundColor Yellow
    Disconnect-PnPOnline
    Read-Host "Press Enter to exit"
    exit
}

# Title must not be mandatory, since we are leaving it empty
$titleField = $allListFields | Where-Object { $_.InternalName -eq 'Title' }
if ($titleField -and $titleField.Required) {
    Write-Host "`nWARNING: the Title column is marked Required." -ForegroundColor Yellow
    Write-Host "Leaving it blank will fail. Make it optional first with:" -ForegroundColor Yellow
    Write-Host "  Set-PnPField -List `"$ListName`" -Identity Title -Values @{Required=`$false}`n" -ForegroundColor Yellow
}

$map        = [ordered]@{}
$unresolved = @()
foreach ($h in $headers) {
    if ($SkipColumns -contains $h) { continue }
    $display = if ($HeaderAliases.ContainsKey($h)) { $HeaderAliases[$h] } else { $h }
    $key = $display.Trim().ToLower()
    if ($byDisplay.ContainsKey($key)) {
        $fld = $byDisplay[$key]
        $map[$h] = @{ Internal = $fld.InternalName; Type = $fld.TypeAsString }
    }
    else { $unresolved += $h }
}

Write-Host "`n--- Mapping (CSV header -> internal name) ---" -ForegroundColor Yellow
$map.GetEnumerator() | ForEach-Object {
    [pscustomobject]@{ CsvHeader=$_.Key; InternalName=$_.Value.Internal; Type=$_.Value.Type }
} | Format-Table -AutoSize

if ($unresolved) {
    Write-Host "NOT MATCHED (these columns will be skipped):" -ForegroundColor Red
    $unresolved | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }
    Write-Host "`nAvailable display names:" -ForegroundColor Yellow
    $listFields | Select-Object Title, InternalName | Format-Table -AutoSize
    Write-Host "Add each to `$HeaderAliases or `$SkipColumns." -ForegroundColor Yellow
}

if ($map.Count -eq 0) {
    Write-Host "Nothing mapped. Stopping." -ForegroundColor Red
    Disconnect-PnPOnline
    Read-Host "Press Enter to exit"
    exit
}

if ((Read-Host "Import $(@($rows).Count) row(s)? (y/n)") -ne 'y') {
    Write-Host "Cancelled." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit
}

# ---- Build values for one row -----------------------------------------
function Get-RowValues {
    param($Row)
    $values = @{}
    foreach ($h in $map.Keys) {
        $internal = $map[$h].Internal

        if ($map[$h].Type -match 'DateTime') {
            $d = ConvertTo-SpDate $Row.$h
            if ($d -eq 'BAD') { throw "Unparseable date in '$h': $($Row.$h)" }
            if ($null -ne $d) { $values[$internal] = $d }
        }
        else {
            $v = ConvertTo-SpValue $Row.$h
            if ($null -ne $v) { $values[$internal] = $v }
        }
    }
    return $values
}

# ---- Import ------------------------------------------------------------
$useCsom = -not (Get-Command New-PnPBatch -ErrorAction SilentlyContinue)
Write-Host ("`nInsert mode: " + $(if ($useCsom) { "CSOM bulk" } else { "PnP batch" })) -ForegroundColor Cyan

$i = 0; $fail = 0
$total = @($rows).Count
$sw = [System.Diagnostics.Stopwatch]::StartNew()

if ($useCsom) {
    $ctx  = Get-PnPContext
    $list = Get-PnPList -Identity $ListName
    $ctx.Load($list); $ctx.ExecuteQuery()
}
else {
    $batch = New-PnPBatch
}

foreach ($row in $rows) {
    $i++
    try {
        $values = Get-RowValues -Row $row
        if ($values.Keys.Count -eq 0) { throw "Row produced no values" }

        if ($useCsom) {
            $ci   = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $item = $list.AddItem($ci)
            foreach ($k in $values.Keys) { $item[$k] = $values[$k] }
            $item.Update()
        }
        else {
            Add-PnPListItem -List $ListName -Values $values -Batch $batch
        }
    }
    catch {
        $fail++
        Write-Host "Row $i ($($row.$KeyColumn)) failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    if (($i % $CommitEvery) -eq 0) {
        if ($useCsom) { $ctx.ExecuteQuery() }
        else { Invoke-PnPBatch -Batch $batch; $batch = New-PnPBatch }
        Write-Host "Committed $i / $total   [$([int]$sw.Elapsed.TotalSeconds)s]" -ForegroundColor Green
    }
}

# final flush
if ($useCsom) { $ctx.ExecuteQuery() }
elseif ($batch.RequestCount -gt 0) { Invoke-PnPBatch -Batch $batch }

$sw.Stop()
Write-Host "`nDone in $([int]$sw.Elapsed.TotalSeconds)s" -ForegroundColor Green
Write-Host "Imported: $($i - $fail)   Errors: $fail`n"

Disconnect-PnPOnline
Read-Host "Press Enter to exit"
