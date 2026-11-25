<#
.SYNOPSIS
    Sets (or removes) the default sensitivity label on many SharePoint document libraries.

.DESCRIPTION
    Reads a CSV of SiteUrl + LibraryName rows (or one site with -AllLibraries), connects to each
    site with PnP.PowerShell, sets the library's default sensitivity label, then reads it back
    to confirm. Writes a results CSV next to the input.

    The default label only applies to files that are NEW or EDITED after it is set, and never
    replaces a label a file already has (unless the existing label is lower priority and was
    applied automatically). Existing untouched files stay unlabelled.

    Needs:
      - PowerShell 7 and PnP.PowerShell 3.x
      - An Entra app registration for PnP interactive login in the target tenant (-ClientId)
        with delegated SharePoint AllSites.FullControl (or Sites.FullControl). No Graph permission
        is needed: labels are looked up through SharePoint's own label endpoint
      - Site owner (or higher) on each site
      - Sensitivity labels for Office files enabled in SharePoint (Set-SPOTenant -EnableAIPIntegration $true)
      - The label must be published to the account running the script, and scoped to "Files"

.PARAMETER CsvPath
    CSV with columns SiteUrl, LibraryName (library title, e.g. "Documents"). Optional column
    Label overrides -Label for that row.

.PARAMETER SiteUrl
    Single site to process instead of a CSV. Use with -LibraryName or -AllLibraries.

.PARAMETER LibraryName
    One or more library titles on -SiteUrl.

.PARAMETER AllLibraries
    With -SiteUrl: every visible document library on the site (skips hidden/system libraries).

.PARAMETER Label
    Label name, "Parent \ Child" path for a sub-label, or GUID. Ignored with -Remove.
    Run with -ListLabels to see what's available.

.PARAMETER ListLabels
    Connects to the first site, lists the file labels published to you, and stops.

.PARAMETER Remove
    Clears the default label instead of setting one.

.PARAMETER LabelExistingFiles
    After the default is set, also labels files ALREADY in the library. For each unlabelled Office
    file (docx/xlsx/pptx and macro versions) and PDF, the script checks the file out, saves the same
    content back and checks it in with "overwrite" (replaces the current version, adds none). That
    save makes SharePoint apply the library default (usually within a minute). The script then puts
    the original Modified date and Modified By back. Net side effects: one new version per file (the
    label being written into the file - unavoidable with any method), and anything that triggers on
    "file modified" (Power Automate flows, alerts) may fire. Skips files that already have a label,
    are checked out, or are empty. Writes a per-file CSV next to the results.
    Test with -WhatIf and -MaxFilesPerLibrary first.

.PARAMETER MaxFilesPerLibrary
    With -LabelExistingFiles: only process this many files per library (0 = all).

.PARAMETER ClientId
    Client ID of the PnP app registration in the target tenant.

.EXAMPLE
    .\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -Label "Confidential" -ClientId <guid> -WhatIf

.EXAMPLE
    .\Set-DefaultLibraryLabel.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance -AllLibraries -Label "Confidential" -ClientId <guid>

.EXAMPLE
    .\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -Remove -ClientId <guid>
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Csv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'Site')]
    [string]$SiteUrl,

    [Parameter(ParameterSetName = 'Site')]
    [string[]]$LibraryName,

    [Parameter(ParameterSetName = 'Site')]
    [switch]$AllLibraries,

    [string]$Label,

    [switch]$Remove,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [string]$ResultsPath,

    # Cache the PnP login on this PC so later runs don't prompt again
    [switch]$PersistLogin,

    [switch]$ListLabels,

    # Also label files already in the library (see .PARAMETER LabelExistingFiles)
    [switch]$LabelExistingFiles,

    # With -LabelExistingFiles: stop after this many files per library (for testing / batching)
    [int]$MaxFilesPerLibrary = 0
)

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'Run this in PowerShell 7 (pwsh). PnP.PowerShell 3.x does not load in Windows PowerShell 5.1.' }
if ($PSCmdlet.ParameterSetName -eq 'Site' -and -not $AllLibraries -and -not $LibraryName -and -not $ListLabels) { throw 'With -SiteUrl, give -LibraryName or -AllLibraries.' }

Import-Module PnP.PowerShell -MinimumVersion 3.0 -ErrorAction Stop

# ---- Build the work list -------------------------------------------------------------------
$rows = if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    Import-Csv -LiteralPath $CsvPath | ForEach-Object {
        [pscustomobject]@{ SiteUrl = $_.SiteUrl.Trim().TrimEnd('/'); LibraryName = $_.LibraryName.Trim(); Label = if ($_.PSObject.Properties['Label'] -and $_.Label) { $_.Label.Trim() } else { $Label } }
    }
} else {
    $libs = if ($AllLibraries) { @('*') } else { $LibraryName }
    $libs | ForEach-Object { [pscustomobject]@{ SiteUrl = $SiteUrl.TrimEnd('/'); LibraryName = $_; Label = $Label } }
}

if (-not $Remove -and -not $ListLabels) {
    $noLabel = @($rows | Where-Object { -not $_.Label })
    if ($noLabel) { throw "No label for: $(($noLabel | ForEach-Object { "$($_.SiteUrl) / $($_.LibraryName)" }) -join '; '). Give -Label, fill the CSV Label column, or use -Remove." }
}

if (-not $ResultsPath) {
    $base = if ($CsvPath) { Split-Path -Parent (Resolve-Path -LiteralPath $CsvPath) } else { $PWD.Path }
    $ResultsPath = Join-Path $base ("DefaultLabel-Results-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}

# ---- Helpers --------------------------------------------------------------------------------
$labelCache = @{}
$script:fileLabels = $null

# File-scoped labels published to the signed-in user, read from SharePoint's own endpoint (the one the
# Library settings panel uses), so no Microsoft Graph permission is needed. Sub-labels are flattened
# and get a Path of "Parent \ Child".
$script:labelLookupTried = $false
$script:labelLookupError = ''
function Get-FileLabels {
    if ($script:labelLookupTried) { return $script:fileLabels }
    $script:labelLookupTried = $true
    $flat = [System.Collections.Generic.List[object]]::new()
    $errs = @()

    # 1) SharePoint's own endpoint (used by the Library settings panel). Some tenants reject PnP app tokens here.
    try {
        $raw = Invoke-PnPSPRestMethod -Method Get -Url "/_api/sp.sensitivitylabelcollection.getsensitivitylabelsasstream?expand='sublabels'&filter='applicableto has ''file'''" -Raw
        function Add-Labels($items, $parent) {
            foreach ($l in $items) {
                $path = if ($parent) { "$parent \ $($l.displayName)" } else { $l.displayName }
                $flat.Add([pscustomobject]@{ Id = [string]$l.id; Name = $l.name; DisplayName = $l.displayName; Path = $path; IsParent = [bool]$l.isParent })
                if ($l.sublabels) { Add-Labels $l.sublabels $l.displayName }
            }
        }
        Add-Labels ($raw | ConvertFrom-Json) $null
    } catch { $errs += "SharePoint: $($_.Exception.Message)" }

    # 2) Microsoft Graph (needs InformationProtectionPolicy.Read on the app)
    if ($flat.Count -eq 0) {
        try {
            $g = @(Get-PnPAvailableSensitivityLabel)
            foreach ($l in $g) {
                $dn = if ($l.DisplayName) { $l.DisplayName } else { $l.Name }
                $parentName = if ($l.Parent) { $l.Parent.Name } else { $null }
                $flat.Add([pscustomobject]@{ Id = [string]$l.Id; Name = $l.Name; DisplayName = $dn; Path = $(if ($parentName) { "$parentName \ $dn" } else { $dn }); IsParent = $false })
            }
        } catch { $errs += "Graph: $($_.Exception.Message)" }
    }

    if ($flat.Count -gt 0) { $script:fileLabels = $flat } else { $script:labelLookupError = $errs -join ' | ' }
    return $script:fileLabels
}

function Resolve-Label([string]$nameOrId) {
    if ($labelCache.ContainsKey($nameOrId)) { return $labelCache[$nameOrId] }
    $all = Get-FileLabels
    $guid = [guid]::Empty
    $isGuid = [guid]::TryParse($nameOrId, [ref]$guid)
    if (-not $all) {
        # Couldn't read the label list. A GUID still works - SharePoint checks it when it's set.
        if ($isGuid) {
            if (-not $script:warnedNoLookup) {
                Write-Warning "Couldn't read the label list, so label GUIDs are used as given (SharePoint checks them when set). Details: $($script:labelLookupError)"
                $script:warnedNoLookup = $true
            }
            $labelCache[$nameOrId] = [pscustomobject]@{ Id = $guid.ToString(); Path = $nameOrId }
            return $labelCache[$nameOrId]
        }
        throw "Couldn't read the label list to look up '$nameOrId' ($($script:labelLookupError)). Pass the label GUID instead: in Security & Compliance PowerShell run Connect-IPPSSession; Get-Label | Format-Table DisplayName, Guid, ParentId"
    }
    $match = if ($isGuid) {
        @($all | Where-Object { $_.Id -eq $guid.ToString() })
    } else {
        # Accepts the label's name, display name, or "Parent \ Child" path
        @($all | Where-Object { $_.Name -eq $nameOrId -or $_.DisplayName -eq $nameOrId -or $_.Path -eq $nameOrId })
    }
    if ($match.Count -eq 0) { throw "Label '$nameOrId' not found, not published to you, or not scoped to files. Available: $(($all | Where-Object { -not $_.IsParent } | ForEach-Object Path) -join '; ')" }
    if ($match.Count -gt 1) { throw "Label '$nameOrId' matches more than one label: $(($match | ForEach-Object Path) -join '; '). Use 'Parent \ Child' or the GUID." }
    if ($match[0].IsParent) { throw "'$($match[0].Path)' is a parent label and can't be applied. Pick one of its sub-labels: $(($all | Where-Object { $_.Path -like "$($match[0].DisplayName) \ *" } | ForEach-Object Path) -join '; ')" }
    $labelCache[$nameOrId] = $match[0]
    return $match[0]
}

function Get-LibraryDefaultLabel($list) {
    $l = Get-PnPList -Identity $list.Id -Includes DefaultSensitivityLabelForLibrary
    return [string]$l.DefaultSensitivityLabelForLibrary
}

$labelableExt = @('docx', 'docm', 'xlsx', 'xlsm', 'xlsb', 'pptx', 'pptm', 'ppsx', 'pdf')

# All files in a library with their current label, paged (works whether or not the label columns exist yet)
function Get-LibraryFiles($list) {
    $viewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query>" +
               "<ViewFields><FieldRef Name='ID'/><FieldRef Name='FileRef'/><FieldRef Name='FileLeafRef'/><FieldRef Name='File_x0020_Size'/>" +
               "<FieldRef Name='CheckoutUser'/><FieldRef Name='_IpLabelId'/><FieldRef Name='_DisplayName'/></ViewFields><RowLimit Paged='TRUE'>500</RowLimit></View>"
    $url = "/_api/web/lists(guid'$($list.Id)')/RenderListDataAsStream"
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $null
    do {
        $u = if ($next) { "$url$($next -replace '^\?', '?')" } else { $url }
        $r = Invoke-PnPSPRestMethod -Method Post -Url $u -Content @{ parameters = @{ RenderOptions = 2; ViewXml = $viewXml } }
        $r.Row | ForEach-Object { $all.Add($_) }
        $next = $r.NextHref
    } while ($next)
    return $all
}

# Re-save one file so SharePoint applies the library default label, then restore Modified/Modified By
function Invoke-Relabel($list, $row, $webRel) {
    $item = Get-PnPListItem -List $list.Id -Id $row.ID -Fields Modified, Editor
    $origMod = $item['Modified']; $origEd = $item['Editor']
    $leaf = $row.FileLeafRef
    $folder = $row.FileRef.Substring(0, $row.FileRef.Length - $leaf.Length - 1)
    if ($webRel -ne '/') { $folder = $folder.Substring($webRel.Length) }
    $folder = $folder.TrimStart('/')
    # Check out -> save the same bytes -> check in with OverwriteCheckIn. The overwrite check-in replaces the
    # current version instead of adding one, but still counts as a save, so SharePoint applies the library
    # default label (tested: labelled within ~40 s). Net effect: one new version, written by the label itself.
    $ms = Get-PnPFile -Url $row.FileRef -AsMemoryStream
    Set-PnPFileCheckedOut -Url $row.FileRef
    try {
        Add-PnPFile -Stream $ms -FileName $leaf -Folder $folder | Out-Null
        Set-PnPFileCheckedIn -Url $row.FileRef -CheckinType OverwriteCheckIn -Comment ''
    } catch {
        try { Undo-PnPFileCheckedOut -Url $row.FileRef } catch { }
        throw
    } finally { $ms.Dispose() }
    $vals = @{ Modified = $origMod }
    if ($origEd -and $origEd.Email) { $vals.Editor = $origEd.Email }
    try { Set-PnPListItem -List $list.Id -Identity $row.ID -Values $vals -UpdateType UpdateOverwriteVersion | Out-Null; return 'Re-saved; Modified/Editor restored' }
    catch {
        # Editor may be a departed user - keep the date at least
        Set-PnPListItem -List $list.Id -Identity $row.ID -Values @{ Modified = $origMod } -UpdateType UpdateOverwriteVersion | Out-Null
        return "Re-saved; Modified restored, Modified By not restored ($($_.Exception.Message))"
    }
}

$fileResults = [System.Collections.Generic.List[object]]::new()

if ($ListLabels) {
    Connect-PnPOnline -Url $rows[0].SiteUrl -Interactive -ClientId $ClientId -PersistLogin:$PersistLogin
    $l = Get-FileLabels
    if ($l) { $l | Select-Object Path, Id, IsParent | Format-Table -AutoSize }
    else { Write-Warning "Couldn't read labels: $($script:labelLookupError). Get GUIDs with Connect-IPPSSession; Get-Label | Format-Table DisplayName, Guid, ParentId" }
    return
}

# ---- Process --------------------------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()
$currentSite = $null

foreach ($siteGroup in ($rows | Group-Object SiteUrl)) {
    $site = $siteGroup.Name
    Write-Host "`n== $site" -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $site -Interactive -ClientId $ClientId -PersistLogin:$PersistLogin
        $currentSite = $site
    } catch {
        foreach ($r in $siteGroup.Group) {
            $results.Add([pscustomobject]@{ SiteUrl = $site; LibraryName = $r.LibraryName; LabelRequested = $r.Label; Before = ''; After = ''; Status = 'Failed'; Detail = "Connect failed: $($_.Exception.Message)" })
        }
        Write-Warning "Could not connect: $($_.Exception.Message)"
        continue
    }

    # Expand '*' into every visible document library
    $targets = foreach ($r in $siteGroup.Group) {
        if ($r.LibraryName -eq '*') {
            Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden -and -not $_.IsCatalog -and -not $_.IsSiteAssetsLibrary -and -not $_.IsSystemList -and $_.EntityTypeName -notin @('FormServerTemplates', 'Style_x0020_Library', 'SitePages', 'SiteAssets') } |
                ForEach-Object { [pscustomobject]@{ SiteUrl = $site; LibraryName = $_.Title; Label = $r.Label } }
        } else { $r }
    }

    foreach ($t in $targets) {
        $res = [pscustomobject]@{ SiteUrl = $site; LibraryName = $t.LibraryName; LabelRequested = if ($Remove) { '(remove)' } else { $t.Label }; Before = ''; After = ''; Status = ''; Detail = '' }
        try {
            $list = Get-PnPList -Identity $t.LibraryName
            if (-not $list) { throw "Library '$($t.LibraryName)' not found" }
            if ($list.BaseTemplate -ne 101) { throw "'$($t.LibraryName)' is not a document library (BaseTemplate $($list.BaseTemplate))" }

            $res.Before = Get-LibraryDefaultLabel $list
            $wantId = if ($Remove) { '' } else { (Resolve-Label $t.Label).Id }

            if ($res.Before -eq $wantId) {
                $res.After = $res.Before; $res.Status = 'NoChange'; $res.Detail = 'Already set'
            } elseif ($PSCmdlet.ShouldProcess("$site / $($t.LibraryName)", $(if ($Remove) { 'Remove default sensitivity label' } else { "Set default sensitivity label to $($t.Label) ($wantId)" }))) {
                if ($Remove) { Set-PnPList -Identity $list.Id -DefaultSensitivityLabelForLibrary $null | Out-Null }
                else { Set-PnPList -Identity $list.Id -DefaultSensitivityLabelForLibrary $wantId | Out-Null }
                $res.After = Get-LibraryDefaultLabel $list
                if ($res.After -eq $wantId) { $res.Status = 'Done' } else { $res.Status = 'Failed'; $res.Detail = "Read-back '$($res.After)' does not match '$wantId'" }
            } else {
                $res.After = $res.Before; $res.Status = 'WhatIf'
            }
        } catch {
            $res.Status = 'Failed'; $res.Detail = $_.Exception.Message
        }
        $colour = switch ($res.Status) { 'Done' { 'Green' } 'Failed' { 'Red' } default { 'Gray' } }
        Write-Host ("  {0,-8} {1}  {2}" -f $res.Status, $res.LibraryName, $res.Detail) -ForegroundColor $colour
        $results.Add($res)

        # ---- Existing files ----
        $defaultInPlace = $res.Status -in 'Done', 'NoChange' -or ($res.Status -eq 'WhatIf' -and -not $Remove)
        if ($LabelExistingFiles -and -not $Remove -and $defaultInPlace) {
            try {
                $webRel = (Get-PnPWeb).ServerRelativeUrl
                $files = Get-LibraryFiles $list
                $todo = 0; $counts = @{}
                foreach ($f in $files) {
                    $ext = ([IO.Path]::GetExtension($f.FileLeafRef)).TrimStart('.').ToLower()
                    $fr = [pscustomobject]@{ SiteUrl = $site; LibraryName = $t.LibraryName; File = $f.FileRef; LabelBefore = [string]$f._DisplayName; Status = ''; Detail = '' }
                    if ($ext -notin $labelableExt) { $fr.Status = 'Skipped'; $fr.Detail = "File type .$ext can't be labelled" }
                    elseif ($f._IpLabelId) { $fr.Status = 'Skipped'; $fr.Detail = 'Already labelled' }
                    elseif ($f.CheckoutUser) { $fr.Status = 'Skipped'; $fr.Detail = "Checked out to $(@($f.CheckoutUser)[0].title)" }
                    elseif ([int64]("0$($f.File_x0020_Size)" -replace '[^\d]', '') -eq 0) { $fr.Status = 'Skipped'; $fr.Detail = 'Empty file' }
                    elseif ($MaxFilesPerLibrary -gt 0 -and $todo -ge $MaxFilesPerLibrary) { $fr.Status = 'Skipped'; $fr.Detail = "Over -MaxFilesPerLibrary $MaxFilesPerLibrary" }
                    else {
                        $todo++
                        if ($PSCmdlet.ShouldProcess($f.FileRef, 'Re-save so the library default label is applied')) {
                            try { $fr.Detail = Invoke-Relabel $list $f $webRel; $fr.Status = 'Resaved' }
                            catch { $fr.Status = 'Failed'; $fr.Detail = $_.Exception.Message }
                        } else { $fr.Status = 'WhatIf' }
                    }
                    $counts[$fr.Status] = 1 + [int]$counts[$fr.Status]
                    $fileResults.Add($fr)
                }
                Write-Host ("           existing files: {0}" -f (($counts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) $($_.Value)" }) -join ', ')) -ForegroundColor Gray
            } catch {
                Write-Warning "  Existing files in '$($t.LibraryName)' failed: $($_.Exception.Message)"
                $fileResults.Add([pscustomobject]@{ SiteUrl = $site; LibraryName = $t.LibraryName; File = ''; LabelBefore = ''; Status = 'Failed'; Detail = $_.Exception.Message })
            }
        }
    }
}

if ($fileResults.Count) {
    $filePath = [IO.Path]::ChangeExtension($ResultsPath, $null).TrimEnd('.') + '-files.csv'
    $fileResults | Export-Csv -LiteralPath $filePath -Encoding utf8 -WhatIf:$false
    Write-Host "`nPer-file results: $filePath"
    Write-Host "Labels are applied by SharePoint shortly after each re-save (usually under a minute, can be a few minutes)."
}

$results | Export-Csv -LiteralPath $ResultsPath -Encoding utf8 -WhatIf:$false
Write-Host "`nResults: $ResultsPath"
$results | Group-Object Status | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Name, $_.Count) }
