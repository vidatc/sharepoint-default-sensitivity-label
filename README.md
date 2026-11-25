# SharePoint default sensitivity label, in bulk

PowerShell (PnP.PowerShell) to set the **default sensitivity label** on many SharePoint Online document libraries at once, instead of clicking through Library settings one library at a time. New and edited files in those libraries then get the label automatically.

Under the hood it's `Set-PnPList -DefaultSensitivityLabelForLibrary`, run against every library in a CSV (or every library on a site), with a dry run, a read-back check, an undo and a results log.

Optional extra: `-LabelExistingFiles` also labels the files already in the libraries, for free, with one new version per file.

## Why

A library's default sensitivity label only applies to files that are uploaded or edited **after** it's set. Files already in the library stay unlabelled forever unless something touches them. Microsoft's other options either cost money per file (Graph `assignSensitivityLabel` is a metered API) or run on Microsoft's schedule (Purview auto-labelling policies).

## What it does

1. Sets (or removes) the default label on each library in a CSV, or every library on a site, and reads it back to confirm.
2. With `-LabelExistingFiles`, for each unlabelled Office file (docx/xlsx/pptx and macro versions) and PDF:
   - checks the file out, saves the **same content** back, and checks it in with **overwrite**, which replaces the current version rather than adding one
   - that save makes SharePoint apply the library default label (usually within a minute)
   - restores the original **Modified** date and **Modified By**
3. Writes a results CSV, plus a per-file CSV when labelling existing files.

Files that already have a label, are checked out, or are empty are skipped.

## Requirements

- PowerShell 7 and [PnP.PowerShell](https://pnp.github.io/powershell/) 3.x
- An Entra app registration for PnP interactive login in your tenant (`-ClientId`), with delegated SharePoint `AllSites.FullControl`
- Site owner (or higher) on each site
- Sensitivity labels enabled for Office files in SharePoint (`Set-SPOTenant -EnableAIPIntegration $true`), plus PDF labelling if you want PDFs done
- Labels published to the account running the script, scoped to **Files**
- Your licensing must include default sensitivity labels for document libraries. Check the Microsoft Purview service description.

## Usage

```powershell
# Dry run
.\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -Label <label GUID> -ClientId <app id> -PersistLogin -WhatIf

# Set defaults only
.\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -Label <label GUID> -ClientId <app id> -PersistLogin

# Set defaults and label existing files (try -MaxFilesPerLibrary 5 first)
.\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -LabelExistingFiles -ClientId <app id> -PersistLogin

# Every visible library on one site
.\Set-DefaultLibraryLabel.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/X -AllLibraries -Label <label GUID> -ClientId <app id>

# Remove the default
.\Set-DefaultLibraryLabel.ps1 -CsvPath .\libraries.csv -Remove -ClientId <app id>
```

CSV columns: `SiteUrl, LibraryName`, and an optional `Label` that overrides `-Label` for that row. See `libraries-sample.csv`.

**Use label GUIDs.** Label names only resolve if your app can read the label list (Graph `InformationProtectionPolicy.Read`). Get GUIDs with:

```powershell
Connect-IPPSSession; Get-Label | Format-Table DisplayName, Guid, ParentId
```

A wrong GUID is rejected by SharePoint, so nothing bad gets set.

## Good to know

- **Default only = no files touched.** Without `-LabelExistingFiles` the script only changes library settings. New and edited files then get the label within a few minutes; existing files stay as they are.
- **Changing or removing a default isn't instant.** In testing, files uploaded a few seconds after a default was removed still got the old label. Microsoft quotes about 15 minutes for the tenant-wide switch (`Set-SPOTenant -DisableDocumentLibraryDefaultLabeling`); allow similar per library.

## Labelling existing files without a script

If you'd rather not re-save files, Microsoft Purview has an auto-labelling policy type that applies each library's default label to existing files. In the Purview portal go to **Information Protection > Auto-labeling policies > Create auto-labeling policy > Auto-apply to default SPO only.**, then pick sites.

- One policy of this type per tenant, up to 10 sites at a time (edit the policy to move on to the next batch).
- Only covers libraries that already have a default label, so set the defaults first (this script).
- Runs on Microsoft's schedule. There's no "run now".
- Can optionally replace lower-priority labels. Leave that off unless you mean it.
- Needs premium licensing. Check the Microsoft Purview service description.

## Side effects of `-LabelExistingFiles`

- **One new version per file.** The label is stored inside the file, so writing it always creates a version, whatever method you use. The overwrite check-in avoids a second version from the save itself.
- Power Automate flows and alerts that trigger on "file modified" may fire.
- Files that are checked out, or locked as records, are skipped or fail. They're listed in the per-file CSV.

## What was tested (demo tenant, Nov 2026)

| Test | Result |
|---|---|
| Set / re-run / remove defaults on several libraries | Read-back matches; re-run is a no-op |
| 3 libraries, 3 different labels, one CSV run | All set in 12 s; a new file in each got its own library's label within 90 s; files already there untouched |
| Default only, on a library with 417 existing files | Set in ~10 s; 5 min later all 417 files unchanged (no label, same version, same Modified) |
| Missing library, bad GUID | Clear failure per row; other rows carry on |
| `-AllLibraries` | Skips Form Templates, Site Assets, Style Library |
| Default set, existing files untouched for 15 min | **Not labelled** |
| Metadata-only edit (Title), with or without a new version | **Not labelled** |
| Plain re-save of the same content | Labelled, but +2 versions (save + label write) |
| **Check out, save, overwrite check-in (what the script does)** | **Labelled within ~40 s, +1 version, Modified and Modified By kept** |
| Graph `assignSensitivityLabel` | 402 Payment Required (metered API) |

## Test data

`test-data/generate_test_files.py` makes 27 small, realistic dummy Word, Excel, PowerPoint and PDF files (fictional property company) for a test site. Needs `openpyxl` and `reportlab`.
