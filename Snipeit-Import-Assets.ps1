# Updates assets using an imported CSV file.
# Each column name must match the ones accepted by Sync-SnipeItAsset,
# which is generally all fields accepted by the snipe-it API, plus a few others.
# See Get-Help Sync-SnipeItAsset for more info on accepted fields.
#
# Requirements:
# * SnipeItPS module: https://github.com/snazy2000/SnipeitPS
# * SnipeIt-Sync-PS.ps1: https://github.com/mattcarras/SnipeItSyncPS
# * RSAT Tools for emailing file owner a success/failure report.
#
# Install-Module SnipeitPS
# Update-Module SnipeitPS
# Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
#
# Author: Matthew Carras
# Source: https://github.com/mattcarras/SnipeItSyncPS

# -- START CONFIGURATION --
# Previously exported credentials
$CREDXML_PATH = "your_exported_credentials.xml"

# Direct filepath to import assets.
$IMPORT_CSV_PATH = "SnipeIt-Import.csv"
# A unique identifier to group the results by.
$IMPORT_CSV_GROUP_BY = "Serial"

# If no fieldset is defined when importing an asset, but there's a valid fieldset with the same name as the category, use that fieldset if needed.
$IMPORT_USE_CATEGORY_MISSING_FIELDSET = $true

# Which fields to sync if the given field matches.
# A value of $true syncs ALL mapped fields.
# Make sure these fields are all unique!
# Defaults to @{ Serial = $true }
$ASSET_FIELD_SYNC_ON_MAP = [ordered]@{
    "Serial" = $true
}

# Which fields are required to be non-blank to create the asset.
# Only checked when creating new assets.
# This is required for full functionality.
$ASSET_FIELD_CREATE_REQUIRED = @("Serial","Model","Manufacturer","Category")

# Status used when creating a new unassigned asset. This can be the status name or ID.
$ASSET_STATUS_CREATE = "Pending"
# Status used when assigning assets, including newly created assets.
$ASSET_STATUS_ASSIGNED = "Ready to Deploy"

# Unique ID field to use in output logs. Defaults to 'Serial'.
$ASSET_FIELD_UNIQUE_ID = 'Serial'

# Throw an error on missing fields rather than just outputting to Write-Verbose.
$ASSET_SYNC_ERROR_MISSING_FIELDS = $true

# To make doubly sure we aren't duplicating any entities, halt if the list of assets, models, manufacturers, fieldsets, categories, etc. are empty
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $true

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "snipeit-import-assets"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Path to copies of imported files.
# If set, imported file will be copied to the given destination and renamed with suffix
# of -Imported
$IMPORT_ARCHIVE_PATH = ".\Imported"
# Maximum number of days worth of imported files to keep.
$IMPORT_ARCHIVE_ROTATE_DAYS = 365

# Email configuration for reports.
# Output from this script will be emailed to the file owner if set.
# Report will contain a truncated list of any encountered errors.
<#
# If filled out, send reports to file owner.
$EMAIL_SMTP = '<smtp server>'
$EMAIL_REPORT_FROM = '<from address>'
# May include multiple destination addresses as an array.
$EMAIL_REPORT_CC = '<CC address>'
$EMAIL_REPORT_BCC = '<BCC address>'
# If NOT filed out, report will be emailed to file owner.
# $EMAIL_REPORT_TO = '<to address>'
#>
# -- END CONFIGURATION --

# -- START --
$startDT = Get-Date

# Check to see if the target file exists. If it doesn't, exit immediately without logging anything.
if (-Not (Test-Path $IMPORT_CSV_PATH -PathType Leaf)) {
    return 0
}

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
    Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd-hhss).log"
Start-Transcript -Path $_logfilepath -Append

# -- START FUNCTIONS --

# Works on a singular grouped asset.
function Import-Asset {
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object]$GroupedAsset,
        
        [parameter(Mandatory=$false)]
        [string]$UniqueIDField='Name',

        [parameter(Mandatory=$false)]
        [string]$DateFormat='yyyy-MM-dd'
    )   
    Begin {
    }
    Process {
        $asset = $null
        $$asset = $null
        $uid = $GroupedAsset[0].$UniqueIDField
        if ([string]::IsNullOrWhiteSpace($sn)) {
            Write-Warning ("[Import-Asset] [{0}] Ignoring empty serial #" -f $GroupedAsset[0].$UniqueIDField)
        } elseif ($GroupedAsset.$UniqueIDField.Count -gt 1) {
            Throw [SnipeItSyncDuplicateNameException] ("[Import-Asset] Error importing [$UniqueIDField] = [$uid] - {0} duplicates detected [{1}]" -f $GroupedAsset.$UniqueIDField.Count,($GroupedAsset.$UniqueIDField -join ', '))
        } else {
            $asset = [PSCustomObject]@{}
            $sn = $GroupedAsset[0].'Serial'
            Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Serial' -Value $sn
            $name = $GroupedAsset[0].'Name'
            Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Name' -Value $name
            $cat = $GroupedAsset[0].'Category'
            if (-Not [string]::IsNullOrEmpty($cat)) {
                Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Category' -Value $cat
                if ([string]::IsNullOrEmpty($GroupedAsset[0].'Fieldset') -And $IMPORT_USE_CATEGORY_MISSING_FIELDSET) {
                    $fieldset = Get-SnipeItFieldsetByName $cat -Verbose
                    if ($fieldset.id -is [int]) {
                        Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Fieldset' -Value $fieldset.id
                    }
                }
            }
            
            foreach ($prop in ($GroupedAsset[0] | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name | Out-String -Stream)) {
                if (-Not [string]::IsNullOrEmpty($prop)) {
                    $val = $GroupedAsset[0].$prop
                    if (-Not [string]::IsNullOrEmpty($val) -And [string]::IsNullOrEmpty($asset.$prop)) {
                        Add-Member -InputObject $asset -MemberType NoteProperty -Name $prop -Value $val -Force
                    }
                }
            }
        }
        return $asset
    }
    End {
    }
}
    
# -- END FUNCTIONS --

# Load custom API
try {
    . .\SnipeIt-Sync-PS.ps1
} catch {
    # Fatal error, exit
    Write-Error $_
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    return -1
}

# Initialize new Snipe-It Session
try {
    Connect-SnipeIt -CredXML $CREDXML_PATH -Verbose
} catch {
    # Fatal error, exit
    Write-Error $_
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    return -2
}

# Initialize the snipe-it caches.
$cacheentities = @("statuslabels","manufacturers","categories","fieldsets","models","assets","fields","users")
If ($ASSET_FIELD_MAP.ContainsKey("company") -Or $ASSET_FIELD_MAP.ContainsKey("company_id")) {
    $cacheentities += @("companies")
}
If ($ASSET_FIELD_MAP.ContainsKey("location") -Or $ASSET_FIELD_MAP.ContainsKey("location_id") -Or $ASSET_FIELD_MAP.ContainsKey("rtd_location_id")) {
    $cacheentities += @("locations")
}
If ($ASSET_FIELD_MAP.ContainsKey("supplier") -Or $ASSET_FIELD_MAP.ContainsKey("supplier_id")) {
    $cacheentities += @("suppliers")
}
$extraParams = @{}
If ($DEBUG_HALT_ON_NULL_CACHE) {
    $extraParams.Add("ErrorOnNullEntities", $cacheentities)
}
Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose @extraParams

Write-Host("[{0}] Importing assets from [{1}]..." -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")),$IMPORT_CSV_PATH)
    
# Array of 
$errorMsgs = $null
$totalErrorCount = 0

# Import assets from CSV.
$imported_assets = Import-CSV -LiteralPath $IMPORT_CSV_PATH | Group-Object $IMPORT_CSV_GROUP_BY | Foreach-Object {
    $asset = $null
    try {
        $asset = Import-Asset -GroupedAsset $_.Group -UniqueIDField 'Serial'
    } catch {
        Write-Error $_
        $errorMsgs += @(("[{0}] {1}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_.Exception.Message))
        $totalErrorCount++
    }
    $asset
}
if ($imported_assets -isnot [array]) {
    $imported_assets = @($imported_assets)
}
Write-Host("[{0}] {1} unique assets loaded from CSV with {2} caught errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $imported_assets.Count, $totalErrorCount)

Write-Host("[{0}] Starting sync..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
$syncErrorCount = 0
$uncaughtSyncErrorCount = 0
$extraParams = @{}
<#
if (-Not [string]::IsNullOrWhitespace($ASSET_DEFAULT_MODEL)) {
    $extraParams.Add("DefaultModel", $ASSET_DEFAULT_MODEL)
}
#>
if ($ASSET_FIELD_SYNC_ON_MAP.Count -gt 0) {
    $extraParams.Add('SyncOnFieldMap', $ASSET_FIELD_SYNC_ON_MAP)
}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_CREATE)) {
    $extraParams.Add('DefaultCreateStatus', $ASSET_STATUS_CREATE)
}
if (-Not [string]::IsNullOrWhitespace($ASSET_ARCHIVED_STATUS_UPDATE)) {
    $extraParams.Add('UpdateArchivedStatus', $ASSET_ARCHIVED_STATUS_UPDATE)
}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_ASSIGNED)) {
    $extraParams.Add('DefaultAssignedStatus', $ASSET_STATUS_ASSIGNED)
}
if (-not [string]::IsNullOrEmpty($ASSET_FIELD_UNIQUE_ID)) {
    $extraParams.Add('UniqueIDField', $ASSET_FIELD_UNIQUE_ID)
}
foreach ($asset in $imported_assets) {
    if ($asset -ne $null) {
        try {
            $sp_asset = Sync-SnipeItAsset -Asset $asset -RequiredCreateFields $ASSET_FIELD_CREATE_REQUIRED -ErrorOnMissingFields:$ASSET_SYNC_ERROR_MISSING_FIELDS -Verbose @extraParams
            if ($sp_asset.id -isnot [int]) {
                $uniqueID = $null
                if (-Not [string]::IsNullOrEmpty($ASSET_FIELD_UNIQUE_ID)) {
                    $uniqueID = $asset.$ASSET_FIELD_UNIQUE_ID
                }
                if ([string]::IsNullOrEmpty($uniqueID)) {
                    $uniqueID = $asset.'Serial'
                }
                # Should always return a valid ID
                $errorMsgs += @(("[{0}] Asset [$uniqueID] did not sync successfully" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
                $uncaughtSyncErrorCount++
            }
        } catch {
            Write-Error $_
            $errorMsgs += @(("[{0}] {1}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_.Exception.Message))
            $syncErrorCount++
        }
    }
}
$totalErrorCount += $syncErrorCount + $uncaughtSyncErrorCount
$successSyncCount = $imported_assets.Count - $syncErrorCount

Write-Host("[{0}] {1} of {2} assets synced successfully. Caught {3} errors during sync ({4} uncaught/invalid IDs returned)" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $successSyncCount, $imported_assets.Count, $syncErrorCount, $uncaughtSyncErrorCount)

Write-Host("[{0}] Encountered {1} total errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $totalErrorCount)

# Get the fileinfo, as well the email address of the file's owner using RSAT tools.
$fileInfo = Get-Item $IMPORT_CSV_PATH | Select Name, Fullname, LastWriteTime, CreationTime, Length, @{N="Owner"; Expression={ (Get-Acl $_.Fullname).Owner }}
$ownerEmail = $null
if (-Not [string]::IsNullOrEmpty($fileInfo.Owner) -And $fileInfo.Owner -match '\\?([^"\[\]:;|=+*?<>/\\]{1,61})$' -And -Not [string]::IsNullOrEmpty($matches[1])) {
    $ownerEmail = Get-ADUser -Identity $matches[1] -Properties mail | Select -ExpandProperty mail
}
$emailReportTo = $ownerEmail
if ([string]::IsNullOrEmpty($emailReportTo)) {
    $emailReportTo = $EMAIL_REPORT_CC
}

# Move a copy of the import file to $IMPORT_ARCHIVE_PATH or delete it
$importArchiveFP = $null
if (Test-Path $IMPORT_ARCHIVE_PATH -PathType Container) {
    if ($totalErrorCount -gt 0) {
        $importArchiveFP = "{0}\{1}-{2}-{3}" -f $IMPORT_ARCHIVE_PATH, $startDT.ToString('yyyyMMdd-hhmmss'), $IMPORT_ARCHIVE_FAILURE_SUFFIX, $fileInfo.Name
    } else {
        $importArchiveFP = "{0}\{1}-{2}" -f $IMPORT_ARCHIVE_PATH, $startDT.ToString('yyyyMMdd-hhmmss'), $fileInfo.Name
    }
    Move-Item $IMPORT_CSV_PATH -Destination $importArchiveFP -Force
    Write-Host("[{0}] Moved [{1}] to [{2}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $IMPORT_CSV_PATH, $importArchiveFP)
} else {
    Remove-Item $IMPORT_CSV_PATH -Force
}

# Output all caught errors to file.
$errorLogFP = $null
if ($errorMsgs.Count -gt 0) {
    $errorLogFP = "{0}\SnipeIt-ImportFailure-{1}.log" -f $LOGFILE_PATH, $startDT.ToString('yyyyMMdd-hhmmss')
    ("Import File: {0}" -f $fileInfo.Fullname) | Out-File $errorLogFP
    $errorMsgs | Out-File -Append $errorLogFP
    Write-Host("[{0}] Outputting copy of error log to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $errorLogFP)
}

# Email out report.
if (-Not [string]::IsNullOrEmpty($EMAIL_SMTP) -And -Not [string]::IsNullOrEmpty($EMAIL_REPORT_FROM) -And -Not [string]::IsNullOrEmpty($emailReportTo)) {
    $params = @{
        From = $EMAIL_REPORT_FROM
        To = $emailReportTo
        Subject = 'Snipe-It Import Success'
        DeliveryNotificationOption = @('OnSuccess', 'OnFailure')
        SmtpServer = $EMAIL_SMTP
        BodyAsHtml = $true
    }
    if (-Not [string]::IsNullOrEmpty($EMAIL_REPORT_CC)) {
        $params['Cc'] = $EMAIL_REPORT_CC
    }
    if (-Not [string]::IsNullOrEmpty($EMAIL_REPORT_BCC)) {
        $params['Bcc'] = $EMAIL_REPORT_BCC
    }
    
    # On error, change subject and add copy of error log along with failed import file.
    if ($totalErrorCount -gt 0) {
        $params['Subject'] = 'Snipe-It Import Failure'
        $params['Priority'] = 'High'
    
        $attachments = $null
        if (Test-Path $errorLogFP -PathType Leaf) {
            $attachments += @($errorLogFP)
        }
        if (Test-Path $importArchiveFP -PathType Leaf) {
            $attachments += @($importArchiveFP)
        }
        if ($attachments -ne $null) {
            $params['Attachments'] = $attachments
        }
    }
    
    $dateFormat = 'MM/dd/yyyy HH:MM:ss'
    $endDT = Get-Date
    $elapsedTS = New-TimeSpan -Start $startDT -End $endDT
    $elapsed = "{0} Hours {1} Minutes {2} Seconds" -f $elapsedTS.Hours, $elapsedTS.Minutes, $elapsedTS.Seconds
    $params['body'] = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii"><title>HTML TABLE</title>
</head><body>
<table>
<tr><td>Snipe-It Asset Import Success:</td><td>{0}</td></tr>
<tr><td>Snipe-It Asset Import Failure:</td><td>{1}</td></tr>
<tr><td>Start Date:</td><td>{2}</td></tr>
<tr><td>Finish Date:</td><td>{3}</td></tr>
<tr><td>Total Time:</td><td>{4}</td></tr>
<tr><td>CSV File Name:</td><td><a href="file://{5}">{5}</a></td></tr>
<tr><td>CSV File Owner:</td><td>{6}</td></tr>
<tr><td>CSV File Size:</td><td>{7}</td></tr>
<tr><td>CSV Date Created:</td><td>{8}</td></tr>
<tr><td>CSV Last Modified:</td><td>{9}</td></tr>
</table>
</body></html>
"@ -f $successSyncCount, $totalErrorCount, $startDT.ToString($dateFormat), $endDT.ToString($dateFormat), $elapsed, $fileInfo.Fullname, $fileInfo.Owner, $fileInfo.Length, $fileInfo.CreationTime.ToString($dateFormat), $fileInfo.LastWriteTime.ToString($dateFormat)
    
    try {
        Send-MailMessage @params
    } catch {
        # Try again without attachment.
        Write-Error $_
        $params['Attachments'] = $null
        Send-MailMessage @params
    }
    Write-Host("[{0}] Emailed report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($emailReportTo -join ", "))
}

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
