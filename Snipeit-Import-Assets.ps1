<#
	.SYNOPSIS
	Updates assets using an imported CSV file.
	
	.DESCRIPTION
	Updates assets using an imported CSV file. Optionally emails results to the file's creator. Uses settings file Snipeit-Import-Assets-Settings.ps1.
	
	.OUTPUTS
	Return Codes
	-1 Error loading settings
	-2 IMPORT_CSV_FILEPATH is null or empty
	-3 Error loading Snipeit-Sync-PS API
	-4 Error connecting to Snipe-It site
	
	.NOTES
	Author: Matthew Carras
	
	Requirements:
	* SnipeItPS module: https://github.com/snazy2000/SnipeitPS
	* SnipeIt-Sync-PS.ps1
	
	How to export encrypted API credentials:
	Install-Module SnipeitPS
	Update-Module SnipeitPS
	Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
#>

# -- LOAD CONFIGURATION --
try {
	. .\Snipeit-Import-Assets-Settings.ps1
} catch {
	Write-Error $_
    return -1
}

# -- START --
$startDT = Get-Date

# Double-check the required variable is filled out in our settings file.
if([string]::IsNullOrWhitespace($IMPORT_CSV_FILEPATH)) {
	Write-Error '$IMPORT_CSV_FILEPATH is null or empty, aborting'
	return -2
}
# Check to see if the target file exists. If it doesn't, exit immediately without logging anything.
if (-Not (Test-Path $IMPORT_CSV_FILEPATH -PathType Leaf)) {
	exit 0
}

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd)"
try {
	$_logfilepath = "${_logfilepath}.log"
	Start-Transcript -Path $_logfilepath -Append
} catch {
	# If we get any error, try again with .1 appended in case it's a file lock.
	$_logfilepath = "${_logfilepath}.1.log"
	Start-Transcript -Path $_logfilepath -Append
}

# -- START FUNCTIONS --

# Works on a singular grouped asset.
function Import-Asset {
	param (
		[parameter(Mandatory=$true,
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[object]$GroupedAsset,
		
		[parameter(Position=1)]
		[string]$UserDomain,
		
		[string]$UniqueIDField='Name',

		[string]$DateFormat='yyyy-MM-dd'
	)	
	
	$asset = $null
	$sn = $GroupedAsset[0].'Serial'
	if ([string]::IsNullOrWhiteSpace($sn)) {
		Write-Warning ("[Import-Asset] [{0}] Ignoring empty serial #" -f $GroupedAsset[0].$UniqueIDField)
	} elseif ($GroupedAsset.'Serial'.Count -gt 1) {
		Throw [SnipeItSyncDuplicateNameException] ("[Import-Asset] Error importing serial # [$sn] - {0} duplicates detected [{1}]" -f $GroupedAsset.'Serial'.Count,($GroupedAsset.$UniqueIDField -join ', '))
	} else {
        $asset = [PSCustomObject]@{}
        Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Serial' -Value $sn
		$name = $GroupedAsset[0].'Name'
        Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Name' -Value $name
		$cat = $GroupedAsset[0].'Category'
		if (-Not [string]::IsNullOrEmpty($cat)) {
            Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Category' -Value $cat
			# Override fieldset with category if it exists
			$fieldset = Get-SnipeItFieldsetByName $cat -Verbose
			if ($fieldset.id -is [int]) {
                Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Fieldset' -Value $fieldset.id
			}
		}
		
		# Format datetime columns correctly
		$pdate = $GroupedAsset[0].'Purchasing date' -as [DateTime]
		if ($pdate -is [DateTime]) {
			$pdate = $pdate.ToString($DateFormat)
            Add-Member -InputObject $asset -MemberType NoteProperty -Name 'Purchasing date' -Value $pdate
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
	
# -- END FUNCTIONS --

# Load custom API
try {
    . .\SnipeIt-Sync-PS.ps1
} catch {
    # Fatal error, exit
    Write-Error $_
    return -3
}

# Initialize new Snipe-It Session
try {
    if ($CREDXML_FILEPATH -eq "snipecred-local-mcarras8.xml") {
        Connect-SnipeIt -CredXML $CREDXML_FILEPATH -IgnoreSelfSignedCert -Verbose
    } else {
        Connect-SnipeIt -CredXML $CREDXML_FILEPATH -Verbose
    }
} catch {
    # Fatal error, exit
    Write-Error $_
    return -4
}

# Initialize the snipe-it caches.
$cacheentities = @("statuslabels","manufacturers","categories","fieldsets","models","assets","fields","users")
If ($DEBUG_HALT_ON_NULL_CACHE) {
	Initialize-SnipeItCache -EntityTypes $cacheentities -ErrorOnNullEntities $cacheentities -Verbose
} else {
	Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose
}

Write-Host("[{0}] Importing assets from [{1}]..." -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")),$IMPORT_CSV_FILEPATH)
	
# Array of 
$caughtErrors = $null
$totalErrorCount = 0

# Import assets from CSV.
$imported_assets = Import-CSV -LiteralPath $IMPORT_CSV_FILEPATH | Group-Object $IMPORT_CSV_GROUP_BY | Foreach-Object {
	$asset = $null
	try {
		$asset = Import-Asset -GroupedAsset $_.Group -UserDomain $USER_DOMAIN -UniqueIDField 'Serial'
	} catch {
        Write-Error $_
		$caughtErrors += @(("[{0}] {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Exception.Message))
		$totalErrorCount++
	}
	$asset
}
# Double-check we have imported at least 1 item.
if ($imported_assets -eq $null -Or ($imported_assets -is [array] -And $imported_assets.Count -le 0) -Or ([string]::IsNullOrEmpty(($imported_assets | Select -ExpandProperty Serial -First 1)))) {
	$msg = "[{0}] An error occurred importing assets. No valid assets were imported. Aborting." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss")
	Write-Error $msg
	$caughtErrors += @( $msg )
	$totalErrorCount++
} else {
	if ($imported_assets -isnot [array]) {
		$imported_assets = @($imported_assets)
	}
	Write-Host("[{0}] {1} unique assets loaded from CSV with {2} caught errors" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), $imported_assets.Count, $totalErrorCount)

	Write-Host("[{0}] Starting sync..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
	$syncErrorCount = 0
	$uncaughtSyncErrorCount = 0
	foreach ($asset in $imported_assets) {
		if ($asset -ne $null) {
			try {
				$sp_asset = Sync-SnipeItAsset -Asset $asset -UniqueIDField "Serial" -SyncOnFieldMap $ASSET_FIELD_SYNC_ON_MAP -RequiredCreateFields $ASSET_FIELD_CREATE_REQUIRED -ErrorOnMissingFields -DefaultCreateStatus $ASSET_STATUS_CREATE -DefaultAssignedStatus $ASSET_STATUS_ASSIGNED -Verbose
				if ($sp_asset.id -isnot [int]) {
					$uncaughtSyncErrorCount++
				}
			} catch {
				Write-Error $_
				$caughtErrors += @(("[{0}] {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Exception.Message))
				$syncErrorCount++
			}
		}
	}
	$totalErrorCount += $syncErrorCount + $uncaughtSyncErrorCount
	$successSyncCount = $imported_assets.Count - $syncErrorCount

	Write-Host("[{0}] {1} of {2} assets synced successfully. Caught {3} errors during sync ({4} uncaught)" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), $successSyncCount, $imported_assets.Count, $syncErrorCount, $uncaughtSyncErrorCount)

	Write-Host("[{0}] Encountered {1} total errors" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), $totalErrorCount)

	# Get the fileinfo, as well the email address of the file's owner using RSAT tools.
	$fileInfo = Get-Item $IMPORT_CSV_FILEPATH | Select Name, Fullname, LastWriteTime, CreationTime, Length, @{N="Owner"; Expression={ (Get-Acl $_.Fullname).Owner }}
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
			$importArchiveFP = "{0}\{1}-Failure-{2}" -f $IMPORT_ARCHIVE_PATH, $startDT.ToString('yyyyMMdd-hhmmss'), $fileInfo.Name
		} else {
			$importArchiveFP = "{0}\{1}-{2}" -f $IMPORT_ARCHIVE_PATH, $startDT.ToString('yyyyMMdd-hhmmss'), $fileInfo.Name
		}
		Move-Item $IMPORT_CSV_FILEPATH -Destination $importArchiveFP -Force
		Write-Host("[{0}] Moved [{1}] to [{2}]" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), $IMPORT_CSV_FILEPATH, $importArchiveFP)
	} else {
		Remove-Item $IMPORT_CSV_FILEPATH -Force
		$importArchiveFP = $fileInfo.Fullname
	}
}

# Output all caught errors to file.
$errorLogFP = $null
if ($caughtErrors.Count -gt 0) {
	$errorLogFP = "{0}\SnipeIt-ImportFailure-{1}.log" -f $LOGFILE_PATH, $startDT.ToString('yyyyMMdd-hhmmss')
	("Import File: {0}" -f $fileInfo.Fullname) | Out-File $errorLogFP
	$caughtErrors | Out-File -Append $errorLogFP
    Write-Host("[{0}] Outputting copy of error log to [{1}]" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), $errorLogFP)
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
"@ -f $successSyncCount, $totalErrorCount, $startDT.ToString($dateFormat), $endDT.ToString($dateFormat), $elapsed, $importArchiveFP, $fileInfo.Owner, $fileInfo.Length, $fileInfo.CreationTime.ToString($dateFormat), $fileInfo.LastWriteTime.ToString($dateFormat)
	
	try {
		Send-MailMessage @params -ErrorAction Stop
	} catch {
		# Try again without attachment.
		Write-Error $_
		$params['Attachments'] = $null
		Send-MailMessage @params
	}
    Write-Host("[{0}] Emailed report to [{1}]" -f ((Get-Date -Format "yyyy/MM/dd HH:mm:ss")), ($emailReportTo -join ", "))
}

Write-Host("[0] Errors encountered: {1}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $errorCount)

$runtimeDiff = ((Get-Date) - $startDT)
Write-Host("[{0}] Total Runtime: {1} hours {2} minutes ({3} total minutes)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $runtimeDiff.Hours, $runtimeDiff.Minutes, $runtimeDiff.TotalMinutes)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
