<#
	.SYNOPSIS
	Disables assets that match criteria for "Assigned User Deleted".
	
	.DESCRIPTION
	Disables assets that match criteria for "Assigned User Deleted" (user deleted from sync). Loads settings from Snipeit-Asset-Sync-Settings.ps1 and Snipeit-Disable-DeletedUser-Systems-Settings.ps1.
	
	.PARAMETER PendingReport
	Optional. Create a report of pending actions only, which can then be processed by -ProcessPendingReport.
	
	.PARAMETER ProcessPendingReport
	Optional. Process a previously exported pending report from -PendingReport.
	
	.PARAMETER EmailReport
	Optional. Email a report after processing.
	
	.PARAMETER NextRunDays
	Optional. The number of days before the next run. Default is 14. Used for reporting.
	
	.PARAMETER LogFilePrefix
	Optional. The prefix used for log files. Defaults to "snipeit-disable-deleteduser-systems".
	
	.PARAMETER DryRun
	Optional. Only output to log, do not process or email. Useful for debugging.
	
	.OUTPUTS
	Return Codes
	-1 Error loading settings files
	-2 -PendingReport given but EXPORT_DISABLED_COMPUTERS_PENDING_PATH is blank
	-3 Error loading Snipeit-Sync-PS API
	-4 Error connecting to Snipe-It site
	-5 Error getting all Snipe-It assets
	-6 No assets returned from Snipe-It
	-7 Constructed filterScript from ASSET_DISABLE_SP_CRITERIA is blank
	-8 -ProcessPendingReport given, but no pending report was found at [$EXPORT_DISABLED_COMPUTERS_PENDING_PATH]
	-9 Unknown error processing filterscript
	-10 Too many matching assets to disable (ASSET_DISABLE_SP_CRITERIA matches all assets)
	-11 No valid results from AD
	
	.NOTES
	Author: Matthew Carras
	
	Requirements:
	* RSAT: Active Directory PowerShell module if you're getting results from AD.
	* SnipeItPS module: https://github.com/snazy2000/SnipeitPS
	* SnipeIt-Sync-PS.ps1
	
	Use -PendingReport and -ProcessPendingReport to give your admin team warning of pending actions.
	The script will check Snipe-It again for each system in the pending report for any updates.
	
	How to export encrypted API credentials:
	Install-Module SnipeitPS
	Update-Module SnipeitPS
	Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
#>
param([switch] $PendingReport, 
	  [switch] $ProcessPendingReport,
	  [switch] $EmailReport,
	  [int] $NextRunDays=14, 
	  [string] $LogFilePrefix, 
	  [switch] $DryRun
)

# -- LOAD CONFIGURATION --
try {
	# Load shared configuration
	. .\Snipeit-Asset-Sync-Settings.ps1
	# Load main configuration
	. .\Snipeit-Disable-DeletedUser-Systems-Settings.ps1
} catch {
	Write-Error $_
    return -1
}

# -- FUNCTIONS START --

function Email-ErrorReport {
	<#
		.SYNOPSIS
		Emails out an error report if we have any errors to report.
		
		.DESCRIPTION
		Emails out an error report if we have any errors to report.
		
		.PARAMETER To
		Required. Report recipients as an array.
		
		.PARAMETER From
		Required. Email address to send as.
		
		.PARAMETER SmtpServer
        Required. The email server to use.
		
		.PARAMETER ErrorCount
		Required. Number of errors to report.
		
		.PARAMETER LogFilePath
		Required. Path to the logfile to link or attach.
		
	#>
	param(		
		[parameter(Mandatory=$true)]
		[string[]]$To,
		[parameter(Mandatory=$true)]
		[string]$From,
		[parameter(Mandatory=$true)]
		[string]$SmtpServer,
		[parameter(Mandatory=$true)]
		[int]$ErrorCount,
		[parameter(Mandatory=$true)]
		[string]$LogFilePath
	)
			
	$_scriptName = split-path $PSCommandPath -Leaf
	
	# Email out notifications of any errors.
	if ($ErrorCount -gt 0 -And -Not [string]::IsNullOrWhiteSpace($SmtpServer) -And -Not [string]::IsNullOrWhiteSpace($From) -And -Not [string]::IsNullOrWhiteSpace(($To | Select -First 1))) {
		$emailParams = @{
			From = $From
			To =  $To
			Subject = "Errors from $_scriptName"
			Body = "There were [$ErrorCount] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
			Priority = "High"
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $SmtpServer
		}
		
		# Stop logging
		Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

		try {
			Send-MailMessage -Attachments $LogFilePath @emailParams -ErrorAction Stop
		} catch {
			Write-Error $_
			$mailParams.Body = "There were [$ErrorCount] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$LogFilePath] for more details."
			Send-MailMessage @emailParams
		}
		Write-Verbose("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($To -join ", "))
	}
}

# -- FUNCTIONS END --

# -- START --
$dateStart = Get-Date
$_scriptName = split-path $PSCommandPath -Leaf

$_logfileprefix = $LOGFILE_PREFIX
if (-Not [string]::IsNullOrWhitespace($LogFilePrefix)) {
    $_logfileprefix = $LogFilePrefix
} else {
    $_logfileprefix = $LOGFILE_PREFIX
}

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${_logfileprefix}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${_logfileprefix}_$(get-date -f yyyy-MM-dd)"
try {
	$_logfilepath = "${_logfilepath}.log"
	Start-Transcript -Path $_logfilepath -Append
} catch {
	# If we get any error, try again with .1 appended in case it's a file lock.
	$_logfilepath = "${_logfilepath}.1.log"
	Start-Transcript -Path $_logfilepath -Append
}

if (($EMAIL_DISABLED_COMPUTERS_REPORT -Or $EmailReport) -And -Not [string]::IsNullOrWhitespace($EMAIL_DISABLED_COMPUTERS_REPORT_TO_GROUPMEMBERS)) {
	Write-Host('[{0}] Compiling list of disabled computer report recipients from [{1}].' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $EMAIL_DISABLED_COMPUTERS_REPORT_TO_GROUPMEMBERS)
	$emailReportTo = Get-ADGroupMember $EMAIL_DISABLED_COMPUTERS_REPORT_TO_GROUPMEMBERS -Recursive | foreach { Get-ADUser $_ -Properties mail | Select -ExpandProperty mail }
} else {
	$emailReportTo = $EMAIL_DISABLED_COMPUTERS_REPORT_TO
}

If ($PendingReport -And [string]::IsNullOrWhitespace($EXPORT_DISABLED_COMPUTERS_PENDING_PATH)) {
	Write-Error "-PendingReport given but EXPORT_DISABLED_COMPUTERS_PENDING_PATH is blank. Exiting."
    return -2
}

# Load custom API
try {
    . .\SnipeIt-Sync-PS.ps1
} catch {
    # Fatal error, exit
    Write-Error $_
	$error_count++
	If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
		Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
	}
    return -3
}

# Get the next run date if we're set to do a pending report.
$nextRun = $null
if ($PendingReport -And $NextRunDays) {
	$nextRun = ((Get-Date).AddDays($NextRunDays)).ToString("yyyy-MM-dd")
}

# Initialize new Snipe-It Session
$spHostURL = $null
try {
    Connect-SnipeIt -CredXML $CREDXML_PATH -Verbose
    # Used for reports.
    $spHostURL = (Import-CliXml $CREDXML_PATH).Username
    if (-Not $spHostURL.EndsWith('/')) {
        $spHostURL += '/'
    }
} catch {
    # Fatal error, exit
    Write-Error $_
	If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
		Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
	}
    return -4
}

# Initialize the snipe-it caches.
#$cacheentities = @("statuslabels","manufacturers","categories","fieldsets","models","assets","fields")
$cacheentities = @("assets")

If ($DEBUG_HALT_ON_NULL_CACHE) {
	Initialize-SnipeItCache -EntityTypes $cacheentities -ErrorOnNullEntities $cacheentities -Verbose
} else {
	Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose
}

$syncExtraParams = @{}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_ARCHIVED_UPDATE)) {
    $syncExtraParams.Add('UpdateArchivedStatus', $ASSET_STATUS_ARCHIVED_UPDATE)
}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_ASSIGNED)) {
	$syncExtraParams.Add('DefaultAssignedStatus', $ASSET_STATUS_ASSIGNED)
}
	
$error_count = 0
$sp_assets_count = 0
$sp_assets_count_found = 0
$modified_assets = [System.Collections.Generic.List[object]]::new()
$sp_assets_unfiltered = $null
$sp_assets = $null
$ad_assets = $null

Write-Host("[{0}] Searching assets..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))

# Get all Snipe-It assets.
try {
	$sp_assets_unfiltered = Get-SnipeItEntityAll "assets" -ReturnValues
} catch {
	Write-Error $_
	$error_count++
	If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
		Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
	}
	return -5
}

if ($sp_assets_unfiltered -eq $null) {
	Write-Error "No assets returned from Snipe-It. Exiting."
	$error_count++
	If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
		Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
	}
	return -6
}

# Construct a filterscript from the disable criteria.
$filterScript = $null
foreach ($criteriaInfo in $ASSET_DISABLE_SP_CRITERIA) {
	$f = ""
	if(-Not [string]::IsNullOrEmpty($criteriaInfo.Field)) {
		if($criteriaInfo.Exclude) {
			$f += "-Not ("
		}
		if($criteriaInfo.IsCustom) {
			$f += "`$_.custom_fields.'$($criteriaInfo.Field)'.value"
		} elseif ($criteriaInfo.SubField) {
			$f += "`$_.'$($criteriaInfo.Field)'.$($criteriaInfo.SubField)"
		} else {
			$f += "`$_.'$($criteriaInfo.Field)'"
		}
		if ($criteriaInfo.ValueMatch) {
			$f += " -match '$($criteriaInfo.Value)'"
		} else {
			$f += " -eq '$($criteriaInfo.Value)'"
		}
		if($criteriaInfo.Exclude) {
			$f += ")"
		}
	}
	if ($filterScript -eq $null) {
		$filterScript = $f
	} else {
		$filterScript += " -AND $f"
	}
}
Write-Host("[{0}] ASSET_DISABLE_SP_CRITERIA filterScript=$filterScript" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))

if ([string]::IsNullOrWhitespace($filterScript)) {
	Write-Error "Constructed filterScript from ASSET_DISABLE_SP_CRITERIA is blank, cannot continue"
	$error_count++
	If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
		Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
	}
	return -7
}

$sp_assets_count = $sp_assets_unfiltered | Measure-Object | Select -ExpandProperty Count
$sp_assets_count_found = $null
if ($ProcessPendingReport) {
	Write-Host ("[{0}] Processing previously saved pending report from [$EXPORT_DISABLED_COMPUTERS_PENDING_PATH]..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
	
	if (-Not (Test-Path $EXPORT_DISABLED_COMPUTERS_PENDING_PATH -PathType Leaf)) {
		Write-Error "-ProcessPendingReport given, but no pending report was found at [$EXPORT_DISABLED_COMPUTERS_PENDING_PATH]"
		$error_count++
		If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
			Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
		}
		return -8
	} else {
		# Load assets from report, then load their information from Snipe-It by the saved ID in the report.
		$pending_assets = Import-CSV $EXPORT_DISABLED_COMPUTERS_PENDING_PATH
		if ($pending_assets -ne $null) {
			$sp_assets = $sp_assets_unfiltered | where {$_.id -in $pending_assets."SnipeIt ID" -and -not [string]::IsNullOrEmpty($_.id)}
		}
		$sp_assets_count_found = $sp_assets | Measure-Object | Select -ExpandProperty Count
		Write-Host("[{0}] [{1}] out of [{2}] total assets filtered based on pending report" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_assets_count_found, $sp_assets_count)
	}
} else {
	Write-Host ("[{0}] Filtering snipe-it assets based on given criteria." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))

	try {
		$scriptBlock = [ScriptBlock]::Create($filterScript)
		$sp_assets = $sp_assets_unfiltered | Where-Object -FilterScript $scriptBlock
	} catch {
		Write-Error $_
		$error_count++
		If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
			Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
		}
		return -9
	}
	
	$sp_assets_count_found = $sp_assets | Measure-Object | Select -ExpandProperty Count
	Write-Host("[{0}] [{1}] out of [{2}] total assets matched criteria" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_assets_count_found, $sp_assets_count)
}

if ($sp_assets_count_found -le 0) {
	Write-Host("[{0}] No assets match criteria. Exiting.")
} else {
	if ($sp_assets_count_found -eq $sp_assets_count) {
		Write-Warning("[{0}] Too many assets match criteria. Exiting.")
		$error_count++
		If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
			Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
		}
		return -10
	}
	
	Write-Host("[{0}] Collecting all assets from AD using searchbase [{1}], this might take a while..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"),($SEARCHBASE_OU -join ", "))
	try {
		$ad_assets = Get-ADComputer -Filter "*" -Searchbase $SEARCHBASE_OU -Properties distinguishedname
		Write-Host("[{0}] {1} assets imported from AD" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $ad_assets.Count)
	} catch {
		Write-Error $_
		$error_count++
	}
	
	# Check if we have at least one valid ad asset.
	if ([string]::IsNullOrEmpty(($ad_assets | Select -ExpandProperty distinguishedname))) {
		Write-Error("No valid assets returned from AD")
		$error_count++
		If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
			Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
		}
		return -11
	}
}

# Loop over every snipe-it asset that matches criteria.
$actionCounter = 0
foreach ($sp_asset in $sp_assets) {
	if([string]::IsNullOrWhitespace($sp_asset.assigned_to.username)) {
		Write-Host("[{0}] Skipping empty assigned_to username for name=[{1}], id=[{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_asset.name, $sp_asset.id)
	} elseif(-not $ASSET_CHECK_USER_DELETED_INCLUDE_UPNS -And $sp_asset.assigned_to.username -like "*@*") {
		Write-Host("[{0}] Skipping name=[{1}], id=[{2}], assigned_to=[{3}] due to being assigned to a valid UPN" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_asset.name, $sp_asset.id, $sp_asset.assigned_to.username)
	} elseif($ProcessPendingReport -Or -Not $ASSET_DISABLE_THRESHOLD_MAX -Or $actionCounter -lt $ASSET_DISABLE_THRESHOLD_MAX) {
		$compFound = $false
		$actionTaken = $null
		$comp = $null
		foreach($fieldinfo in $ASSET_CHECK_USER_DELETED_AD_MATCH_FIELDMAP) {
			$adField = $fieldinfo.ADField
			$spField = $fieldinfo.SnipeItField
			if (-Not [string]::IsNullOrEmpty($adField) -And -Not [string]::IsNullOrEmpty($spField)) {
				if ($fieldinfo.SnipeItFieldIsCustom) {
					$sp_value = $sp_asset.custom_fields.$spField.Value
				} else {
					$sp_value = $sp_asset.$spField
				}
				if(-Not [string]::IsNullOrEmpty($sp_value)) {	
					if ($adField -eq "SID") {
						$comp = $ad_assets | where {[string]($_.$adField.Value) -eq $sp_value}
					} else {
						$comp = $ad_assets | where {[string]($_.$adField) -eq $sp_value}
					}
					if ($comp.Count -gt 1) {
						Write-Warning("[{0}] Returned $($comp.Count) results checking for $adField=[$sp_value], skipping: name=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_asset.name, $sp_asset.id, $sp_asset.assigned_to.username)
					} else {
						if (-not [string]::IsNullOrEmpty($comp.distinguishedname)) {
							$compFound = $true
							$compOU = $null
							# Get OU from distinguishedname
							if ($comp.distinguishedname -match "^CN=[^,]+,(OU=.+)") {
								$compOU = $Matches[1]
							}
							if ([string]::IsNullOrWhitespace($compOU) -Or $compOU -in $EXCLUDED_COMPUTER_OUS) {
								Write-Host("[{0}] Skipping due to excluded OU DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
								$actionTaken = "Skipped (Excluded)"
							} else {
								if ($comp.Enabled) {
									
									try {
										if ($PendingReport) {
											Write-Host("[{0}] Would disable DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
										} elseif ($DryRun) {
											Write-Host("[{0}] Would disable (dryrun) DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
										} else {
											Write-Host("[{0}] Disabling DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
											Disable-ADAccount -Identity $comp.distinguishedname
										}
										$actionTaken = "Disabled"
									} catch {
										Write-Error $_
										$error_count++
									}
								}
								if ($comp.distinguishedname -ne "CN=$($comp.name),$DISABLED_COMPUTER_OU") {
									try {
										if ($PendingReport) {
											Write-Host("[{0}] Would move DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
										} elseif ($DryRun) {
											Write-Host("[{0}] Would move (dryrun) DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
										} else {
											Write-Host("[{0}] Moving DN=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.distinguishedname, $sp_asset.id, $sp_asset.assigned_to.username)
											Move-ADObject -Identity $comp.distinguishedname -TargetPath $DISABLED_COMPUTER_OU
										}
										if ($actionTaken -ne $null) {
											$actionTaken += ",Moved"
										} else {
											$actionTaken = "Moved"
										}
									} catch {
										Write-Error $_
										$error_count++
									}
								}
							}
							break
						}
					}
				}
			}
		}
		if (-not $compFound) {
			Write-Host("[{0}] System not found in AD, skipping: name=[{1}], id=[{2}], assigned_to=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $sp_asset.name, $sp_asset.id, $sp_asset.assigned_to.username)
			If ($ASSET_REPORT_NOT_FOUND_IN_AD) {
				$actionTaken = "Skipped (Not Found)"
			} else {
				$actionTaken = $null
			}
		}
		if ($actionTaken -ne $null) {
			$o = [PSCustomObject]@{
				"Name"=$comp.Name
				"SnipeIt ID"=$sp_asset.id
				"Link"=$spHostURL + "hardware/" + $sp_asset.id
			}
			# Add additional fields from Snipe-It if set.
			foreach ($f in $EXPORT_DISABLED_COMPUTERS_EXTRA_FIELDS) {
				if ($f.IsCustom) {
					$v = $sp_asset.custom_fields.($f.field).value
				} elseif ($f.SubField) {
					$v = $sp_asset.($f.Field).($f.SubField)
				} else {
					$v = $sp_asset.($f.Field)
				}
					
				Add-Member -InputObject $o -NotePropertyName $f.Field -NotePropertyValue $v
			}
			if($PendingReport) {
				Add-Member -InputObject $o -NotePropertyName 'Pending Action' -NotePropertyValue $actionTaken
			} else {
				Add-Member -InputObject $o -NotePropertyName 'Action' -NotePropertyValue $actionTaken
			}
			
			$modified_assets.Add($o)
			
			# Exit out if max action threshold has been reached.
			$actionCounter++
			if(-Not $ProcessPendingReport -And $ASSET_DISABLE_THRESHOLD_MAX -And $actionCounter -ge $ASSET_DISABLE_THRESHOLD_MAX) {
				Write-Host("[{0}] Max action threshold of [$ASSET_DISABLE_THRESHOLD_MAX] reached. Processing halted." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
			}
		}
	}
}

$modified_count = ($modified_assets | where {$_.Action -notmatch "Skipped"} | Measure-Object).Count
$modified_excluded_count = ($modified_assets | where {$_.action -match "Skipped"} | Measure-Object).Count
Write-Host("[{0}] {1} computers modified, not including {2} skipped." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $modified_count, $modified_excluded_count)
$modified_total_count = ($modified_assets | Measure-Object).Count

# Export a pending report. This report may be blank.
if ($PendingReport) {
	try {
		if ($modified_total_count -eq $null -Or $modified_total_count -le 0) {
			# If no assets modified, clear the pending report.
			Clear-Content -Force $EXPORT_DISABLED_COMPUTERS_PENDING_PATH
		} else {
			$modified_assets | Export-CSV -NoTypeInformation -Force $EXPORT_DISABLED_COMPUTERS_PENDING_PATH
		}
	} catch {
		Write-Error $_
		Write-Host('[{0}] Waiting 600 seconds and trying again...' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
		Start-Sleep -Seconds 600
		$modified_assets | Export-CSV -NoTypeInformation -Force $EXPORT_DISABLED_COMPUTERS_PENDING_PATH
	}
	if (Test-Path $EXPORT_DISABLED_COMPUTERS_PENDING_PATH -PathType Leaf) {
		Write-Host('[{0}] Pending disabled computer report has been saved to [{1}]. Next Run: {2}' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $EXPORT_DISABLED_COMPUTERS_PENDING_PATH, $nextRun)
    }
# Export a list of actions taken.
} elseif(-not [string]::IsNullOrWhitespace($EXPORT_DISABLED_COMPUTERS_REPORT_PATH) -And $modified_total_count -gt 0) {
	try {
		$modified_assets | Export-CSV -NoTypeInformation -Force $EXPORT_DISABLED_COMPUTERS_REPORT_PATH
	} catch {
		Write-Error $_
		Write-Host('[{0}] Waiting 600 seconds and trying again...' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
		Start-Sleep -Seconds 600
		$modified_assets | Export-CSV -NoTypeInformation -Force $EXPORT_DISABLED_COMPUTERS_REPORT_PATH
	}
	if (Test-Path $EXPORT_DISABLED_COMPUTERS_REPORT_PATH -PathType Leaf) {
		Write-Host('[{0}] Disabled computer report has been saved to [{1}].' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $EXPORT_DISABLED_COMPUTERS_REPORT_PATH)
    }
}

# Email out a report on deleted users.
If ($EMAIL_DISABLED_COMPUTERS_REPORT -Or $EmailReport) {
	Write-Host('[{0}] Report is ENABLED.' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
	
	If ([string]::IsNullOrWhiteSpace($EMAIL_SMTP) -Or [string]::IsNullOrWhiteSpace($EMAIL_REPORT_FROM) -Or [string]::IsNullOrEmpty($emailReportTo) -Or [string]::IsNullOrEmpty($EMAIL_REPORT_SUBJECT)) {
		Write-Warning "-EmailReport given, however one of the required email parameters is blank. No email will be sent."
	
	} elseif (-Not $DryRun -And ($modified_total_count -gt 0 -Or $EMAIL_REPORT_ON_NO_ACTIONS)) {
		If ($PendingReport) {
			Write-Host('[{0}] Preparing email for pending disabled systems report' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
		} else {
			Write-Host('[{0}] Preparing email for disabled systems report' -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
		}
		
		$body = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii"><title>HTML TABLE</title>
</head><body>
<p>[$sp_assets_count_found] out of [$sp_assets_count] match criteria for systems previously assigned to deleted users.</p>
"@

		If ($PendingReport) {
			if (-Not [string]::IsNullOrEmpty($nextRun)) {
				$nextRun = "<u>on $nextRun</u>"
			} else {
				$nextRun = "the next time this script is run"
			}
			$body += @"
<p>[$modified_count] of these systems are pending disabling ($modified_excluded_count skipped).</p>
<p>A pending report has been saved at [<a href="file://$EXPORT_DISABLED_COMPUTERS_PENDING_PATH">$EXPORT_DISABLED_COMPUTERS_PENDING_PATH</a>].</p>
<p><b>Computers listed in this report will be disabled $nextRun if their status has not been updated in Snipe-It</b> (not including those marked 'Excluded').</p>
<p>If any items need to be updated, make sure to change the status and reassign the asset in Snipe-It before the script is next run.</p>
"@
		} else {
			$body += @"
<p><b>[$modified_count] of these systems have been disabled or moved</b> ($modified_excluded_count skipped).</p>
<p>A report of all actions taken has been saved to [<a href="file://$EXPORT_DISABLED_COMPUTERS_REPORT_PATH">$EXPORT_DISABLED_COMPUTERS_REPORT_PATH</a>].</p>
"@
		}
$body += @"
<br>
<p>This message was automatically generated from [$_scriptName] running on [${ENV:COMPUTERNAME}].
"@

		$emailParams = @{
			From = $EMAIL_REPORT_FROM
			To = $emailReportTo
			Subject = $EMAIL_REPORT_SUBJECT
			Body = $body
			Priority = "Normal"
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $EMAIL_SMTP
			BodyAsHtml = $true
		}
		if ($PendingReport) {
			if (-Not [string]::IsNullOrEmpty($EMAIL_REPORT_PENDING_SUBJECT)) {
				$emailParams["Subject"] = $EMAIL_REPORT_PENDING_SUBJECT
			}
			$emailParams["Priority"] = "High"
		}
		try {
			Send-MailMessage @emailParams -ErrorAction Stop
		} catch {
			Write-Error $_
			$error_count++
		}

		Write-Host("[{0}] Emailed report to [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($emailReportTo -join ", "))
	}
}

Write-Host("[{0}] Caught {1} errors" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $error_count)

$runtimeDiff = ((Get-Date) - $dateStart)
Write-Host("[{0}] Total Runtime: {1} hours {2} minutes ({3} total minutes)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $runtimeDiff.Hours, $runtimeDiff.Minutes, $runtimeDiff.TotalMinutes)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

If (-Not $DryRun -And -Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1))) {
	Email-ErrorReport -To $EMAIL_ERROR_REPORT_TO -From $EMAIL_ERROR_REPORT_FROM -SmtpServer $EMAIL_SMTP -ErrorCount $error_count -LogFilePath $_logfilepath -Verbose
}
