<#
	.SYNOPSIS
	Syncs assets with Snipe-It from SCCM and AD exports.
	
	.DESCRIPTION
	Syncs assets with Snipe-It from SCCM and AD exports. Requires RSAT, SnipeitPS, and Snipeit-Sync-PS powershell modules. Uses settings file Snipeit-Asset-Sync-Settings.ps1.

	.OUTPUTS
	Return Codes
	-1 Error loading settings file
	-2 Error loading Snipeit-Sync-PS API
	-3 Error connecting to Snipe-It site

	.NOTES
	Uses Snipeit-Asset-Sync-Settings.ps1 settings file.
	
	Requirements:
	* RSAT: Active Directory PowerShell module.
	* SnipeItPS module (1.10.225 or newer): https://github.com/snazy2000/SnipeitPS
	* SnipeIt-Sync-PS.ps1: https://github.com/mattcarras/SnipeItSyncPS

	Install-Module SnipeitPS
	Update-Module SnipeitPS

	Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
	
	Author: Matthew Carras
	Source: https://github.com/mattcarras/SnipeItSyncPS
#>

# -- LOAD CONFIGURATION --
try {
	. .\Snipeit-Asset-Sync-Settings.ps1
} catch {
	Write-Error $_
    return -1
}

# -- START --
$dateStart = Get-Date
$_scriptName = split-path $PSCommandPath -Leaf
$error_count = 0

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

# Return Platform from OperatingSystem name
function Get-ComputerPlatformFromOS {
	param (
		[parameter(Mandatory=$false,
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[alias("OperatingSystem")]
		[string]$OS
	)	
	Begin {
	}
	Process {
		switch($OS) {
			{$_ -imatch "Mac"} {
				return "Mac"
			}
			{$_ -imatch "Linux"} {
				return "Linux"
			}
			{-Not [string]::IsNullOrWhitespace($_)} {
				return "PC"
			}
			default {
				return ''
			}
		}
	}
	End {
	}
}

# Return whether model or manufacturer matches a virtual machine
Function Get-IsVirtualMachine {	
	param (
		[parameter(Mandatory=$false, Position=0)]
		[string]$Model,
		
		[parameter(Mandatory=$False, Position=1)]
		[string]$Manufacturer
	)
	
	return ($Model -imatch "Virtual" -Or $Model -eq "HVM domU" -Or $Manufacturer -eq "Xen" -Or $Manufacturer -eq "QEMU")
}

# Return Computer Form Factor from WMI ChassisType
function Get-ComputerFormFactorFromChassis {
	param (
		[parameter(Mandatory=$false,
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$ChassisType
	)	
	Begin {
	}
	Process {
		switch($ChassisType) { 
			{$_ -in "3", "4", "5", "6", "7", "15", "16"} { 
				return "Desktop" 
			} 
			{$_ -in "13"} { 
				return "All-In-One"
			}
			{$_ -in "8", "9", "10", "11", "12", "14", "18", "21","31","32"} { 
				return "Laptop" 
			} 
			{$_ -in "30"} { 
				return "Tablet"
			} 
			{$_ -in "17","23"} { 
				return "Server"
			}
			Default { 
				return ''
			}
		}
	}
	End {
	}
}

# Format asset to properties expected by Snipe-It.
function Format-AssetForSyncing {
	param (
		[parameter(Mandatory=$true,
					Position = 0,
					ValueFromPipeline = $true,
					ValueFromPipelineByPropertyName=$true)]
		[object[]]$Asset,
		
		[parameter(Mandatory=$false)]
		[hashtable]$PropertyMap = @{
			"Serial"="SerialNumber"
			"Name"="Name"
			"Model"="Model"
			"Manufacturer"="Manufacturer"
			"Category"="Type"
		}
	)
	Begin {
		# Compute the given Property Map into an array for Select-Object.
		$SelectArray = $PropertyMap.GetEnumerator() | where {-Not [string]::IsNullOrWhitespace($_.Value)} | foreach {
            $val = $_.Value
			# The format of 'yyyy-MM-dd' is required for compatibility with Snipe-It.
			@{N=$_.Name; Expression=[Scriptblock]::Create("if (`$_.'$val' -is [DateTime]) { ([DateTime]`$_.'$val').ToString('yyyy-MM-dd') } else { `$_.'$val' }") }
		}
	}
	Process {
		return $Asset | Select $SelectArray
	}
	End {
	}
}

# Import a list of computers from an exported SCCM Report.
# TODO: Make this function generic.
function Import-AssetsFromCSV {
	param (
		[parameter(Mandatory=$true,
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[ValidateScript({Test-Path $_ -PathType Leaf})]
		[string]$Filepath,
		
		[parameter(Mandatory=$false)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$GroupBy
	)	
	
	Write-Verbose("[Import-AssetsFromCSV] Importing latest asset exports from [$Filepath]...")
	
	# SCCM Headers: Computer_Name,Unique_Identifier,SID,Domain_or_Workgroup,SMBIOS_GUID,MAC_Address,User_Name,Console_User,Primary_users,Last_Logon_Time,Operating_Sytem,OS_Build,Is_Virtual_Machine0,LastActiveTime,ADLastLogonTime,Manufacturer,Model,Serial_Number,Chassis
	# Group by a unique field (in this case, Unique_Identifier, the SCCM resource ID)
	$assets = Import-CSV -LiteralPath (Get-ChildItem $FilePath | Sort-Object {$_.LastWriteTime} | Select -Last 1 | Select -ExpandProperty FullName) | Group-Object $GroupBy | Foreach-Object {
		$lastActive = ($_.Group.LastActiveTime | Select -First 1) -as [DateTime]
		$lastLogon = ($_.Group.ADLastLogonTime | Select -First 1) -as [DateTime]
		$primaryUser = ($_.Group.Primary_User | Select -Unique) -join "; "
		if ([string]::IsNullOrWhitespace($primaryUser)) {
			$primaryUser = ($_.Group.Console_User | Select -Unique) -join "; "
		}
		$model = $_.Group.Model | Select -First 1
		$manufacturer = $_.Group.Manufacturer | Select -First 1
		
		[PsCustomObject]@{
			"Model" = $model
			"Manufacturer" = $manufacturer
			"LastActiveTime" = $lastActive
			"ADLastLogonTime" = $lastLogon
			"Computer_Name" = $_.Group.Computer_Name | Select -First 1
			"Serial_Number" = $_.Group.Serial_Number | Select -First 1
			"SMBIOS_GUID" = $_.Group.SMBIOS_GUID | Select -First 1
			"SID" = $_.Group.SID | Select -First 1
			"User_Name" = $_.Group.User_Name | Select -First 1
			"Primary_User" = $primaryUser
			"Type" = Get-ComputerFormFactorFromChassis -ChassisType ($_.Group.Chassis | Select -First 1)
			"Platform" = Get-ComputerPlatformFromOS -OS ($_.Group.Operating_Sytem | Select -First 1)
			"IsVirtualMachine" = (($_.Group.Is_Virtual_Machine0 | where {$_ -eq $true}).Count -gt 0) -Or (Get-IsVirtualMachine -Model $model -Manufacturer $manufacturer)
			"OS_Build"=$_.Group.OS_Build | Select -First 1
			"Exists In SCCM" = $true
		}
	}
	
    Write-Verbose("[Import-AssetsFromCSV] {0} unique assets imported from CSV" -f $assets.Count)
	return $assets
}

# Return a set location from the IP.
function Get-LocationFromIP {
	param (
		[parameter(Mandatory=$true,
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$IP
	)
	Begin {
	}
	Process {
		$location = $null
		<#
		switch ($IP) {
			{$_ -like "10.160.1.*"} {
				$location = 'Building A'
			}
		}
		#>
		return $location
	}
	End {
	}
}

# Import assets from AD.
# TODO: Make this function generic.
function Import-AssetsFromAD {
	param (
		[parameter(Mandatory=$false)]
		[string[]]$SearchBase,
	
		[parameter(Mandatory=$false)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$Filter="*"
	)
	
    Write-Verbose("[Import-AssetsFromAD] Collecting all assets from AD using searchbases [{0}], this might take a while..." -f ($SearchBase -join ", "))

	$props = @("distinguishedname","LastLogonDate","OperatingSystem","OperatingSystemVersion","IPV4Address")
	if ($SearchBase.Count -gt 0) {
		$assets = $Searchbase | foreach { Get-ADComputer -SearchBase $_ -Filter $Filter -Properties $props}
	} else {
		$assets = Get-ADComputer -Filter $Filter -Properties $props
	}
	$assets = $assets | foreach {
		$location = $null
		if (-Not [string]::IsNullOrEmpty($_.IPV4Address)) {
			$location = Get-LocationFromIP $_.IPV4Address
		}
		# Format OperatingSystemVersion to match OS_Build from SCCM
		[PsCustomObject]@{
			"Computer_Name" = $_.name
			"SID" = $_.SID.Value
			"ADLastLogonTime" = $_.LastLogonDate -as [DateTime]
			"Platform" = Get-ComputerPlatformFromOS -OS $_.OperatingSystem
			"Location" = $location
			"OS_Build" = ($_.OperatingSystemVersion -replace " \(",".") -replace "\)",""
			"Exists In AD" = $true
		}
	}

    Write-Verbose("[Import-AssetsFromCSV] {0} assets imported from AD" -f $assets.Count)
	return $assets
}

# Helper function to get a remote WMI Object as a job in case of possible timeout.
function Get-WMIObjectAsJob {
	param (
		[parameter(Mandatory=$true, 
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$ComputerName,
		
		[parameter(Mandatory=$true)]
		[string]$Class,
		
		[parameter(Mandatory=$false)]
		[string]$Filter,
		
		[parameter(Mandatory=$false)]
		[ValidateRange(0,[int]::MaxValue)]
		[int]$Timeout=30
	)
	Begin {
		$extraParams = @{}
		if ($Filter -is [string]) {
			$extraParams.Add("Filter", $Filter)
		}
	}
	Process {
		$job = Get-WmiObject -Class $Class -ComputerName $ComputerName -AsJob @extraParams | Wait-Job -Timeout $Timeout
		if ($job.State -eq 'Completed') {
			return Receive-Job -Job $job
		}
		Throw "[Get-WMIObjectAsJob] [$ComputerName] timed out"
	}
	End {
	}
}

# Get Asset info directly from WMI. Give ${ENV:COMPUTERNAME} to pull from the current computer.
function Import-AssetFromWMI {
	param (
		[parameter(Mandatory=$true, 
				   Position=0,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$ComputerName
	)
	Begin {
	}
	Process {
		$owmi = Get-WmiObjectAsJob $ComputerName -Class 'Win32_BIOS'
		$serial = $owmi.SerialNumber
		$manufacturer = $owmi.Manufacturer
		$owmi = Get-WmiObjectAsJob $ComputerName -Class 'Win32_ComputerSystem'
		$name = $owmi.Name
		$model = $owmi.Model
		if (-Not [string]::IsNullOrWhitespace($owmi.Manufacturer)) {
			$manufacturer = $owmi.Manufacturer
		}
		$owmi = Get-WmiObjectAsJob $ComputerName -Class 'Win32_SystemEnclosure'
		if ($owmi.ChassisTypes.Count -gt 0) {
			$type = Get-ComputerFormFactorFromChassis ($owmi.ChassisTypes | Select -First 1)
		}
		$owmi = Get-WmiObjectAsJob $ComputerName -Class 'Win32_OperatingSystem'
		if ($owmi.Caption -ne $null) {
			$platform = Get-ComputerPlatformFromOS -OS $owmi.Caption
		}
		
		return [PsCustomObject]@{
			"Computer_Name" = $name
			"Model" = $model
			"Manufacturer" = $manufacturer
			"Serial_Number" = $serial
			"Type" = $type
			"Platform" = $platform
			"IsVirtualMachine" = Get-IsVirtualMachine -Model $model -Manufacturer $manufacturer
		}
	}
	End {
	}
}

# Joins two arrays on the given key.
function Join-Assets {
	param (
		[parameter(Mandatory=$true, Position=0)]
        [AllowEmptyCollection()]
		[array]$Left,
		
		[parameter(Mandatory=$true, Position=1)]
        [AllowEmptyCollection()]
		[array]$Right,
		
		[parameter(Mandatory=$true, Position=2)]
		[string]$On
	)
	
    Write-Verbose("[Join-Assets] Joining {0} assets (left) with {1} assets (right) on [{2}]..." -f $Left.Count,$Right.Count,$On)

    if ($Right.Count -eq 0) {
        return $Left
    } elseif ($Left.Count -eq 0) {
        return $Right
    }
		
	return ($Left + $Right) | Group-Object -Property $On | foreach { 
		if ($_.Count -eq 1) {
			[PSCustomObject]($_.Group | Select -First 1)
		} else {
			$o = [PSCustomObject]@{}
			foreach ($p in ($_.Group | foreach { $_ | Get-Member -MemberType NoteProperty } | Select -ExpandProperty Name -Unique)) {
				$val = $_.Group[0].$p
				if ($val -is [bool] -Or $_.Group[1].$p -is [bool]) {
					$val = ($val -Or $_.Group[1].$p)
				} elseif ($val -eq $null -Or ($val -is [string] -And [string]::IsNullOrEmpty($val) -And -Not [string]::IsNullOrEmpty($_.Group[1].$p)) -Or ($val -is [DateTime] -And $_.Group[1].$p -is [DateTime] -And $_.Group[1].$p -gt $val)) {
					$val = $_.Group[1].$p
				}
				Add-Member -InputObject $o -MemberType NoteProperty -Name $p -Value $val -Force
			}
			$o
		}
	}
}

# -- END FUNCTIONS --

# Load custom API
try {
    . .\SnipeIt-Sync-PS.ps1
} catch {
    # Fatal error, exit
    Write-Error $_
    return -1
}

# Initialize new Snipe-It Session
try {
    Connect-SnipeIt -CredXML $CREDXML_PATH -Verbose
} catch {
    # Fatal error, exit
    Write-Error $_
    return -2
}

# Import assets from exported CSV reports and optionally AD, and join the results.
If ($IMPORT_CSV_PATH -And $IMPORT_CSV_GROUP_BY) {
	$sccm_assets = Import-AssetsFromCSV -Filepath $IMPORT_CSV_PATH -GroupBy $IMPORT_CSV_GROUP_BY -Verbose
	# Filter out only assets set from $SCCM_IMPORT_PLATFORMS (default "PC" only)
	if ($SCCM_IMPORT_PLATFORMS -isnot [array]) {
		# Default to PC only
		$sccm_assets = $sccm_assets | where {$_.Platform -eq "PC" -Or ($SCCM_IMPORT_EMPTY_PLATFORM -And [string]::IsNullOrWhitespace($_.Platform))}
	} elseif($SCCM_IMPORT_PLATFORMS.Count -eq 0) {
		Write-Warning "SCCM_IMPORT_PLATFORMS is set but empty, ignoring platform restrictions"
	} else {
		$sccm_assets = $sccm_assets | where {($_.Platform -in $SCCM_IMPORT_PLATFORMS) -Or ($SCCM_IMPORT_EMPTY_PLATFORM -And [string]::IsNullOrWhitespace($_.Platform))}
	}
} else {
	$sccm_assets = @()
}

#Write-Debug($sccm_assets | Select -First 1)

if ($AD_IMPORT_SEARCHBASES.Count -gt 0) {
	$ad_assets = Import-AssetsFromAD -SearchBase $AD_IMPORT_SEARCHBASES -Verbose
    # Export all joined assets
    if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container) -And $EXPORTS_PREFIX_AD -is [string]) {
	    # Rotate previous exports
	    if ($EXPORTS_ROTATE_DAYS -is [int] -And $EXPORTS_ROTATE_DAYS -gt 0) {
		    Get-ChildItem "${EXPORTS_PATH}\${EXPORTS_PREFIX_AD}_*" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$EXPORTS_ROTATE_DAYS) } | Remove-Item -Force
	    }
	    $fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_AD}_$(get-date -f yyyy-MM-dd).csv"
	    Write-Host("[{0}] Exporting assets from AD to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
	    $ad_assets | Export-CSV $fp -NoTypeInformation -Force
    }
    # Filter out only assets set from $AD_IMPORT_PLATFORMS (default "PC" only)
	if ($AD_IMPORT_PLATFORMS -isnot [array]) {
		# Default to PC only
		$ad_assets = $ad_assets | where {$_.Platform -eq "PC" -Or ($AD_IMPORT_EMPTY_PLATFORM -And [string]::IsNullOrWhitespace($_.Platform))}
	} elseif($AD_IMPORT_PLATFORMS.Count -eq 0) {
		Write-Warning "AD_IMPORT_PLATFORMS is set but empty, ignoring platform restrictions"
	} else {
		$ad_assets = $ad_assets | where {$_.Platform -in $AD_IMPORT_PLATFORMS -Or ($AD_IMPORT_EMPTY_PLATFORM -And [string]::IsNullOrWhitespace($_.Platform))}
	}
} else {
	$ad_assets = @()
}
#Write-Debug($ad_assets | Select -First 1)

# Join the results together.
$joined_assets = Join-Assets -Left $sccm_assets -Right $ad_assets -On "SID" -Verbose

#Write-Debug($joined_assets | Select -First 1)

# Export all joined assets
if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container) -And $EXPORTS_PREFIX_SCCM_AD -is [string]) {
	# Rotate previous exports
	if ($EXPORTS_ROTATE_DAYS -is [int] -And $EXPORTS_ROTATE_DAYS -gt 0) {
		Get-ChildItem "${EXPORTS_PATH}\${EXPORTS_PREFIX_SCCM_AD}_*" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$EXPORTS_ROTATE_DAYS) } | Remove-Item -Force
	}
	$fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_SCCM_AD}_$(get-date -f yyyy-MM-dd).csv"
	Write-Host("[{0}] Exporting assets from SCCM and AD to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
	$joined_assets | Export-CSV $fp -NoTypeInformation -Force
}

# Example of how to get results directly from SCCI using a WMI query.
# Fields with WMI Timestamps will need to be converted into DateTime like so: 
#   $lastactive = ([WMI] '').ConvertToDateTime($_.SMS_CombinedDeviceResources.LastActiveTime)
#
# $results = Get-WmiObject -Query $WQL -ComputerName $ProviderMachineName -Namespace "root\sms\site_$SiteCode"

# Initialize the snipe-it caches.
$cacheentities = @("statuslabels","manufacturers","categories","fieldsets","models","assets","fields")
If ($ASSET_FIELD_MAP.ContainsKey("company") -Or $ASSET_FIELD_MAP.ContainsKey("company_id")) {
	$cacheentities += @("companies")
}
If ($ASSET_FIELD_MAP.ContainsKey("location") -Or $ASSET_FIELD_MAP.ContainsKey("location_id") -Or $ASSET_FIELD_MAP.ContainsKey("rtd_location_id")) {
	$cacheentities += @("locations")
}
If ($DEBUG_HALT_ON_NULL_CACHE) {
	Initialize-SnipeItCache -EntityTypes $cacheentities -ErrorOnNullEntities $cacheentities -Verbose
} else {
	Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose
}

# Filter out those that don't exist in SCCM and VMs, and format for syncing
# Nulls out Location if anything other than a Desktop or All-In-One
# Note date format is hardcoded as yyyy-MM-dd to ensure compatibility with snipe-it, which should reformat the date as needed.
# Example below filters out VMs and only includes Location for Desktops and All-In-One (in case it's set by IP).
# $formatted_assets = $joined_assets | where {-Not $_.IsVirtualMachine} | Select *,@{N="Location"; Expression={ If ($_.Type -eq "Desktop" -Or $_.Type -eq "All-In-One") { $_.Location } Else { $null }}} -ExcludeProperty Location | Format-AssetForSyncing -PropertyMap $ASSET_FIELD_MAP
# Filtering below only syncs items from SCCM
$formatted_assets = $joined_assets | where {$_."Exists In SCCM" -eq $true} | Format-AssetForSyncing -PropertyMap $ASSET_FIELD_MAP

# Export all formatted assets
if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container) -And $EXPORTS_PREFIX_FORMATTED -is [string]) {
	$fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_FORMATTED}.csv"
	Write-Host("[{0}] Exporting formatted copy of assets to be processed to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
	$formatted_assets | Export-CSV $fp -NoTypeInformation -Force
}

# Extra parameters are found in Snipeit-Asset-Sync-Settings.ps1
$syncExtraParams = @{}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_ARCHIVED_UPDATE)) {
    $syncExtraParams.Add('UpdateArchivedStatus', $ASSET_STATUS_ARCHIVED_UPDATE)
}
if (-Not [string]::IsNullOrWhitespace($ASSET_STATUS_ASSIGNED)) {
	$syncExtraParams.Add('DefaultAssignedStatus', $ASSET_STATUS_ASSIGNED)
}
if (-Not $ASSET_SYNC_ARCHIVED_INCLUDE) {
	$syncExtraParams.Add('SkipArchived', $true)
}
	
$error_count = 0
if ($ENABLE_SYNC -ne $true) {
    Write-Host('[{0}] Skipping sync, please set $ENABLE_SYNC=$true to start syncing' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
} else {
    Write-Host("[{0}] Starting sync..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
    foreach ($asset in $formatted_assets) {
	    try {
			# UniqueIDField - Used mainly for reporting
			# SyncFields - Which fields to sync
			# SyncOnFieldMap - Map of asset fields to snipe-it fields
			# RequiredCreateFields - 
		    $sp_asset = Sync-SnipeItAsset -Asset $asset -UniqueIDField "Name" -SyncFields $ASSET_FIELD_MAP.Keys -SyncOnFieldMap $ASSET_FIELD_SYNC_ON_MAP -RequiredCreateFields $ASSET_FIELD_CREATE_REQUIRED -DefaultCreateStatus $ASSET_STATUS_CREATE -DefaultModel $ASSET_DEFAULT_MODEL -OnlyUpdateBlankFields "Location" -Verbose -VerboseLevel 1 @syncExtraParams
	    } catch {
		    Write-Error $_
		    $error_count += 1
	    }
    }
}

if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container)) {
    Write-Host("[{0}] Preparing to export assets from snipe-it..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
    try {  
		$sp_assets = Get-SnipeItEntityAll "assets" -ReturnValues | Format-SnipeItAsset -AddDepartment -AddDepartmentId -Verbose
        # Ensure we have all possible custom fields in output
        # initial_props are always ordered first, the other columns are semi-sorted
        $initial_props = @('asset_tag','name','serial','status_label','assigned_to','Department','manufacturer','model','category')
        $props = $sp_assets | % { Get-Member -MemberType NoteProperty -InputObject $_ | Select -ExpandProperty Name } | Select -Unique | where {$_ -notin $initial_props}
		if ($props -is [string]) {
			$props = @($props)
		}
        if ($props -is [array]) {
			# Rearrange custom field columns, if they exist
			if($EXPORTS_CUSTOMFIELD_ORDER -is [array]) {
                # Rearrange custom field columns, if they exist
                foreach($customfield in $EXPORTS_CUSTOMFIELD_ORDER) {
                    if ($customfield -in $props) {
                        $initial_props += @($customfield)
                    }
                }
            }
            $props = $initial_props + ($props | where {$_ -notin $initial_props})
        } else {
            # Should never get here
            $props = $initial_props
        }
        $sp_assets = $sp_assets | Select $props | Sort -Property 'Department'

		# Export all assets.
        if ($EXPORTS_PREFIX_SNIPEIT -is [string]) {
            $fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_SNIPEIT}_$(get-date -f yyyy-MM-dd).csv"
            Write-Host("[{0}] Exporting assets from snipe-it to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
            # Rotate previous exports
	        if ($EXPORTS_ROTATE_DAYS -is [int] -And $EXPORTS_ROTATE_DAYS -gt 0) {
		        Get-ChildItem "${EXPORTS_PATH}\${EXPORTS_PREFIX_SNIPEIT}_*.csv" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$EXPORTS_ROTATE_DAYS) } | Remove-Item -Force
	        }
            $sp_assets | Export-CSV $fp -NoTypeInformation -Force
			if ($EXPORTS_PREFIX_SNIPEIT_LATEST -is [string]) {
				Write-Host("[{0}] Making a copy of latest export from snipe-it to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), "${EXPORTS_PATH}\${EXPORTS_PREFIX_SNIPEIT_LATEST}.csv")
				Copy-Item $fp "${EXPORTS_PATH}\${EXPORTS_PREFIX_SNIPEIT_LATEST}.csv" -Force
			}
        }
    } catch {
        Write-Error $_
        $error_count++
    }
}

Write-Host("[{0}] Caught {1} errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

$runtimeDiff = ((Get-Date) - $dateStart)
Write-Host("[{0}] Total Runtime: {1} hours {2} minutes ({3} total minutes)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $runtimeDiff.Hours, $runtimeDiff.Minutes, $runtimeDiff.TotalMinutes)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {  
    # Email out notifications of any errors.
    if ($error_count -gt 0 -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1)))	{
		$emailParams = @{
			From = $EMAIL_ERROR_REPORT_FROM
			To =  $EMAIL_ERROR_REPORT_TO
			Subject = "Errors from $_scriptName"
			Body = "There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
			#Priority = "High"
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $EMAIL_SMTP
		}
        try {
			Send-MailMessage -Attachments $_logfilepath @emailParams -ErrorAction Stop
		} catch {
			Write-Error $_
			$mailParams.Body = "There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details."
			Send-MailMessage @emailParams
		}
        Write-Host("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($EMAIL_ERROR_REPORT_TO -join ", "))
    }
}
