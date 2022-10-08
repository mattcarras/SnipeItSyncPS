# Syncs asset information with Snipe-It.
#
# Requirements:
# * RSAT: Active Directory PowerShell module if you're getting results from AD.
# * SnipeItPS module: https://github.com/snazy2000/SnipeitPS
# * SnipeIt-Sync-PS.ps1: https://github.com/mattcarras/SnipeItSyncPS
#
# Install-Module SnipeitPS
# Update-Module SnipeitPS
# Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
#
# Author: Matthew Carras
# Source: https://github.com/mattcarras/SnipeItSyncPS

# -- START CONFIGURATION --
# Previously exported credentials from Export-SnipeItCredentials
$CREDXML_PATH = "your_exported_credentials.xml"

# This must evaluate to $true to actually start syncing. Otherwise the script skips syncing entirely.
# It also gives a debug breakpoint, if you have debugging enabled.
$ENABLE_SYNC = $false

# File or path to import assets from previously exported reports.
<#
$IMPORT_CSV_PATH = "\\path\to\SCCM_Export*.csv"
# A unique identifier to group the results by.
$IMPORT_CSV_GROUP_BY = "Unique_Identifier"
#>

# Mapping for Snipe-It field names.
# "SnipeItAssetField"="Asset Field Name"
# A Snipe-It field may map to more than one field in the asset.
$ASSET_FIELD_MAP = @{ 
    "Serial"="Serial_Number"
    "Name"="Computer_Name"
    "Model"="Model"
    "Manufacturer"="Manufacturer"
    "SMBIOS GUID"="SMBIOS_GUID"
    "SCCM LastActiveTime"="LastActiveTime"
    "AD LastLogonTime"="ADLastLogonTime"
    "Category"="Platform"
    "Fieldset"="Platform"
    "System Form Factor"="Type"
    "AD SID"="SID"
    "LastLogonUser"="User_Name"
    "Primary Users"="Primary_User"
}

# Searchbases to optionally sync information from AD.
<#
$AD_IMPORT_SEARCHBASES = @("OU=Domain Computers,DC=Fabrikam,DC=COM")
#>

# Which fields to sync if the given field matches.
# A value of $true syncs ALL fields in $ASSET_FIELD_MAP.
# Matched fields which return more than 1 result are never used.
$ASSET_FIELD_SYNC_ON_MAP = [ordered]@{
    "Serial" = $true
    "SMBIOS GUID" = $true
    "AD SID" = @("Name", "SCCM LastActiveTime", "LastLogonUser", "AD LastLogonTime")
    "Name" = @("SCCM LastActiveTime", "LastLogonUser", "AD LastLogonTime")
}
# Which fields are required to be non-blank before creating an asset.
# Only checked when creating new assets.
# This must have at least one entry.
$ASSET_FIELD_CREATE_REQUIRED = @("Name","Serial","SMBIOS GUID","Model","Manufacturer","Category")
# Status used when creating. This can be the status name or ID. This must be a valid status.
$ASSET_CREATE_STATUS = "Pending"
# Optional status used to change archived assets to when encountered in sync. This can be the status name or ID.
# $ASSET_ARCHIVED_STATUS_CHANGE_TO = "New"
# Default model name used when no model is found. This only accepts model names.
# If given, this model must already be created in Snipe-It.
$ASSET_DEFAULT_MODEL = "_Unknown PC Model_"

# To make doubly sure we aren't duplicating any entities, halt if the list of assets, models, manufacturers, fieldsets, categories, etc. are empty
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $false

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "snipeit-asset-sync"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Filepath for exports of assets from imported sources and snipe-it
<#
$EXPORTS_PATH = "\\path\to\Snipe-It\Exports"
$EXPORTS_PREFIX_FORMATTED = "assets_formatted"
$EXPORTS_PREFIX_SNIPEIT = "assets_snipeit"
$EXPORTS_ROTATE_DAYS = 365
#>

# Email configuration for reports
<#
$EMAIL_SMTP = '<smtp server>'
# If filled out, send error reports
$EMAIL_ERROR_REPORT_FROM = '<from address>'
# May include multiple destination addresses as an array.
$EMAIL_ERROR_REPORT_TO = '<to address>'
#>
# -- END CONFIGURATION --

# -- START --

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
    Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append

# -- START FUNCTIONS --
function Get-ComputerPlatformFromOS {
    <#
        .SYNOPSIS
        Returns platform from Operating System string ("Mac", "Linux", or "PC"), or empty string if not found.
        
        .DESCRIPTION
        Returns platform from Operating System string ("Mac", "Linux", or "PC"), or empty string if not found. This is intended to be used with values from SCCM and AD.
        
        .PARAMETER OS
        The operating system string.

        .OUTPUTS
        Either "Mac", "Linux", or "PC".
        
        .Example
        PS> Get-ComputerPlatformFromOS $_.OperatingSystem
    #>
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

Function Get-IsVirtualMachine {
    <#
        .SYNOPSIS
        Returns whether given Model and/or Manufacturer is a Virtual Machine.
        
        .DESCRIPTION
        Returns whether given Model and/or Manufacturer is a Virtual Machine.
        
        .PARAMETER Model
        The system model.

        .PARAMETER Manufacturer
        The system manufacturer.
        
        .OUTPUTS
        True if matching a Virtual Machine, false otherwise.
        
        .Example
        PS> Get-IsVirtualMachine $_.Model $_.Manufacturer
    #>
    param (
        [parameter(Mandatory=$false, Position=0)]
        [string]$Model,
        
        [parameter(Mandatory=$False, Position=1)]
        [string]$Manufacturer
    )
    
    return ($Model -imatch "Virtual" -Or $Model -eq "HVM domU" -Or $Manufacturer -eq "Xen" -Or $Manufacturer -eq "QEMU")
}

function Get-ComputerFormFactorFromChassis {
    <#
        .SYNOPSIS
        Return system form factor from WMI / SCCM chassis type.
        
        .DESCRIPTION
        Return system form factor from WMI / SCCM chassis type.
        
        .PARAMETER ChassisType
        The chassis type to evaluate.
        
        .OUTPUTS
        The system form factor as a string if matching, empty string otherwise.
        
        .Example
        PS> Get-ComputerFormFactorFromChassis "13"
    #>
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
            default { 
                return ''
            }
        }
    }
    End {
    }
}

function Format-AssetForSyncing {
    <#
        .SYNOPSIS
        Formats one or more asset object(s) for syncing with Snipe-It.
        
        .DESCRIPTION
        Formats one or more asset object(s) for syncing with Snipe-It, using a property map to convert properties into the format required by Sync-SnipeItAsset.
        
        .PARAMETER Asset
        Required. One or more assets to format.
        
        .PARAMETER PropertyMap
        A hashtable of "SnipeItAssetField"="AssetProperty".

        .PARAMETER DateFormat
        The string format for all dates. Defaults to 'yyyy-MM-dd', which is the format currently required by Snipe-It.
        
        .OUTPUTS
        The asset object(s) formatted for use with Sync-SnipeItAsset.
        
        .Example
        PS> $ad_assets | Format-AssetForSyncing
    #>
    param (
        [parameter(Mandatory=$true,
                    Position = 0,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName=$true)]
        [object[]]$Asset,
        
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$PropertyMap = @{
            "Serial"="SerialNumber"
            "Name"="Name"
            "Model"="Model"
            "Manufacturer"="Manufacturer"
            "Category"="Type"
        },
        
        [parameter(Mandatory=$false)]
        [string]$DateFormat='yyyy-MM-dd'
    )
    Begin {
        # Compute the given Property Map into an array for Select-Object.
        $SelectArray = $PropertyMap.GetEnumerator() | where {-Not [string]::IsNullOrWhitespace($_.Value)} | foreach {
            $val = $_.Value
            @{N=$_.Name; Expression=[Scriptblock]::Create("if (`$_.'$val' -is [DateTime]) { ([DateTime]`$_.'$val').ToString('$DateFormat') } else { `$_.'$val' }") }
        }
    }
    Process {
        return $Asset | Select $SelectArray
    }
    End {
    }
}

function Get-WMIObjectAsJob {
    <#
        .SYNOPSIS
        Helper function to get a remote WMI Object as a job in case of possible timeout.
        
        .DESCRIPTION
        Helper function to get a remote WMI Object as a job in case of possible timeout.
        
        .PARAMETER ComputerName
        Required. The local or remote computer name.
        
        .PARAMETER Class
        Required. The class to query.

        .PARAMETER Filter
        The optional string filter to apply.
        
        .PARAMETER Timeout
        Time in seconds to wait for a response. Defaults to 30 seconds.
        
        .OUTPUTS
        The result of the job. Throws an error on timeout.
        
        .Example
        PS> Get-WMIObjectAsJob -ComputerName "Foobar" -Class "Win32_BIOS"
    #>
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

function Import-AssetFromWMI {
    <#
        .SYNOPSIS
        Get information about given computer through local or remote WMI.
        
        .DESCRIPTION
        Get information about given computer through local or remote WMI.
        
        .PARAMETER ComputerName
        Required. The local or remote computer name. Give ${ENV:COMPUTERNAME} to query the local system.
        
        .OUTPUTS
        An object with the queried info from WMI.
        
        .Example
        PS> Import-AssetFromWMI -ComputerName ${ENV:COMPUTERNAME}
    #>
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
    <#
        .SYNOPSIS
        Joins two object arrays, grouping by the given shared property.
        
        .DESCRIPTION
        Joins two object arrays, grouping by the given shared property.
        
        The Left object is always considered the authority if values are non-null on both sides, unless it's a DateTime, in which case whichever DateTime is newer is used.
        
        .PARAMETER Left
        Required. The left side array of objects. This is the considered the authority, except for DateTimes.
        
        .PARAMETER Right
        Required. The right side array of objects.
        
        .OUTPUTS
        An array of objects with merged properties from both arrays.
        
        .Example
        PS> Join-Assets $sccm_assets $ad_assets "SID"
    #>
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

# An example of a function which imports assets from an exported SCCM report.
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
            "Exists in SCCM" = $true
        }
    }
    
    Write-Verbose("[Import-AssetsFromCSV] {0} unique assets imported from CSV" -f $assets.Count)
    return $assets
}

# An example of a function which imports from AD.
function Import-AssetsFromAD {
    param (
        [parameter(Mandatory=$false)]
        [string[]]$SearchBase,
    
        [parameter(Mandatory=$false)]
        [ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
        [string]$Filter="*"
    )

    $props = @("distinguishedname","LastLogonDate","OperatingSystem")
    if ($SearchBase.Count -gt 0) {
        Write-Verbose("[Import-AssetsFromAD] Collecting all computers from AD using Searchbases [{0}] and Filter [{1}], this might take a while..." -f ($SearchBase -join ", "), $Filter)
        $assets = $Searchbase | foreach { Get-ADComputer -SearchBase $_ -Filter $Filter -Properties $props}
    } else {
        Write-Verbose("[Import-AssetsFromAD] Collecting all computers from AD using Filter [{0}], this might take a while..." -f $Filter)
        $assets = Get-ADComputer -Filter $Filter -Properties $props
    }
    $assets = $assets | foreach {
        [PsCustomObject]@{
            "Computer_Name" = $_.name
            "SID" = $_.SID.Value
            "ADLastLogonTime" = $_.LastLogonDate -as [DateTime]
            "Platform" = Get-ComputerPlatformFromOS -OS $_.OperatingSystem
            "Exists in AD" = $true
        }
    }

    Write-Verbose("[Import-AssetsFromCSV] {0} assets imported from AD" -f $assets.Count)
    return $assets
}
# -- END FUNCTIONS --

# Dot-source custom sync API
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

# Import assets from exported CSV reports and optionally AD, and join the results.
If ((Test-Path $IMPORT_CSV_PATH -PathType Leaf) -And -Not [string]::IsNullOrEmpty($IMPORT_CSV_GROUP_BY)) {
    $sccm_assets = Import-AssetsFromCSV -Filepath $IMPORT_CSV_PATH -GroupBy $IMPORT_CSV_GROUP_BY -Verbose
} else {
    $sccm_assets = @()
}
if ($AD_IMPORT_SEARCHBASES.Count -gt 0) {
    $ad_assets = Import-AssetsFromAD -SearchBase $AD_IMPORT_SEARCHBASES -Verbose | where {$_.Platform -eq "PC"}
} else {
    $ad_assets = @()
}
# Join the results together.
$joined_assets = Join-Assets -Left $sccm_assets -Right $ad_assets -On "SID" -Verbose

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
If ($ASSET_FIELD_MAP.ContainsKey("supplier") -Or $ASSET_FIELD_MAP.ContainsKey("supplier_id")) {
    $cacheentities += @("suppliers")
}

$extraParams = @{}
If ($DEBUG_HALT_ON_NULL_CACHE) {
    $extraParams.Add("ErrorOnNullEntities", $cacheentities)
}
Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose @extraParams

# Filter out those that don't exist in SCCM and VMs, and format for syncing
$formatted_assets = $joined_assets | where {$_."Exists in SCCM" -eq $true -And -Not $_.IsVirtualMachine} | Format-AssetForSyncing -PropertyMap $ASSET_FIELD_MAP

# Export all formatted assets
if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container)) {
    # Rotate exports
    if ($EXPORTS_ROTATE_DAYS -is [int] -And $EXPORTS_ROTATE_DAYS -gt 0) {
        Get-ChildItem "${EXPORTS_PATH}\*" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$EXPORTS_ROTATE_DAYS) } | Remove-Item -Force
    }
    if ($EXPORTS_PREFIX_FORMATTED -is [string]) {
        $fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_FORMATTED}.csv"
        if (Test-Path $fp -PathType Leaf) {
            Remove-Item $fp -Force
        }
        Write-Host("[{0}] Exporting formatted assets to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
        $formatted_assets | Export-CSV -NoTypeInformation $fp
    }
}

$error_count = 0
if (-Not $ENABLE_SYNC) {
    Write-Host('Please set $ENABLE_SYNC to $true when ready to start syncing.')
    Write-Debug('Debug breakpoint due to $ENABLE_SYNC not set.')
} else {
    Write-Host("[{0}] Starting sync..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
    $extraParams = @{}
    if (-Not [string]::IsNullOrWhitespace($ASSET_DEFAULT_MODEL)) {
        $extraParams.Add("DefaultModel", $ASSET_DEFAULT_MODEL)
    }
    if (-Not [string]::IsNullOrWhitespace($ASSET_ARCHIVED_STATUS_CHANGE_TO)) {
        $extraParams.Add('UpdateArchivedStatus', $ASSET_ARCHIVED_STATUS_CHANGE_TO)
    }
    foreach ($asset in $formatted_assets) {
        try {
            $sp_asset = Sync-SnipeItAsset -Asset $asset -UniqueIDField "Name" -SyncFields $ASSET_FIELD_MAP.Keys -SyncOnFieldMap $ASSET_FIELD_SYNC_ON_MAP -RequiredCreateFields $ASSET_FIELD_CREATE_REQUIRED -DefaultCreateStatus $ASSET_CREATE_STATUS -Verbose @extraParams
        } catch {
            Write-Error $_
            $error_count += 1
        }
    }
}

Write-Host("[{0}] Caught {1} errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

if ($EXPORTS_PATH -is [string] -And (Test-Path $EXPORTS_PATH -PathType Container) -And $EXPORTS_PREFIX_SNIPEIT -is [string]) {
    $fp = "${EXPORTS_PATH}\${EXPORTS_PREFIX_SNIPEIT}_$(get-date -f yyyy-MM-dd).csv"
    if (Test-Path $fp -PathType Leaf) {
        Remove-Item $fp -Force
    }
    try {
        Write-Host("[{0}] Exporting assets from snipe-it to CSV file [{1}]..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $fp)
        $formatted_assets = $sp_assets | Format-SnipeItAsset -AddDepartment
        # Ensure we have all possible custom fields in output
        # initial_props are always ordered first, the other columns are semi-sorted
        $initial_props = @('asset_tag','name','serial','status_label','assigned_to','Department','manufacturer','model','category')
        $props = $formatted_assets | % { Get-Member -MemberType NoteProperty -InputObject $_ | Select -ExpandProperty Name } | Select -Unique | where {$_ -notin $initial_props}
        if ($props -is [array]) {
            $props = $initial_props + $props
        } elseif ($props -is [string]) {
            $props = $initial_props + @($props)
        } else {
            # Should never get here
            $props = $initial_props
        }

        $formatted_assets | Select $props | Export-CSV -NoTypeInformation $fp
    } catch {
        Write-Error $_
        $error_count++
    }
}

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

# Email out notifications of any errors.
if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP) -And ($error_count -gt 0 -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And ($EMAIL_ERROR_REPORT_TO -is [string] -Or ($EMAIL_ERROR_REPORT_TO -is [array] -And $EMAIL_ERROR_REPORT_TO.Count -gt 0)))) {
    $params = @{
        From = $EMAIL_ERROR_REPORT_FROM
        To = $EMAIL_ERROR_REPORT_TO
        Subject = 'Errors from Snipeit-Asset-Sync'
        Body = "There were [$error_count] caught errors from [Snipeit-Asset-Sync.ps1] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
        Priority = 'High'
        DeliveryNotificationOption = @('OnSuccess', 'OnFailure')
        SmtpServer = $EMAIL_SMTP
    }
    try {
        # Attempt to send with an attachment. If that throws an error for some reason, try sending without it.
        Send-MailMessage -Attachments $_logfilepath @params
    } catch {
        Write-Error $_
        $params['Body'] = "There were [$error_count] caught errors from [Snipeit-Asset-Sync.ps1] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details."
        Send-MailMessage @params
    }
    Write-Host("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($EMAIL_ERROR_REPORT_TO -join ", "))
}