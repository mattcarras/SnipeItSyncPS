# Syncs users from AD with Snipe-It.
#
# Requirements:
# * RSAT: Active Directory PowerShell module.
# * SnipeItPS module (1.10.225 or newer): https://github.com/snazy2000/SnipeitPS
# * SnipeIt-Sync-PS.ps1: https://github.com/mattcarras/SnipeItSyncPS
# 
# Install-Module SnipeitPS
# Update-Module SnipeitPS
# Export credentials: Export-SnipeItCredentials -File "snipeit_creds.xml" -URL "<URL>" -APIKey "<APIKEY>"
#
# Author: Matthew Carras
# Source: https://github.com/mattcarras/SnipeItSyncPS

# Parameter definitions.
# Given parameters override configuration below.
param([switch] $DisableSync, [switch] $ADSyncDeletedUsersPurge, [switch] $EmailDeletedUsersReport, [string] $LogFilePrefix)

# -- START CONFIGURATION --
# Previously exported credentials from Export-SnipeItCredentials
$CREDXML_PATH = "your_exported_credentials.xml"

# This must evaluate to $true to actually start syncing. Otherwise the script skips syncing entirely.
# It also gives a debug breakpoint, if you have debugging enabled.
# Note the -DisableSync switch overrides this setting.
$ENABLE_SYNC = $false

# Target group(s) of users to sync with Snipe-It.
# This should be one or more hashtables in the form of:
#   groupname = group name or array of group names
#   nested = If $true, use -Recursive lookup. May fail if >5000 members returned.
#   activated = If $true, allow login for users in this group.
#   groups = int or int array of groups to assign in Snipe-It (requires SnipeItPS 1.10.225 or newer)
#   ldap_import = Set the ldap_import flag on the Snipe-It user.
$AD_GROUP_TARGETS = @(
    @{"groupname" = "Domain Users"; "nested"=$false; "ldap_import"=$true}
)

# AD properties to sync
# "SnipeitField"="AD Property Name"
# Only these fields will sync.
$AD_GROUP_PROPERTY_MAP = @{
    "username"="UserPrincipalName"
    "employee_num"="SID"
    "first_name"="givenname"
    "last_name"="surname"
    "department"="department"
    "company"="company"
    "jobtitle"="title"
    "email"="mail"
    #"manager"="manager"
    #"location"="physicaldeliveryofficename"
    #"location_address"="foo"   # AD Attribute to sync in the Address field if new location.
}
# Filter the results based on the given map of Properties.
# These properties do not need to be defined in the property map.
# If a hashtable, requires the "Value" and "operator" keys, where "operator" can be any operator supported by PowerShell.
# Otherwise assume the "-ne" operator by default and the value is a string.
$AD_GROUP_PROPERTY_FILTER_MAP = @{
}

# Sync SID to employee_num.
$AD_SYNC_ON_EMPLOYEE_NUM = $true

# Only sync the email address if the user is login-enabled. Ignored if not syncing the email field.
$AD_SYNC_EMAIL_FOR_LOGIN_ONLY = $true

# Purge users that no longer exist in the target AD groups.
# You must have either $AD_SYNC_ON_EMPLOYEE_NUM set to $true or set all your groups with ldap_import=$true.
# Note the -ADSyncDeletedUsersPurge switch overrides this setting.
$AD_SYNC_DELETED_USERS_PURGE = $false

# Only report on deleted users, do not flag or purge.
# Note the -ADSyncDeletedUsersPurge switch overrides this setting.
$AD_SYNC_DELETED_USERS_REPORT_ONLY = $false

# Skip processing deleted users entirely.
# Note the -ADSyncDeletedUsersPurge switch overrides this setting.
$AD_SYNC_DELETED_USERS_SKIP = $false

# Path to save latest deleted users report
# $AD_SYNC_DELETED_USERS_EXPORT_PATH = "path\to\deleted_users_report.csv"

# Reassign any equipment assigned to a deleted user to a special department user if true.
$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT = $false
# Only reassign equipment if the user was deleted from AD entirely.
#$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_ONLY_DELETED = $true
# Change to the given status ID when reassigning assets if set. This status must already exist.
#$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_STATUS_ID = 1

# Create special users for each department to allow assigning assets to departments.
#$AD_SYNC_DEPARTMENT_USERS = $true
# Only create special department users when the following company is set.
#$AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY = "My Company"

# To make doubly sure we aren't duplicating any entities, halt if the list of users, depts, and/or locations are empty.
# This is useful if you know all the entities (users, departments, companies, and locations) should return at least 1 result.
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $false

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "snipeit-ad-sync"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Email configuration for reports
<#
$EMAIL_SMTP = '<smtp server>'
# If filled out, send error reports
$EMAIL_ERROR_REPORT_FROM = '<from address>'
# May include multiple destination addresses as an array.
$EMAIL_ERROR_REPORT_TO = '<to address>'
#>

<#
# You may also give the -EmailDeletedUsersReport script parameter.
# Using this in combination with the -DisableSync and -ADSyncDeletedUsersPurge script parameters allows
# for purging users and emailing out the results.
$EMAIL_DELETED_USERS_REPORT = $false
$EMAIL_DELETED_USERS_REPORT_FROM = '<from address>'
# May include multiple destination addresses as an array.
$EMAIL_DELETED_USERS_REPORT_TO = '<to address>'
# Overrides $EMAIL_DELETED_USERS_REPORT_TO
#$EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS = 'snipeit-reports-group'
$EMAIL_DELETED_USERS_REPORT_SUBJECT = 'Weekly Inactive Snipe-It Users Report'
# Field to check for reporting EOL assets owned by reassigned users (optional).
#$EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD = "End of Life"
#>
# -- END CONFIGURATION --

# -- START --
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
$_logfilepath = "${LOGFILE_PATH}\${_logfileprefix}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append

if (($EMAIL_DELETED_USERS_REPORT -Or $EmailDeletedUsersReport) -And [string]::IsNullOrWhitespace($EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS)) {
    $emailDeletedUsersReportTo = Get-ADGroupMember $EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS -Recursive | foreach { Get-ADUser $_ -Properties mail | Select -ExpandProperty mail }
} else {
    $emailDeletedUsersReportTo = $EMAIL_DELETED_USERS_REPORT_TO
}
    
# -- START FUNCTIONS --
function Get-ADUsersByGroup {
    <#
        .SYNOPSIS
        Collect all AD users from given target group(s).
        
        .DESCRIPTION
        Collect all AD users from given target group(s). If you want to check all users give a global group like Domain Users.
        
        .PARAMETER TargetGroup
        Required. The AD Group(s) to check.
        
        .PARAMETER ADProperties
        The AD properties to return with each user.
        
        .PARAMETER ADPropertyFilterMap
        A hashtable of filters to exclude from the target groups, where each key is a the name of the Property. If the value is a string, assume "Property" -ne "Value". If the value is a hashtable, assume it has the "Value" and "operator" keys, where the operator can be any operator supported by Powershell's Where-Object.
        
        .PARAMETER Nested
        If true calls Get-ADGroupMember with the -Recursive switch instead of Get-ADGroup. Note this may fail to return more than 5000 members depending on your environment.

        .PARAMETER IncludeDisabled
        If true include disabled users.

        .OUTPUTS
        The returned users from AD.
        
        .Example
        PS> Get-ADUsersByGroup "Domain Users" -ADProperties @("department","company","title","manager")
    #>
    param (     
        [parameter(Mandatory=$true,
                    Position = 0,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName=$true)]
        [string[]]$TargetGroup,
        
        [parameter(Mandatory=$false)]
        [AllowEmptyCollection()]
        [string[]]$ADProperties = @("givenname","surname","department","company","title","manager","physicaldeliveryofficename","mail"),
        
        [parameter(Mandatory=$false)]
        [hashtable]$ADPropertyFilterMap = @{},
        
        [parameter(Mandatory=$false)]
        [switch]$Nested,

        [parameter(Mandatory=$false)]
        [switch]$IncludeDisabled
    )
    
    $ad_users = $null
    foreach ($group in $TargetGroup) {
        # Get all users from AD
        Write-Verbose ("[Get-ADUsersByGroup] Collecting all users from AD group [$group] (Nested=$Nested, With FilterMap={0})..." -f ($ADPropertyFilterMap.Count -gt 0))
        
        if ($Nested) {
            # May not work with >5000 results
            $ad_users += Get-ADGroupMember $group -Recursive -ErrorAction Stop | ?{$_.objectClass -eq 'user'}
        } else {
            $ad_users += Get-ADGroup $group -Properties Member -ErrorAction Stop | ?{$_.objectClass -eq 'user'} | Select -ExpandProperty Member
        }
    }
    if ($ad_users -ne $null) {
        # Get extra attributes for each user
        $props = $ADProperties
        if ($props -ne $null -And -Not $props -is [array]) {
            $props = @($props)
        }
        $props += ($ADPropertyFilterMap.Keys + @("distinguishedname")) | Sort -Unique
        Write-Debug "[Get-ADUsersByGroup] Properties: $props"
        Write-Verbose ("[Get-ADUsersByGroup] Getting properties for {0} users..." -f $ad_users.Count)
        $ad_users = $ad_users | foreach { Get-ADUser $_ -Properties $props }
    
        # Create dynamic filter from given parameters
        $filterscript = ($ADPropertyFilterMap.GetEnumerator() | Foreach-Object { if ($_.Value -is [hashtable]) { $op=$_.Value.operator; $val=$_.Value.Value } else { $op="ne"; $val=$_.Value }; "`$_.{0} -{1} `"{2}`"" -f $_.Key, $op, $val}) -join " -AND "
        if (-Not $IncludeDisabled) {
            $filter = "`$_.Enabled -eq `$true"
            if ([string]::IsNullOrWhitespace($filterscript)) {
                $filterscript = $filter
            } else {
                $filterscript += " -AND $filter"
            }
        }
        Write-Debug "[Get-ADUsersByGroup] AD Group Filter: $filterscript"
        if (-Not [string]::IsNullOrWhitespace($filterscript)) {
            $ad_users = $ad_users | Where-Object -FilterScript ([scriptblock]::create($filterscript))
        }
    }
    Write-Verbose ("[Get-ADUsersByGroup] Total filtered AD users collected: {0}" -f $ad_users.Count)
    
    return $ad_users
}

# Format user to properties expected by Snipe-It.
function Format-UserForSyncing {
    <#
        .SYNOPSIS
        Formats one or more user object(s) for syncing with Snipe-It.
        
        .DESCRIPTION
        Formats one or more user object(s) for syncing with Snipe-It, using a property map to convert properties into the format required by Sync-SnipeItUser.
        
        .PARAMETER User
        Required. One or more user objects to format.
        
        .PARAMETER PropertyMap
        A hashtable of "SnipeItUserField"="UserProperty". Just like Sync-SnipeItUser, the "username", "first_name", and "last_name" keys are required.

        .OUTPUTS
        The user objects formatted for use with Sync-SnipeItUser.
        
        .Example
        PS> $ad_users | Format-UserForSyncing
    #>
    param (
        [parameter(Mandatory=$true,
                    Position = 0,
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName=$true)]
        [object[]]$User,
        
        # Note: "activated"="_activated", "groups"="_groups", and "ldap_import"="_ldap_import" are added by default.
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({-Not [string]::IsNullOrWhitespace($_["username"]) -Or -Not [string]::IsNullOrWhitespace($_["first_name"]) -Or -Not [string]::IsNullOrWhitespace($_["last_name"])})]
        [hashtable]$PropertyMap = @{
            "first_name"="givenname"
            "last_name"="surname"
            "username"="samaccountname"
            "employee_num"="SID"
            "department"="department"
            "company"="company"
            "jobtitle"="title"
            "manager"="manager"
            "location"="physicaldeliveryofficename"
            "email"="mail"
        }
    )
    Begin {
        # Compute the given Property Map into an array for Select-Object.
        $SelectArray = $PropertyMap.GetEnumerator() | where {-Not [string]::IsNullOrWhitespace($_.Value)} | foreach { 
            $val = $_.Value
            if ($val -eq "SID") {
                @{N=$_.Name; Expression=[Scriptblock]::Create("[string]`$_.'$val'.Value") }
            } else {
                @{N=$_.Name; Expression=[Scriptblock]::Create("`$_.'$val'") }
            }
        }
        # Add the "ldap_import"="_ldap_import" mapping if it doesn't already exist.
        if (-Not $PropertyMap.ContainsKey("ldap_import")) {
            $SelectArray += @(@{N="ldap_import"; Expression=[Scriptblock]::Create("`$_._ldap_import -eq `$true")})
        }
        # Add the "activated"="_activated" mapping if it doesn't already exist.
        if (-Not $PropertyMap.ContainsKey("activated")) {
            $SelectArray += @(@{N="activated"; Expression=[Scriptblock]::Create("`$_._activated -eq `$true")})
        }
        # Add the "groups"="_groups" mapping if it doesn't already exist.
        if (-Not $PropertyMap.ContainsKey("groups")) {
            $SelectArray += @(@{N="groups"; Expression=[Scriptblock]::Create("if (`$_._groups -is [int] -Or `$_._groups -is [array]) { `$_._groups } else { `$null }")})
        }

        # Add distinguishedname in case we need to add manager references.
        $SelectArray += @("distinguishedname")
    }
    Process {
        return $User | Select $SelectArray
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
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    return -2
}

# Initialize cache if the field is defined
$cacheentities = @("users")
If ($AD_GROUP_PROPERTY_MAP.ContainsKey("company")) {
    $cacheentities += @("companies")
}
If ($AD_GROUP_PROPERTY_MAP.ContainsKey("location")) {
    $cacheentities += @("locations")
}
If ($AD_GROUP_PROPERTY_MAP.ContainsKey("department")) {
    $cacheentities += @("departments")
}
$extraParams = @{}

If ($DEBUG_HALT_ON_NULL_CACHE) {
    $extraParams.Add("ErrorOnNullEntities", $cacheentities)
}
Initialize-SnipeItCache -EntityTypes $cacheentities -Verbose @extraParams

# Fetch groups of AD users, combining them by distinguishedname
$extraParams = @{}
if ($AD_GROUP_PROPERTY_FILTER_MAP -is [hashtable]) {
    $extraParams.Add("ADPropertyFilterMap", $AD_GROUP_PROPERTY_FILTER_MAP)
}
$_props = ($AD_GROUP_PROPERTY_MAP.Values | where {$_ -ne "SID"})
$ad_users = $AD_GROUP_TARGETS | foreach { 
    if($_.groupname -is [string]) { 
        if ($_.ldap_import -is [bool]) { 
            $ldap_import = $_.ldap_import
        }
        $activated = $null
        if ($_.activated -is [bool]) { 
            $activated = $_.activated 
        }
        $groups = $null 
        if ($_.groups -is [int] -Or $_.groups -is [array]) {
            $groups = $_.groups 
        }
        
        Get-ADUsersByGroup -TargetGroup $_.groupname -Nested:$_.nested -ADProperties $_props @extraParams -Verbose | Select *,@{N="_ldap_import"; Expression={ $ldap_import }},@{N="_activated"; Expression={ $activated }},@{N="_groups"; Expression={ $groups }}
    }
} | Group-Object -Property distinguishedname | foreach {
    # Group the results by distinguishedname and merge into a new object
    if ($_.Count -eq 1) {
        [PSCustomObject]$_.Group
    } else {
        $u = [PSCustomObject]@{}
        foreach($p in ($_.Group | Select -First 1 | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name)) {
            # Get first non-null value found (if any)
            $val = $_.Group.$p | where {$_ -ne $null} | Select -First 1
            Add-Member -InputObject $u -MemberType NoteProperty -Name $p -Value $val -Force
        }
        $u
    }
}

Write-Host("[{0}] Formatting users..." -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))

# Format the users to have the properties expected by Snipe-It. Also converts the "_activated" property to "activated" and "_groups" to "groups".
$formatted_users = $ad_users | Format-UserForSyncing -PropertyMap $AD_GROUP_PROPERTY_MAP

# Null the email field if $AD_SYNC_EMAIL_FOR_LOGIN_ONLY is set and user is not activated.
if ($AD_SYNC_EMAIL_FOR_LOGIN_ONLY -And $AD_GROUP_PROPERTY_MAP.ContainsKey("email")) {
    $formatted_users = $formatted_users | Select *,@{N="email"; Expression={ if ($_.activated -eq $true) { $_.email } else { $null }}} -ExcludeProperty email
}
# Fill out the references to managers, if we're syncing it.
if (-Not [string]::IsNullOrWhitespace($AD_GROUP_PROPERTY_MAP["manager"])) {
    # Double-check a user isn't set as a manager to themselves.
    $formatted_users = $formatted_users | Select *,@{N="manager"; Expression={$manager = $_.manager; if (-Not [string]::IsNullOrWhitespace($manager) ) { if ($_.distinguishedname -eq $manager) { Write-Warning("User with username [{0}], employee_num [{1}] has self as manager, skipping adding manager reference" -f $_.username, $_.employee_num); $null } elseif (($user = $formatted_users | where {$_.distinguishedname -eq $manager} | Select -First 1) -And -Not [string]::IsNullOrWhitespace($user.username)) { $user } else { $null }}}} -ExcludeProperty "manager"
}


# Sync users with Snipe-It.
$error_count = 0
if (-Not $ENABLE_SYNC -Or $DisableSync) {
    if (-Not $ENABLE_SYNC) {
        Write-Host('Please set $ENABLE_SYNC to $true when ready to start syncing.')
        Write-Debug('Debug breakpoint due to $ENABLE_SYNC not set.')
    } else {
        Write-Host('Not syncing due to given -DisableSync switch.')
    }
} else {
    $extraParams = @{}
    if ($AD_SYNC_ON_EMPLOYEE_NUM) {
        $extraParams.Add("SyncOnEmployeeNum", $true)
    }
    Write-Host("[{0}] Starting sync..." -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
    foreach($user in $formatted_users) {
        try {
            $sp_user = Sync-SnipeItUser -User $user -Verbose @extraParams
        } catch {
            Write-Error $_
            $error_count += 1
        }
    }
}

# Create users for assigning assets to departments
if ($AD_SYNC_DEPARTMENT_USERS) {
    $extraParams = @{}
    if (-Not [string]::IsNullOrEmpty($AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)) {
        $extraParams.Add("RestrictCompany", $AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)
        $extraParams.Add("SkipEmptyCompany", $true)
    }
    Sync-SnipeItDeptUsers -SyncCompany -SkipEmptyDepartment -Verbose @extraParams
}

# Flag users that no longer exist in targeted AD groups and delete them if they have 0 assignments of all types
$inactive_users = $null
$inactive_users_undeletable = $null
$inactive_users_reassigned = $null
$inactive_users_deletable_count = 0
if ($formatted_users.Count -gt 0 -And -Not $AD_SYNC_DELETED_USERS_SKIP) {
    $_all_ldap_import = ($AD_GROUP_TARGETS | where {$_.ldap_import -eq $true}).Count -eq $AD_GROUP_TARGETS.Count
    if (-Not $AD_SYNC_ON_EMPLOYEE_NUM -And -Not $_all_ldap_import) {
        Write-Host('[{0}] Skipping deleted users -- please either set AD_SYNC_ON_EMPLOYEE_NUM or make sure all target AD groups set to use ldap_import' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
    } else {
        $duParams = @{}
        if ($AD_SYNC_ON_EMPLOYEE_NUM) {
            $duParams.Add("CompareEmployeeNum", $true)
        }
        if ($_all_ldap_import) {
            $duParams.Add("OnlyIfLdapImport", $true)
            if ($AD_SYNC_ON_EMPLOYEE_NUM) {
                $duParams.Add("AlsoCompareUsername", $true)
            }
        }
        if (-Not $ADSyncDeletedUsersPurge) {
            if ($AD_SYNC_DELETED_USERS_REPORT_ONLY) {
                Write-Host('[{0}] Will only report on inactive/deletable snipe-it users' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
                $duParams.Add("OnlyReport", $true)
            } elseif (-Not $AD_SYNC_DELETED_USERS_PURGE) {
                Write-Host('[{0}] NOT purging inactive/deletable snipe-it users' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
                $duParams.Add("DontDelete", $true)
            }
        }
        if (-Not $duParams.DontDelete -And -Not $duParams.OnlyReport) {
            Write-Host('[{0}] PURGING inactive/deletable snipe-it users' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
        }
        $inactive_users = Remove-SnipeItInactiveUsers -CompareUsers $formatted_users -Verbose @duParams 
        
        if ($inactive_users -ne $null) {
            Write-Host('[{0}] Processing inactive users' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
            If($AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT) {
                $ruParams = @{}
                If(-Not [string]::IsNullOrEmpty($AD_GROUP_PROPERTY_MAP["username"])) {
                    $ruParams.Add("ADPropertyUsername", $AD_GROUP_PROPERTY_MAP["username"])
                }
                If(-Not [string]::IsNullOrEmpty($AD_GROUP_PROPERTY_MAP["employee_num"])) {
                    $ruParams.Add("ADPropertyEmployeeNum", $AD_GROUP_PROPERTY_MAP["employee_num"])
                }
                If($AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_ONLY_DELETED) {
                    $ruParams.Add("OnlyReassignDeleted", $true)
                }
                If(-Not [string]::IsNullOrEmpty($AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)) {
                    $ruParams.Add("DepartmentalUserCompany", $AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)
                }

                $results = $null
                try {           
                    $results = Update-SnipeItInactiveUserReassignment -InactiveUsers $inactive_users -Status $AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_STATUS_ID -ExpectedCheckinDate (Get-Date) -Verbose @ruParams
                } catch {
                    Write-Error $_
                    $error_count += 1
                }
                if($results -ne $null) {
                    If($results.error_count -gt 0) {
                        $error_count += $results.error_count
                    }
                    $inactive_users_undeletable = $results.undeletable
                    $inactive_users_reassigned = $results.reassigned
                    If($inactive_users_reassigned.Count -gt 0) {
                        # Attempt to remove the reassigned users, making sure to refresh the cache.
                        $inactive_users_undeletable_2ndpass = Remove-SnipeItInactiveUsers -CompareUsers $formatted_users -Verbose -RefreshCache @duParams | where {$_.available_actions.delete -eq $false}
                        # Filter out already reassigned users.
                        $inactive_users_undeletable = $inactive_users_undeletable | where {$inactive_users_undeletable_2ndpass.id -contains $_.id}
                    }
                }
            }
            
            $inactive_users_undeletable_count = $inactive_users_undeletable.Count
            $inactive_users_deletable = $inactive_users | where {$_.available_actions.delete -eq $true} | Select -ExpandProperty username
            $inactive_users_deletable_count = $inactive_users_deletable.Count
            $inactive_users_deletable = $inactive_users_deletable -join ", "
            if (-Not [string]::IsNullOrEmpty($inactive_users_undeletable)) {
                Write-Host('[{0}] Inactive snipe-it users that no longer exist in target groups and CANNOT be deleted (still have active assignments and cannot be reassigned): {1}' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($inactive_users_undeletable._UsernameWithDept -join ", "))
            }
            if (-Not [string]::IsNullOrEmpty($inactive_users_deletable)) {
                Write-Host('[{0}] Inactive snipe-it users that no longer exist in target groups and can/have been deleted: {1}' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $inactive_users_deletable)
            }
            if (-Not [string]::IsNullOrWhiteSpace($AD_SYNC_DELETED_USERS_EXPORT_PATH)) {
                $inactive_users | Select *,@{N="_DELETABLE_"; Expression={ $_.available_actions.delete -eq $true }} | Format-SnipeItEntity | Select username,first_name,last_name,employee_num,jobtitle,department,name,location,manager,notes,* -ExcludeProperty username,first_name,last_name,employee_num,jobtitle,department,name,location,manager,notes | Export-CSV -NoTypeInformation -Force $AD_SYNC_DELETED_USERS_EXPORT_PATH
                if (Test-Path $AD_SYNC_DELETED_USERS_EXPORT_PATH -PathType Leaf) {
                    Write-Host('[{0}] Inactive user report has been saved to [{1}].' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $AD_SYNC_DELETED_USERS_EXPORT_PATH)
                }
            }
        }
    }
}

Write-Host("[{0}] Caught {1} errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

# Email out notifications
if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {    
    # Email out a report on deleted users.
    if ($EMAIL_DELETED_USERS_REPORT -Or $EmailDeletedUsersReport) {
        if (-Not [string]::IsNullOrEmpty($inactive_users_undeletable) -And $inactive_users -ne $null -And -Not [string]::IsNullOrWhiteSpace($EMAIL_DELETED_USERS_REPORT_FROM) -And -Not [string]::IsNullOrEmpty($emailDeletedUsersReportTo) -And -Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_SUBJECT)) {
            # Get all assets to check EOL dates
            $sp_assets = Get-SnipeItEntityAll "assets" -ReturnValues
            # Construct email
            $groups = $AD_GROUP_TARGETS.groupname -join ", "
            $total_count = $inactive_users.Count
            $datestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
            if (($AD_SYNC_DELETED_USERS_PURGE -Or $ADSyncDeletedUsersPurge) -And -Not $AD_SYNC_DELETED_USERS_REPORT_ONLY) {
                $deletable_action = "have been removed"
            } else {
                $deletable_action = "can be removed"
            }
            $body = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=us-ascii"><title>HTML TABLE</title>
</head><body>
<p>There are [$total_count] users in snipe-it that no longer exist in target AD group(s): ${groups}</p>
<p>[$inactive_users_deletable_count] of these users ${deletable_action}.</p>
"@
            if ($AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT) {
            $body += ("<p>[{0}] of these users had their assets reassigned to their department.</p>" -f $inactive_users_reassigned.Count)
            }
            $body += @"
<p>A user must have all their assignments checked in before they can be deleted from snipe-it. Users which cannot be deleted or reassigned:</p>
<table border="1">
"@
            If(-Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD)) {
                $body += "<tr><td>Username</td><td>Department (Last Sync)</td><td>Exists in AD</td><td>Non-EOL Assignments</td><td>Total Assignments</td></tr>"
            } else {
                $body += "<tr><td>Username</td><td>Department (Last Sync)</td><td>Exists in AD</td><td>Total Assignments</td></tr>"
            }
            # Double-check whether the user still exists in AD at all.
            foreach($user in $inactive_users_undeletable) {
                $existsInAD = $user._ExistsInAD
                if (-Not $existsInAD) { $existsInAD = "<b>False</b>" }
                $username = $user.username
                if (-Not [string]::IsNullOrEmpty($spHostURL)) {
                    $username = '<a href="{0}users/{1}">{2}</a>' -f $spHostURL, $user.id, $user.username
                }
                # Just in case one of these counts do not resolve to an integer.
                $total = $null
                $totalNonEol = $null
                try {
                    if ($user.assets_count -gt 0 -And -Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD)) {
                        $totalNonEol = ($sp_assets | where {$_.assigned_to.id -eq $user.id -And ($_.custom_fields.$EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD.value -as [DateTime]) -gt (Get-Date)} | Measure-Object).Count
                        # If greater than 0, bold the result.
                        if (-Not [string]::IsNullOrEmpty($totalNonEol) -And $totalNonEol -gt 0) {
                            $totalNonEol = '<b>{0}</b>' -f $totalNonEol
                        }
                    }
                    $total = $user.assets_count + $user.licenses_count + $user.consumables_count + $user.accessories_count
                } catch {
                    Write-Error $_
                    $total = 'ERROR'
                    $totalNonEol = 'ERROR'
                }
                # Add row for user.
                If(-Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD)) {
                    $body += ('<tr><td>{0}</td><td>{1}</td><td style="text-align: center;">{2}</td><td style="text-align: center;">{3}</td><td style="text-align: center;">{4}</td></tr>' -f $username, $user.department.name, $existsInAD, $totalNonEol, $total)
                } else {
                    $body += ('<tr><td>{0}</td><td>{1}</td><td style="text-align: center;">{2}</td><td style="text-align: center;">{3}</td></tr>' -f $username, $user.department.name, $existsInAD, $total)
                }
            }
            $body += @"
</table>

<p>A report has been saved to [<a href="file://$AD_SYNC_DELETED_USERS_EXPORT_PATH">$AD_SYNC_DELETED_USERS_EXPORT_PATH</a>].</p>

<p>This message generated on [$datestamp] from [Snipeit-AD-Sync.ps1] running on [${ENV:COMPUTERNAME}].</p>
</body></html>
"@
            Send-MailMessage -From $EMAIL_DELETED_USERS_REPORT_FROM -To $emailDeletedUsersReportTo -Subject $EMAIL_DELETED_USERS_REPORT_SUBJECT -Body $body -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $EMAIL_SMTP -BodyAsHtml
            Write-Host("[{0}] Emailed inactive user report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($emailDeletedUsersReportTo -join ", "))
        }
    }

    # Stop logging
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

    # Email out notifications of any errors.
    if ($error_count -gt 0 -And -Not [string]::IsNullOrEmpty($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrEmpty($EMAIL_ERROR_REPORT_TO)) {
        $params = @{
        From = $EMAIL_ERROR_REPORT_FROM
        To = $EMAIL_ERROR_REPORT_TO
        Subject = 'Errors from Snipeit-AD-Sync'
        Body = "There were [$error_count] caught errors from [Snipeit-AD-Sync.ps1] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
        Priority = 'High'
        DeliveryNotificationOption = @('OnSuccess', 'OnFailure')
        SmtpServer = $EMAIL_SMTP
        }
        try {
            # Attempt to send with an attachment. If that throws an error for some reason, try sending without it.
            Send-MailMessage -Attachments $_logfilepath @params
        } catch {
            Write-Error $_
            $params['Body'] = "There were [$error_count] caught errors from [Snipeit-AD-Sync.ps1] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details."
            Send-MailMessage @params
        }
        Write-Host("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($EMAIL_ERROR_REPORT_TO -join ", "))
    }
}

# Stop logging if we haven't stopped already
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null