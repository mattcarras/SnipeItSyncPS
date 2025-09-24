<#
	.SYNOPSIS
	Syncs users with Snipe-It from AD.
	
	.DESCRIPTION
	Syncs users with Snipe-It from AD. Requires RSAT, SnipeitPS, and Snipeit-Sync-PS powershell modules. Uses settings file Snipeit-AD-Sync-Settings.ps1.

	.PARAMETER DisableSync
	Disables syncing. Intended to be used with other switches. Overrides settings file.
	
	.PARAMETER ADSyncDeletedUsersPurge
	Deletes Snipe-It users that no longer exist in target AD groups and have no active assignments. Can attempt to reassign assets if set. Overrides settings file.
	
	.PARAMETER EmailDeletedUsersReport
	Emails a report of deleted users. Overrides settings file.
	
	.OUTPUTS
	Return Codes
	-1 Error loading settings file
	-2 Error loading Snipeit-Sync-PS API
	-3 Error connecting to Snipe-It site

	.NOTES
	Uses Snipeit-AD-Sync-Settings.ps1 settings file.
	
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
param([switch] $DisableSync, [switch] $ADSyncDeletedUsersPurge, [switch] $EmailDeletedUsersReport, [string] $LogFilePrefix)

# -- LOAD CONFIGURATION --
try {
	. .\Snipeit-AD-Sync-Settings.ps1
} catch {
	Write-Error $_
    return -1
}

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

if (($EMAIL_DELETED_USERS_REPORT -Or $EmailDeletedUsersReport) -And -Not [string]::IsNullOrWhitespace($EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS)) {
	Write-Host('[{0}] Compiling list of deleted users report recipients from [{1}].' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS)
	$emailDeletedUsersReportTo = Get-ADGroupMember $EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS -Recursive | foreach { Get-ADUser $_ -Properties mail | Select -ExpandProperty mail }
} else {
	$emailDeletedUsersReportTo = $EMAIL_DELETED_USERS_REPORT_TO
}
	
# -- START FUNCTIONS --	
function Get-ADUsersByGroup {
	<#
		.SYNOPSIS
		Collect all AD users from given target group(s), filtering the results.
		
		.DESCRIPTION
		Collect all AD users from given target group(s), filtering the results. If you want to check all users give a global group like Domain Users.
		
		.PARAMETER TargetGroup
        Required. The AD Group(s) to check.
		
		.PARAMETER ADProperties
        The AD properties to return with each user.
		
		.PARAMETER ADPropertyFilter
        A filterscript to use on the results. Use backticks for property references. E.g. "`$_.distinguishedname -like '*,OU=Users,*'"
		
		.PARAMETER Nested
		Will recurse over groups if given. This may take a while with large groups.
		
        .PARAMETER IncludeDisabled
        If true include disabled users.
		
		.PARAMETER ExitOnError
		Exit on error fetching group membership.
		
		.PARAMETER RecurseLoopCount
		This is used when the function is called recursively.
		
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
		[string]$ADPropertyFilter,
		
		[parameter(Mandatory=$false)]
		[switch]$Nested,

        [parameter(Mandatory=$false)]
		[switch]$IncludeDisabled,
		
		[parameter(Mandatory=$false)]
		[switch]$ExitOnError,
		
		[parameter(Mandatory=$false)]
		[int]$RecurseLoopCount=0
	)
	
	$ad_users = $null
	$props = $ADProperties
	if ($props -ne $null -And -Not $props -is [array]) {
		$props = @($props)
	}
	# We'll use the memberof property to determine if we already got this user.
	$props += @("distinguishedname","memberof") | Select -Unique
	Write-Debug "[Get-ADUsersByGroup] Properties: $props"
		
	foreach ($group in $TargetGroup) {
		# Get all users from AD
		Write-Verbose ("[Get-ADUsersByGroup] Collecting all users from AD group [$group] (Nested=$Nested, With Filter={0})..." -f (-not [string]::IsNullOrEmpty($ADPropertyFilter)))
		
		if ($Nested) {
			try {
				# May not work with >5000 results
				$ad_users += Get-ADGroupMember $group -Recursive -ErrorAction Stop | where {$_.objectClass -eq 'user'}
			} catch [System.TimeoutException],[TimeoutException] {
				Write-Warning ("[Get-ADUsersByGroup] Timeout detected. Trying again, recursing over each member. Please wait...")
				# If we have a timeout, try again recursing over each nested group found.
				# If we have a very high recurse count, assume we're in an infinite loop and throw an error.
				if ($RecurseLoopCount -gt 20) {
					$errorMsg = "Recurse count is too high ($RecurseLoopCount), may be infinite loop, aborting"
					if ($ExitOnError) {
						Write-Error $errorMsg
						exit -1
					} else {
						throw $errorMsg
					}
				}
				try {
					# Manually recurse over nested groups.
					# An alternative is using LDAP_MATCHING_RULE_IN_CHAIN, but it's quite slower.
					# Get the group info.
					$adgroup = Get-ADGroup $group
					# Get all user members of this group.
					$childUsers = Get-ADUser -LDAPFilter "(&(objectCategory=user)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -Properties $props -ErrorAction Stop
					Write-Debug("[Get-ADUsersByGroup] [group=$group] Found $($childUsers.Count) users")
					# Get all nested groups.
					$childGroups = Get-ADGroup -LDAPFilter "(&(objectCategory=group)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -ErrorAction Stop | Select -ExpandProperty Name
					Write-Debug("[Get-ADUsersByGroup] [group=$group] Found $($childGroups.Count) groups")
					# Call this function recursively for all groups found.
					if (($childGroups | Measure-Object).Count -gt 0) {
						$ad_users += Get-ADUsersByGroup -TargetGroup $childGroups -ADProperties $ADProperties -Nested -IncludeDisabled:$IncludeDisabled -ExitOnError:$ExitOnError -RecurseLoopCount ($RecurseLoopCount + 1)
					}
				} catch {
					if ($ExitOnError) {
						Write-Error $_
						exit -1
					} else {
						throw
					}
				}
			} catch {
				if ($ExitOnError) {
					Write-Error $_
					exit -1
				} else {
					throw
				}
			}
		} else {
			# No nested groups.
			try {
				$adgroup = Get-ADGroup $group
				$ad_users += Get-ADUser -LDAPFilter "(&(objectCategory=user)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -Properties $props -ErrorAction Stop
			} catch {
				if ($ExitOnError) {
					Write-Error $_
					exit -1
				} else {
					throw
				}
			}
		}
	}
    if ($ad_users -ne $null) {		
		# Get extra attributes for each user
		Write-Verbose ("[Get-ADUsersByGroup] Getting properties for {0} users..." -f ($ad_users | Measure).Count)
		# Make sure to dedupe users here.
		# Only fetch the user if they are missing the "memberof" property
		try {
			$ad_users = $ad_users | Select -Unique | foreach { if($_.memberof -ne $null) { $_ } else { Get-ADUser $_ -Properties $props } }
		} catch {
			if ($ExitOnError) {
				Write-Error $_
				exit -1
			} else {
				throw
			}
		}
		
		$filterscript = $ADPropertyFilter
		if (-Not $IncludeDisabled) {
			if (-Not [string]::IsNullOrWhitespace($filterscript)) {
				$filterscript += ' -AND '
			}
			$filterscript += "`$_.Enabled -eq `$true"
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

		.Parameter PreferredFirstNameAttr
		Attribute for preferred first name, if set.
		
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
		[ValidateScript({
			(($_["username"] -is [hashtable] -And -Not [string]::IsNullOrWhitespace(($_["username"].Properties | Select -First 1))) -Or
				($_["username"] -is [string] -And -Not [string]::IsNullOrWhitespace($_["username"]))) -And 
			(($_["first_name"] -is [hashtable] -And -Not [string]::IsNullOrWhitespace(($_["first_name"].Properties | Select -First 1))) -Or
				($_["first_name"] -is [string] -And -Not [string]::IsNullOrWhitespace($_["first_name"]))) -And 
			(($_["last_name"] -is [hashtable] -And -Not [string]::IsNullOrWhitespace(($_["last_name"].Properties | Select -First 1))) -Or
				($_["last_name"] -is [string] -And -Not [string]::IsNullOrWhitespace($_["last_name"])))
		})]
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
	}
	Process {
		return $User | foreach {
			$u = [PSCustomObject]@{}
			# Iterate over the property map.
			foreach($pair in $PropertyMap.GetEnumerator()) {
				# Iterate through Properties array in order of preference
				if ($pair.Value.Properties -is [array]) {
					if ($pair.Value.ScriptBlock -isnot [ScriptBlock]) {
						foreach($p in $pair.Value.Properties) {
							if ($p -eq "SID") {
								$val = [string]$_.$p.Value
							} elseif(-Not [string]::IsNullOrWhitespace($p) -And (-Not [string]::IsNullOrWhitespace($_.$p) -Or ($pair.Value.AllowWhitespace -And -Not [string]::IsNullOrEmpty($_.$p)))) {
								$val = ($_.$p -join ";")
							}
							break
						}
					} else {
						# Run function scriptblock
						$val = ($pair.Value.ScriptBlock.Invoke($_) -join ";")
					}
				} elseif(-Not [string]::IsNullOrWhitespace($pair.Value)) {					
					if ($pair.Value -eq "SID") {
						$val = [string]$_.($pair.Value).Value
					} else {
						$val = ($_.($pair.Value) -join ";")
					}
				}
				Add-Member -InputObject $u -MemberType NoteProperty -Name $pair.Name -Value $val -Force
			}
			
			# Add the "ldap_import"="_ldap_import" mapping if it doesn't already exist.
			if (-Not $PropertyMap.ContainsKey("ldap_import")) {
				Add-Member -InputObject $u -MemberType NoteProperty -Name "ldap_import" -Value ($_._ldap_import -eq $true) -Force
			}
			
			# Add the "activated"="_activated" mapping if it doesn't already exist.
			if (-Not $PropertyMap.ContainsKey("activated")) {
				Add-Member -InputObject $u -MemberType NoteProperty -Name "activated" -Value ($_._activated -eq $true) -Force
			}
			
			# Add the "groups"="_groups" mapping if it doesn't already exist.
			if (-Not $PropertyMap.ContainsKey("groups")) {
				$groups = $null
				if($_._groups -is [int] -Or $_._groups -is [array]) {
					$groups = $_._groups
				}
				Add-Member -InputObject $u -MemberType NoteProperty -Name "groups" -Value $groups -Force
			}
			
			# Add distinguishedname in case we need to add manager references.
			Add-Member -InputObject $u -MemberType NoteProperty -Name "distinguishedname" -Value $_.distinguishedname -Force
			
			# Return formatted user object
			$u
		}
		# return $User | Select $SelectArray
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
    return -3
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
if (-Not [string]::IsNullOrWhitespace($AD_GROUP_PROPERTY_FILTER)) {
    $extraParams.Add("ADPropertyFilter", $AD_GROUP_PROPERTY_FILTER)
}
$_props = @()
foreach ($v in $AD_GROUP_PROPERTY_MAP.Values) {
	$p = $null
	if ($v.Properties.Count -gt 0 -Or $v -is [hashtable]) {
		$_props += ($v.Properties | where {$_ -ne "SID"})
	} elseif ($v -ne "SID") {
		$_props += @($v)
	}
}

# Add in the property filter attributes, if set.
If(-Not [string]::IsNullOrWhitespace(($AD_GROUP_PROPERTY_FILTER_ATTRS | Select -First 1))) {
	$_props += $ADPropertyFilterAttrs
}
$_props = $_props | Sort -Unique

$doExitOnError = ($ADSyncDeletedUsersPurge -Or $AD_SYNC_DELETED_USERS_PURGE)
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
		
		Get-ADUsersByGroup -TargetGroup $_.groupname -Nested:$_.nested -ADProperties $_props @extraParams -ExitOnError:$doExitOnError -Verbose | Select *,@{N="_ldap_import"; Expression={ $ldap_import }},@{N="_activated"; Expression={ $activated }},@{N="_groups"; Expression={ $groups }}
	}
# Group the results by distinguishedname and merge into a new object
} | Group-Object -Property distinguishedname | foreach {
	$u = $_.Group
	# If a user is in multiple groups, merge the results
	if ($_.Count -gt 1) {
		$u = @{}
		# Loop over all properties
		foreach($p in ($_.Group | Select -First 1 | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name)) {
			$group = $_.Group
            # Get first non-null value found (if any)
			# Unless property is one of the built-in ones we've added
			switch($p) {
				"_groups" {
					$val = $group.$p | where {$_ -ne $null} | Select -Unique
				}
				($_ -in "_ldap_import","_activated") {
					$val = $true -in $group.$p
				}
				default {
					$val = $group.$p | where {$_ -ne $null} | Select -First 1
				}
			}
			$u.Add($p, $val)
		}
	}
	[PSCustomObject]$u
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
	$syncable_users = $formatted_users
    $extraParams = @{}
    if ($AD_SYNC_ON_EMPLOYEE_NUM) {
	    $extraParams.Add("SyncOnEmployeeNum", $true)
    }
	if ($AD_SYNC_DONTCREATECOMPANY) {
		# Add the -DontCreateCompanyIfNotFound parameter and exclude the company field from the list of users to sync.
		$extraParams.Add("DontCreateCompanyIfNotFound", $true)
		$syncable_users = $syncable_users | Select * -ExcludeProperty Company
	}
    Write-Host("[{0}] Starting sync for [{1}] total users..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($syncable_users | Measure).Count)
    foreach($user in $syncable_users) {
	    try {
		    $sp_user = Sync-SnipeItUser -User $user -Verbose @extraParams
	    } catch {
		    Write-Error $_
		    $error_count += 1
	    }
    }

	# Create users for assigning assets to departments
	if ($AD_SYNC_DEPARTMENT_USERS) {
		$extraParams = @{}
		Write-Host("[{0}] Syncing departmental users based on Snipe-It departments" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
		
		if($AD_SYNC_DEPARTMENT_USERS_FROM_AD_DEPARTMENTS) {
			# Assumes Department and Company fields are mapped in AD properties.
			try {
				$departments = 	$formatted_users | Select @{N="Department"; Expression={ if ($_.Department -eq $null) { $null } else { $_.Department.Trim() }}},Company | where {-not [string]::IsNullOrWhitespace($_.Department) -And ([string]::IsNullOrEmpty($AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY) -Or ($_.Company -ne $null -And $_.Company.Trim() -eq $AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY))} | Select -ExpandProperty Department -Unique
				Write-Host("[{0}] Filtering based on {1} AD departments" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($departments | Measure).Count)
				
				Sync-SnipeItDeptUsers -Departments $departments -SkipEmptyDepartment -Verbose
			} catch {
				Write-Error $_
				$error_count += 1
			}
		} else {
			if (-Not [string]::IsNullOrEmpty($AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)) {
				$extraParams.Add("RestrictCompany", $AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY)
				$extraParams.Add("SkipEmptyCompany", $true)
			}
			
			try {
				Sync-SnipeItDeptUsers -SyncCompany -SkipEmptyDepartment -Verbose @extraParams
			} catch {
				Write-Error $_
				$error_count += 1
			}
		}
	}
}

# Flag users that no longer exist in targeted AD groups and delete them if they have 0 assignments of all types
$inactive_users = $null
$inactive_users_undeletable = $null
$inactive_users_reassigned = $null
$inactive_users_deletable_count = 0
$inactive_users_reassigned_count = 0
if (($formatted_users | Measure-Object).Count -gt 0 -And -Not $AD_SYNC_DELETED_USERS_SKIP) {
    $_all_ldap_import = ($AD_GROUP_TARGETS | where {$_.ldap_import -eq $true}).Count -eq $AD_GROUP_TARGETS.Count
    if ($AD_SYNC_ON_EMPLOYEE_NUM -Or $_all_ldap_import) {
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

                $results = $null
                try {			
                    $results = Update-SnipeItInactiveUserReassignment -InactiveUsers $inactive_users -Status $AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_STATUS_ID -ExpectedCheckinDate (Get-Date) -Verbose @ruParams
					if($results -ne $null) {
						If($results.error_count -gt 0) {
							$error_count += $results.error_count
						}
						$inactive_users_undeletable = $results.undeletable
						$inactive_users_reassigned = $results.reassigned
						$inactive_users_reassigned_count = ($inactive_users_reassigned | Measure-Object).Count
						If($inactive_users_reassigned_count -gt 0) {
							# Attempt to remove the reassigned users, making sure to refresh the cache.
							$inactive_users_undeletable_2ndpass = Remove-SnipeItInactiveUsers -CompareUsers $formatted_users -Verbose -RefreshCache @duParams | where {$_.available_actions.delete -eq $false}
							# Filter out already reassigned users.
							$inactive_users_undeletable = $inactive_users_undeletable | where {$inactive_users_undeletable_2ndpass.id -contains $_.id}
						}
					}
                } catch {
                    Write-Error $_
                    $error_count += 1
                }
            }
			
            $inactive_users_undeletable_count = ($inactive_users_undeletable | Measure-Object).Count
            $inactive_users_deletable = $inactive_users | where {$_.available_actions.delete -eq $true} | Select -ExpandProperty username
            $inactive_users_deletable_count = ($inactive_users_deletable | Measure-Object).Count
            $inactive_users_deletable = $inactive_users_deletable -join ", "
            if (-Not [string]::IsNullOrEmpty($inactive_users_undeletable)) {
                Write-Host('[{0}] Inactive snipe-it users that no longer exist in target groups and CANNOT be deleted (still have active assignments and cannot be reassigned): {1}' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($inactive_users_undeletable._UsernameWithDept -join ", "))
            }
            if (-Not [string]::IsNullOrEmpty($inactive_users_deletable)) {
                Write-Host('[{0}] Inactive snipe-it users that no longer exist in target groups and can/have been deleted: {1}' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $inactive_users_deletable)
            }
            if (-Not [string]::IsNullOrWhiteSpace($AD_SYNC_DELETED_USERS_EXPORT_PATH)) {
                $inactive_users | Select *,@{N="_DELETABLE_"; Expression={ $_.available_actions.delete -eq $true }} | Format-SnipeItEntity | Export-CSV -NoTypeInformation -Force $AD_SYNC_DELETED_USERS_EXPORT_PATH
                if (Test-Path $AD_SYNC_DELETED_USERS_EXPORT_PATH -PathType Leaf) {
                    Write-Host('[{0}] Inactive user report has been saved to [{1}].' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $AD_SYNC_DELETED_USERS_EXPORT_PATH)
                }
            }
        }
    }
}

# Email out notifications
if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {    
    # Email out a report on deleted users.
    if ($EMAIL_DELETED_USERS_REPORT -Or $EmailDeletedUsersReport) {
		Write-Host('[{0}] Deleted users report is ENABLED.' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
		
        if (-Not [string]::IsNullOrEmpty($inactive_users_undeletable) -And $inactive_users -ne $null -And -Not [string]::IsNullOrWhiteSpace($EMAIL_DELETED_USERS_REPORT_FROM) -And -Not [string]::IsNullOrEmpty($emailDeletedUsersReportTo) -And -Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_SUBJECT)) {
			Write-Host('[{0}] Preparing deleted users email.' -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
			
			# Get all assets to check EOL dates
			If (-Not [string]::IsNullOrEmpty($EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD)) {
				$sp_assets = Get-SnipeItEntityAll "assets" -ReturnValues
			}
			
			# Construct email
			$groups = $AD_GROUP_TARGETS.groupname -join ", "
            $total_count = ($inactive_users | Measure-Object).Count
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
				$body += ("<p>[{0}] of these users had their assets reassigned to their department.</p>" -f $inactive_users_reassigned_count)
			}
$body += @"
<p>A user must have all their assignments checked in before they can be deleted from snipe-it. Users which cannot be deleted or reassigned:</p>
<table border="1">
<tr><td>Username</td><td>Department (Last Sync)</td><td>Exists in AD</td><td>Non-EOL Assignments</td><td>Total Assignments</td></tr>
"@
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
				$body += ('<tr><td>{0}</td><td>{1}</td><td style="text-align: center;">{2}</td><td style="text-align: center;">{3}</td><td style="text-align: center;">{4}</td></tr>' -f $username, $user.department.name, $existsInAD, $totalNonEol, $total)
			}
			$body += @"
</table>

<p>A report has been saved to [<a href="file://$AD_SYNC_DELETED_USERS_EXPORT_PATH">$AD_SYNC_DELETED_USERS_EXPORT_PATH</a>].</p>

<p>This message generated on [$datestamp] from [Snipeit-AD-Sync.ps1] running on [${ENV:COMPUTERNAME}].</p>
</body></html>
"@
            Send-MailMessage -From $EMAIL_DELETED_USERS_REPORT_FROM -To $emailDeletedUsersReportTo -Subject $EMAIL_DELETED_USERS_REPORT_SUBJECT -Body $body -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $EMAIL_SMTP -BodyAsHtml
            Write-Host("[{0}] Emailed inactive user report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($emailDeletedUsersReportTo -join ", "))
        }
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
