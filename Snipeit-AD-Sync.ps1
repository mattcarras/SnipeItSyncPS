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

# -- START CONFIGURATION --
# Previously exported credentials from Export-SnipeItCredentials
$CREDXML_PATH = "your_exported_credentials.xml"

# This must evaluate to $true to actually start syncing. Otherwise the script skips syncing entirely.
# It also gives a debug breakpoint, if you have debugging enabled.
$ENABLE_SYNC = $false

# Target group(s) of users to sync with Snipe-It.
# This should be one or more hashtables in the form of:
# 	groupname = group name or array of group names
#	nested = If $true, use -Recursive lookup. May fail if >5000 members returned.
#	activated = If $true, allow login for users in this group.
#	groups = int or int array of groups to assign in Snipe-It (requires SnipeItPS 1.10.225 or newer)
#	ldap_import = Set the ldap_import flag on the Snipe-It user.
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
# It's recommended to have either $AD_SYNC_ON_EMPLOYEE_NUM set to $true or set all your groups with ldap_import=$true.
$AD_SYNC_PURGE_DELETED_USERS = $false

# Flag deleted users by removing ldap_import (if set) and disabling login.
# This will be done if they are NOT purged.
$AD_SYNC_FLAG_DELETED_USERS = $false

# Create special users for each department to allow assigning assets to departments.
$AD_SYNC_DEPARTMENT_USERS = $true

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
$EMAIL_FROM_ERROR_REPORT = '<from address>'
$EMAIL_TO_ERROR_REPORT = '<to address>'
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
			$ad_users += Get-ADGroupMember $group -Recursive -ErrorAction Stop
		} else {
			$ad_users += Get-ADGroup $group -Properties Member -ErrorAction Stop | Select -ExpandProperty Member
		}
	}
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

# Dot-source custom sync API
try {
	. .\SnipeIt-Sync-PS.ps1
} catch {
	# Fatal error, exit
	Write-Error $_
	Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
	return -1
}

try {
	Connect-SnipeIt -CredXML $CREDXML_PATH -Verbose
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
$extraParams = @{}
if ($AD_SYNC_ON_EMPLOYEE_NUM) {
	$extraParams.Add("SyncOnEmployeeNum", $true)
}

$error_count = 0
if (-Not $ENABLE_SYNC) {
	Write-Host('Please set $ENABLE_SYNC to $true when ready to start syncing.')
	Write-Debug('Debug breakpoint due to $ENABLE_SYNC not set.')
} else {
	Write-Host("[{0}] Starting sync..." -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")))
	foreach($user in $formatted_users) {
		try {
			$sp_user = Sync-SnipeItUser -User $user -Verbose @extraParams
		} catch {
			Write-Error $_
			$error_count += 1
		}
	}
	
	# Create users for assigning assets to departments
	if ($AD_SYNC_DEPARTMENT_USERS) {
		Sync-SnipeItDeptUsers -SyncCompany -SkipEmptyDepartment -Verbose
	}

	# Flag users that no longer exist in targeted AD groups and delete them if they have 0 assignments of all types
	if ($AD_SYNC_ON_EMPLOYEE_NUM -Or $AD_GROUP_TARGETS_FLAG_LDAP_IMPORT) {
		$extraParams = @{}
		if ($AD_SYNC_ON_EMPLOYEE_NUM) {
			$extraParams.Add("CompareEmployeeNum", $true)
		}
		if (($AD_GROUP_TARGETS | where {$_.ldap_import -eq $true}).Count -eq $AD_GROUP_TARGETS.Count) {
			$extraParams.Add("OnlyIfLdapImport", $true)
			if ($AD_SYNC_ON_EMPLOYEE_NUM) {
				$extraParams.Add("AlsoCompareUsername", $true)
			}
		}
		if (-Not $AD_SYNC_PURGE_DELETED_USERS) {
			$extraParams.Add("DontDelete", $true)
		}
		if (-Not $AD_SYNC_FLAG_DELETED_USERS) {
			$extraParams.Add("OnlyReport", $true)
		}
		$inactive_users = Remove-SnipeItInactiveUsers -CompareUsers $formatted_users @extraParams -Verbose
		# DEBUG - need to send this in a report
		if ($inactive_users.Count -gt 0) {
			Write-Host("Inactive snipe-it users that no longer exist in AD: {0}" -f ($inactive_users.username -join ", "))
		}
	}
}

Write-Host("[{0}] Caught {1} errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

# Stop logging, in case you're running from console.
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

# Email out notifications of any errors.
if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {
	if ($error_count -gt 0 -And -Not [string]::IsNullOrWhiteSpace($EMAIL_FROM_ERROR_REPORT) -And -Not [string]::IsNullOrWhiteSpace($EMAIL_TO_ERROR_REPORT)) {
		Send-MailMessage -From $EMAIL_FROM_ERROR_REPORT -To $EMAIL_TO_ERROR_REPORT -Subject 'Errors from Snipeit-AD-Sync' -Body "There were [$error_count] caught errors from Snipeit-AD-Sync.ps1 running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details." -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $EMAIL_SMTP
	}
}