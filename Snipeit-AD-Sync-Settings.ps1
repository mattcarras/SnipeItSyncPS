# MJC 8-13-22
# Settings for Snipeit-AD-Sync.ps1.

# Previously exported credentials. Must be exported and imported under the same account.
$CREDXML_PATH = "snipeit-creds.xml"

# This must evaluate to $true to actually start syncing. Otherwise the script skips syncing entirely.
# It also gives a debug breakpoint, if you have debugging enabled.
# Note the -DisableSync switch overrides this setting.
$ENABLE_SYNC = $true

# Target group(s) of users to sync with Snipe-It.
# This should be one or more hashtables in the form of:
# 	groupname = group name or array of group names
#	nested = If True, use -Recursive lookup. May fail if >5000 members returned.
#	activated = If True, allow login for users in this group.
#	groups = int or int array of groups to assign in Snipe-It (requires SnipeItPS 1.10.225 or newer)
#	ldap_import = Set the ldap_import flag on the Snipe-It user.
$AD_GROUP_TARGETS = @(
	@{"groupname" = "Domain Users"; "nested"=$false; "ldap_import"=$true}
	#@{"groupname" = "SnipeItAdmins"; "nested"=$false; "activated"=$true; "groups"=2; "ldap_import"=$true}
)

# AD properties to sync
# "SnipeitField"="AD Property Name"
# Only these fields will sync.
# May also be given in the following alternate formats:
# -- Format 1 (Basic) --
# @{"Properties"=@("prop1","prop2"); "AllowWhitespace"=$false }
# Where Properties contains all properties in order of preference.
# The properties will be checked for non-null/whitespace values in the order given.
# -- Format 2 (Advanced) --
# @{"Properties=@("prop1","prop2"); "ScriptBlock"={ ... }}
# Where Properties contains all properties referenced in the ScriptBlock.
# This can allow different field values based on certain conditions.
<#
# Example:
# @{"Properties"=@("title","extensionattribute1","extensionattribute2"); 
	"ScriptBlock"={
			param(
				$o
			)
			If ($o.extensionattribute1 -ne 'Staff' -And -Not [string]::IsNullOrWhitespace($o.extensionattribute2)) { 
				return $o.extensionattribute2
			} else { 
				return $o.title
			}
		}
   }
#>
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

# AD Property
# Filter the results based on the given map of Properties.
# These properties do not need to be defined in the property map.
# If a hashtable, requires the "Value" and "operator" keys, where "operator" can be any operator supported by PowerShell.
# Otherwise assume the "-ne" operator by default and the value is a string.
# $AD_GROUP_PROPERTY_FILTER = "(`$_.distinguishedname -like '*,OU=Users,*')"
# List of attributes required by the filter (if not already part of the property map).
# $AD_GROUP_PROPERTY_FILTER_ATTRS = @()

# If set, don't sync companies. Only use it for filtering users.
$AD_SYNC_DONTCREATECOMPANY = $false

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
$AD_SYNC_DELETED_USERS_EXPORT_PATH = ".\Exports\snipeit_ad_deleted_users.csv"

# Reassign any equipment assigned to a deleted user to a special department user if true.
$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT = $true
# Only reassign equipment if the user was deleted from AD entirely.
$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_ONLY_DELETED = $true
# Change to the given status ID when reassigning assets if set. This status must already exist.
$AD_SYNC_DELETED_USERS_REASSIGN_TO_DEPARTMENT_STATUS_ID = 10

# Create special users for each department to allow assigning assets to departments.
$AD_SYNC_DEPARTMENT_USERS = $true
# Only create departmental users if their department exists in AD.
# Assumes Department and Company fields are mapped in AD properties.
$AD_SYNC_DEPARTMENT_USERS_FROM_AD_DEPARTMENTS = $true
# Only create special department users when the following company is set.
$AD_SYNC_DEPARTMENT_USERS_RESTRICT_COMPANY = "Constco"

# To make doubly sure we aren't duplicating any entities, halt if the list of users, depts, and/or locations are empty.
# This is useful if you know all the entities (users, departments, companies, and locations) should return at least 1 result.
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $true

# Path and prefix for the Start-Transcript logfiles.
# Note the -LogFilePrefix "<string>" parameter overrides this prefix.
$LOGFILE_PATH = ".\Logs\Snipeit-Sync-PS"
$LOGFILE_PREFIX = "snipeit-ad-sync"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Email configuration for reports
<#
$EMAIL_SMTP = 'smtp.constco.com'
# If filled out, send error reports
$EMAIL_ERROR_REPORT_FROM = 'no-reply@constco.com'
# Can be string or array of strings.
$EMAIL_ERROR_REPORT_TO = @('it@constco.com')
# You may also give the -EmailDeletedUsersReport script parameter.
# Using this in combination with the -DisableSync and -ADSyncDeletedUsersPurge script parameters allows
# for purging users and emailing out the results.
$EMAIL_DELETED_USERS_REPORT = $false
$EMAIL_DELETED_USERS_REPORT_FROM = 'no-reply@constco.com'
$EMAIL_DELETED_USERS_REPORT_SUBJECT = 'Weekly Inactive Snipe-It Users Report'
# Can be string or array of strings.
$EMAIL_DELETED_USERS_REPORT_TO = $null
# Overrides $EMAIL_DELETED_USERS_REPORT_TO
$EMAIL_DELETED_USERS_REPORT_TO_GROUPMEMBERS = 'USS-IT-SnipeItReports'
# Field to check for EOL assets.
$EMAIL_DELETED_USERS_REPORT_ASSET_EOL_CUSTOMFIELD = "End of Life"
#>
