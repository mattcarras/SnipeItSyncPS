# Settings for Snipeit-Disable-DeletedUser-Systems.ps1.
# These settings will override settings from Snipeit-Asset-Sync-Settings.ps1.
# MJC 9-24-25

# Root of OU and searchbase for AD queries.
$ROOT_DC = "DC=constco,DC=com"
$SEARCHBASE_OU = "OU=Computers,$ROOT_DC"
# Move into this OU after disabling.
$DISABLED_COMPUTER_OU = "OU=Disabled-Computers,DC=constco,DC=com"

# Exclude the following OUs from processing. They will still show up in the report as "Excluded (Skipped)", but no action should be taken on them.
# $EXCLUDED_COMPUTER_OUS = @("OU=Computers,DC=constco,DC=com")

# A set of criteria to match from Snipe-It for deciding which computers to disable/move.
# Systems must match all criteria given.
# Field - String. The field in Snipe-It.
# Value - String. The value of the field in Snipe-It.
# SubField - String. If given, check this subfield of the given Field.
# IsCustom - Boolean. If given, the field is a custom field.
# ValueMatch - Boolean. If given, treat the value as regex using -match. Note this will be ignored if the value is an array.
# Exclude - Boolean. If given, exclude this criteria instead of including it.
# To give multiple values use regex with ValueMatch.
#
# Example with multiple values (one blank):
# { 
#    Field = "Usage"
#    Value = "^(Dedicated|)"
# 	 IsCustom = $true
#    ValueMatch = $false
#	 Exclude = $false
# }
$ASSET_DISABLE_SP_CRITERIA = @(
	# Only include systems marked with status "Assigned User Deleted".
	@{
		Field = "status_label"
		SubField = "name"
		Value = "Assigned User Deleted"
	},
	# Only match PCs.
	@{
		Field = "category"
		SubField = "name"
		Value = "PC"
	}
)
	
# If true, do not skip systems that are assigned to a user with a valid UPN (with "@" symbol).
# Generally this should not happen unless the asset was manually reassigned without updating the status.
# This also confirms if the asset was reassigned to a Departmental User.
$ASSET_CHECK_USER_DELETED_INCLUDE_UPNS = $false
# Field in Snipe-It used for ADSID. Given blank to disable.
# $ASSET_CUSTOMFIELD_ADSID = "AD SID"
# Field map of AD and Snipe-It fields, in matching order.
# If matching a custom field in Snipe-It then set SnipeItFieldIsCustom=$true
<#
$ASSET_CHECK_USER_DELETED_AD_MATCH_FIELDMAP = @(
	@{ 
		ADField="SID"
		SnipeItField=$ASSET_CUSTOMFIELD_ADSID
		SnipeItFieldIsCustom=$true
	},
	@{
		ADField="Name"
		SnipeItField="name"
		SnipeItFieldIsCustom=$false
	}
)
#>
$ASSET_CHECK_USER_DELETED_AD_MATCH_FIELDMAP = @(
	@{
		ADField="Name"
		SnipeItField="name"
		SnipeItFieldIsCustom=$false
	}
)

# Maximum number to disable/move at once. Give $null to disable.
# Used in case the data in Snipe-It is incorrect at a large scale.
$ASSET_DISABLE_THRESHOLD_MAX = 15

# Also include computers not found in AD in the report if $true.
$ASSET_REPORT_NOT_FOUND_IN_AD = $false

# Export report Settings
$EXPORT_DISABLED_COMPUTERS_REPORT_PATH = ".\Exports\snipeit_ad_disabled_computers.csv"
# Export of pending computers to disable.
$EXPORT_DISABLED_COMPUTERS_PENDING_PATH = ".\Exports\snipeit_ad_disabled_computers_pending.csv"
# Snipe-It additional fields to include in export.
<#
$EXPORT_DISABLED_COMPUTERS_EXTRA_FIELDS = @(
	@{
			Field = "Primary Users"
			IsCustom = $true
	},
	@{
			Field = "AD LastLogonTime"
			IsCustom = $true
	},
	@{
			Field = "SCCM LastActiveTime"
			IsCustom = $true
	}
)
#>

# Email Report Settings
<#
$EMAIL_DISABLED_COMPUTERS_REPORT = $true
$EMAIL_DISABLED_COMPUTERS_REPORT_TO_GROUPMEMBERS = "SnipeItAdmins"
$EMAIL_REPORT_FROM = "no-reply@constco.com"
$EMAIL_REPORT_SUBJECT = "Snipe-It Deleted User Disabled Systems Report"
$EMAIL_REPORT_PENDING_SUBJECT = "Snipe-It Pending Deleted User Disabled Systems Report"
# If enabled, also email report when no actions were taken.
$EMAIL_REPORT_ON_NO_ACTIONS = $true
#>

# Prefix for logfile.
$LOGFILE_PREFIX = "snipeit-disable-deleteduser-systems"



