# Settings for Snipeit-Asset-Sync.ps1.
# MJC 4-17-24

# Set to True to start syncing.
# Still exports if set to false.
$ENABLE_SYNC = $true

# Previously exported credentials
$CREDXML_PATH = "snipeit-creds.xml"

# File or path to import assets from previously exported reports.
$IMPORT_CSV_PATH = ".\Reports\SCCM\SCCM_Asset_Export*.csv"
# A unique identifier from SCCM to group the results by.
$IMPORT_CSV_GROUP_BY = "Unique_Identifier"

# Mapping for Snipe-It field names.
# "SnipeItAssetField"="Asset Field Name"
# A Snipe-It field may map to more than one field in the asset.
$ASSET_FIELD_MAP = @{ 
	"Serial"="Serial_Number"
	"Name"="Computer_Name"
	"Model"="Model"
	"Manufacturer"="Manufacturer"
	#"SMBIOS GUID"="SMBIOS_GUID"
	#"SCCM LastActiveTime"="LastActiveTime"
	#"AD LastLogonTime"="ADLastLogonTime"
	"Category"="Platform"
	"Fieldset"="Platform"
	#"System Form Factor"="Type"
	#"AD SID"="SID"
	#"LastLogonUser"="User_Name"
	#"Primary Users"="Primary_User"
	"Location"="Location"
	#"OS Version"="OS_Build"
}

# Searchbases to optionally sync information from AD.
# $AD_IMPORT_SEARCHBASES = @("OU=Computers,DC=constco,DC=com")

# Platforms to sync from AD. Defaults to "PC"
$AD_IMPORT_PLATFORMS = @("PC", "Linux")
# Also sync entries with empty Platform if true.
$AD_IMPORT_EMPTY_PLATFORM = $true

# Platforms to sync from SCCM. Defaults to "PC"
$SCCM_IMPORT_PLATFORMS = @("PC", "Linux")
# Also sync entries with empty Platform if true.
$SCCM_IMPORT_EMPTY_PLATFORM = $true

# Which fields to sync if the given field matches.
# A value of $true syncs ALL mapped fields.
# Make sure these fields are all unique!
# Defaults to @{ Serial = $true }
$ASSET_FIELD_SYNC_ON_MAP = [ordered]@{
	Serial = $true
}
<#
$ASSET_FIELD_SYNC_ON_MAP = [ordered]@{
	"Serial" = $true
	"SMBIOS GUID" = $true
	"AD SID" = @("Name", "SCCM LastActiveTime", "LastLogonUser", "AD LastLogonTime", "PC Checkboxes")
	"Name" = @("SCCM LastActiveTime", "LastLogonUser", "AD LastLogonTime", "PC Checkboxes")
}
#>

# Which fields are required to be non-blank to create the asset.
# Only checked when creating new assets.
$ASSET_FIELD_CREATE_REQUIRED = @("Name","Serial","Model","Manufacturer","Category")
# $ASSET_FIELD_CREATE_REQUIRED = @("Name","Serial","SMBIOS GUID","Model","Manufacturer","Category")
# Status used when creating. This can be the status name or ID.
$ASSET_STATUS_CREATE = "New"
# Optional default status used to change archived assets to when encountered in sync. This can be the status name or ID.
# $ASSET_STATUS_ARCHIVED_UPDATE = "New"
# Optional default status used when assigning assets, must be deployable. This overrides other default statuses.
# $ASSET_STATUS_ASSIGNED = "Ready to Deploy"
# Include updating archived assets if they're found in the source (otherwise they're skipped).
# $ASSET_SYNC_ARCHIVED_INCLUDE = $true

# Default model name used when no model is found. This only accepts model names.
$ASSET_DEFAULT_MODEL = "_Unknown PC Model_"

# To make doubly sure we aren't duplicating any entities, halt if the list of assets, models, manufacturers, fieldsets, categories, etc. are empty
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $true

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "snipeit-asset-sync"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Filepath for exports of assets from sources and snipe-it
$EXPORTS_PATH = ".\Exports"
# Re-arrange fields in the CSV output, if they exist
# $EXPORTS_CUSTOMFIELD_ORDER = [ordered]@()
# $EXPORTS_PREFIX_SCCM_AD = "assets_sccm_ad" 	# Assets exported from SCCM and AD. Rotated on $EXPORTS_ROTATE_DAYS.
$EXPORTS_PREFIX_AD = "assets_ad" 	            # Assets exported from AD only. Rotated on $EXPORTS_ROTATE_DAYS.
$EXPORTS_PREFIX_FORMATTED = "assets_formatted"	# Assets formatted before synced with Snipe-It. Not rotated, only latest is kept.
$EXPORTS_PREFIX_SNIPEIT = "assets_snipeit"		# Assets exported from Snipe-It. Rotated on $EXPORTS_ROTATE_DAYS.
$EXPORTS_ROTATE_DAYS = 365
# $EXPORTS_PREFIX_SNIPEIT_LATEST = "assets_snipeit_latest" # Copy of the latest assets exported from Snipe-It.

# Email configuration for reports
#$EMAIL_SMTP = 'smtp.constco.com'
# If filled out, send error reports
# $EMAIL_ERROR_REPORT_FROM = 'no-reply@constco.com'
# $EMAIL_ERROR_REPORT_TO = @('it@constco.com')


