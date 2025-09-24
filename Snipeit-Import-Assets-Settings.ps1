# Settings for Snipeit-Import-Assets.ps1.
# MJC 4-17-24

# Previously exported credentials
$CREDXML_PATH = "snipeit-creds.xml"

# Direct filepath to import assets.
$IMPORT_CSV_FILEPATH = ".\Imports\SnipeIt-Import.csv"
# A unique identifier to group the results by.
$IMPORT_CSV_GROUP_BY = "Serial"

# Mapping for Snipe-It field names.
# "SnipeItAssetField"="CSV Column Name"
# A Snipe-It field may map to more than one field in the asset.
<#
$ASSET_FIELD_MAP = @{ 
	"Serial"="Serial #"
	"Name"="Asset Name"
	"Model"="Model"
	"model_number"="Model #"
	"Manufacturer"="Manufacturer"
	"SMBIOS GUID"="SMBIOS GUID"
	"Category"="Category"
	"Fieldset"="Platform"
	"System Form Factor"="Form Factor"
	"Location"="Location"
	"Purchasing date"="Purchasing date"
	"purchase_date"="Purchasing date"
	"assigned_to"="assigned_to"
	"asset_tag"="Asset Tag"
	"supplier"="Supplier"
	"order_number"="Order #"
}
#>

# Which fields to sync if the given field matches.
# A value of $true syncs ALL mapped fields.
# Make sure these fields are all unique!
# Defaults to @{ Serial = $true }
$ASSET_FIELD_SYNC_ON_MAP = [ordered]@{
	"Serial" = $true
}
# Which fields are required to be non-blank to create the asset.
# Only checked when creating new assets.
$ASSET_FIELD_CREATE_REQUIRED = @("Serial","Model","Manufacturer","Category")
# Status used when creating. This can be the status name or ID.
$ASSET_STATUS_CREATE = "New"
# Status used when assigning assets, including newly created assets.
$ASSET_STATUS_ASSIGNED = "Ready to Deploy"

# To make doubly sure we aren't duplicating any entities, halt if the list of assets, models, manufacturers, fieldsets, categories, etc. are empty
# Ignored if not syncing the relevant fields.
$DEBUG_HALT_ON_NULL_CACHE = $true

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs\Snipeit-Sync-PS"
$LOGFILE_PREFIX = "snipeit-import-assets"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Path to copies of imported files.
$IMPORT_ARCHIVE_PATH = ".\Imports"
# Maximum number of days worth of imported files to keep.
$IMPORT_ARCHIVE_ROTATE_DAYS = 365

# Email configuration for reports.
# Output from this script will be emailed to the file owner.
$EMAIL_SMTP = 'smtp.constco.com'
$EMAIL_REPORT_FROM = 'no-reply@constco.com'
$EMAIL_REPORT_CC = @('it@constco.com')
#$EMAIL_REPORT_BCC = ''

