# High-level implementation of Snipe-It API and helper functions for syncing users and assets.
# Requirements:
# * SnipeItPS module: https://github.com/snazy2000/SnipeitPS
# 
# Install-Module SnipeitPS
# Update-Module SnipeitPS
#
# Author: Matthew Carras
# Source: https://github.com/mattcarras/SnipeItSyncPS

# Status codes to retry on error.
# Sometimes got 429 or 422 from self-hosted solutions.
# Leave this empty if you want to skip retrying.
$SNIPEIT_RETRY_ON_STATUS_CODES = @("429","Too Many Requests","422","Unprocessable Entity")

# https://github.com/snazy2000/SnipeitPS
# We use this later to access private functions.
$_SNIPEITPSFEATURES = Import-Module SnipeitPS -PassThru

function Export-SnipeItCredentials {
    <#
        .SYNOPSIS
        Export SnipeIt API credentials to an encrypted XML file.
        
        .DESCRIPTION
        Export SnipeIt API credentials to an encrypted XML file. Returns the credentials on success.
        
        .PARAMETER File
        Required. The filename and path to export to.
        
        .PARAMETER URL
        Required. The API URL for your SnipeIt instance.
        
        .PARAMETER APIKey
        Required. The API Key for your SnipeIt instance.
        
        .OUTPUTS
        The credentials converted into type "System.Management.Automation.PSCredential".
        
        .Example
        PS> Export-SnipeItCredentials -File "ap_creds.xml" -URL $URL -APIKey $APIKey
    #>
    param (
        [parameter(Mandatory=$true, Position=0)]
        [System.IO.FileInfo]$File,
        
        [parameter(Mandatory=$true, Position=1)]
        [System.URI]$URL,
        
        [parameter(Mandatory=$true, Position=2)]
        [string]$APIKey
    )
    $creds = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $URL,(ConvertTo-SecureString -String $APIKey -AsPlainText -Force)
    $creds | Export-Clixml $File
    return $creds
}

function Connect-SnipeIt {
    <#
        .SYNOPSIS
        Initializes a new connection to the given Snipe-It instance.
        
        .DESCRIPTION
        Initializes a new connection to the given Snipe-It instance. Also initializes the cache.
        
        .PARAMETER CredXML
        Required. The filename of the credential XML file.
        
        .PARAMETER IgnoreSelfSignedCert
        Dummies out a function to allow connecting to an instance with self-signed certificates.
        
        .OUTPUTS
        Any returns from Connect-SnipeItPS.
        
        .Example
        PS> Connect-SnipeIt -CredXML "creds.xml"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [alias('CredentialXML')]
        [System.IO.FileInfo]$CredXML,
        
        [parameter(Mandatory=$false)]
        [switch]$IgnoreSelfSignedCert
    )
    $creds = Import-CliXml $CredXML
    
    # Initialize cache
    $script:_SnipeItCache = @{}
    foreach ($key in @("departments","locations","companies","manufacturers","categories","models","fieldsets","statuslabels","users","assets","fields")) {
        $script:_SnipeItCache["sp_${key}"] = $null
    }
    
    # Ignore self-signed certificates by dummying out a function in .Net.
    if ($IgnoreSelfSignedCert -And -not("dummy" -as [type])) {
        Write-Verbose "Ignoring self-signed certs"
        add-type -TypeDefinition @"
        using System;
        using System.Net;
        using System.Net.Security;
        using System.Security.Cryptography.X509Certificates;

        public static class Dummy {
            public static bool ReturnTrue(object sender,
                X509Certificate certificate,
                X509Chain chain,
                SslPolicyErrors sslPolicyErrors) { return true; }

            public static RemoteCertificateValidationCallback GetDelegate() {
                return new RemoteCertificateValidationCallback(Dummy.ReturnTrue);
            }
        }
"@
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = [dummy]::GetDelegate()
    }

    Write-Verbose ("[{0}] Connecting to Snipe-It instance at [{1}]" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"),$creds.username)

    return Connect-SnipeItPS -siteCred $creds
}

function Get-SnipeItCustomFieldMap {
    <#
        .SYNOPSIS
        Return a map of custom field name => {custom field object} or database custom field name => {custom field object} with the -DBName switch.
        
        .DESCRIPTION
        Return a map of custom field name => {custom field object} or database custom field name => {custom field object} with the -DBName switch. This is used internally.
        
        .PARAMETER Fields
        Fields to evaluate. Internally all custom fields are given. You could also give a fieldset's subset of fields.
        
        .PARAMETER DBName
        Returns a field mapped to custom field database names instead of display names.
        
        .OUTPUTS
        A reference to the field map.
        
        .Example
        PS> $fieldMap = Get-SnipeItCustomFieldMap
    #>
    param ( 
        [parameter(Mandatory=$false,
                   Position=0)]
        [object[]]$Fields,
        
        [parameter(Mandatory=$false)]
        [switch]$DBName
    )
    $fieldMap = @{}
    if ($DBName) {
        $cacheKey = 'dbFieldMap'
    } else {
        $cacheKey = 'fieldMap'
    }
    if ($Fields.Count -eq 0 -And $script:_SnipeItCache.Count -gt 0) {
        if ($script:_SnipeItCache.$cacheKey.Count -eq 0) {
            $_fields = Get-SnipeItEntityAll "fields"
            if ($_fields.Count -gt 0) {
                $fieldMap = Get-SnipeItCustomFieldMap ($_fields.Values | foreach { $_ }) -DBName:$DBName
            }
        } else {
            $fieldMap = $script:_SnipeItCache.$cacheKey
        }
    } else {
        if ($DBName) {
            $groupBy = 'db_column_name'
        } else {
            $groupBy = 'Name'
        }
        $Fields | Group-Object -Property $groupBy | foreach {
            if ($_.Count -gt 1) {
                # Should never get here
                Write-Warning ("[Get-SnipeItCustomFieldMap] Cannot map custom field name [{0}] due to it matching {1} names" -f $_.Name, $_.Count)
            } else {
                $fieldMap[$_.Name] = $_.Group
            }
        }
    }
    if ($script:_SnipeItCache.Count -gt 0 -And ($script:_SnipeItCache.$cacheKey.Count -ge $fieldMap.Count -Or $Fields.Count -eq 0)) {
        $script:_SnipeItCache.$cacheKey = $fieldMap
    }
    return $fieldMap
}

function Restore-SnipeItAssetCustomFields {
    <#
        .SYNOPSIS
        Restores the custom_fields object inside of an asset.
        
        .DESCRIPTION
        Restores the custom_fields object inside of an asset, in case it's missing (can the case for assets returned from New- or Set- operations). This is a workaround for snipe-it issue #11725. This is used internally before the cache is updated.
        
        .PARAMETER Asset
        Required. The asset object to process.
        
        .OUTPUTS
        The same asset object with the custom_fields object restored.
        
        .Example
        PS> $assets = $Assets | Restore-SnipeItAssetCustomFields
    #>
    param ( 
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object]$Asset
    )
    Begin {
        $fieldMap = Get-SnipeItCustomFieldMap -DBName
    }
    Process {
        if ($Asset -is [PSObject] -And $fieldMap -ne $null -And $Asset.custom_fields -eq $null -And ($Asset | Get-Member -MemberType NoteProperty | where {$_.Name -eq "custom_fields"}).Count -eq 0) {
            $customdbfields = $Asset | Get-Member -MemberType NoteProperty | where {$_.Name -like "_snipeit_*"} | Select -ExpandProperty Name | Out-String -Stream
            if ($customdbfields.Count -gt 0) {
                $customfields = $customdbfields | foreach {
                    $fieldinfo = $fieldMap[$_]
                    if (-Not [string]::IsNullOrEmpty($fieldinfo.Name)) {
                        [PSCustomObject]@{
                            $fieldinfo.Name = [PSCustomObject]@{
                                field = $_
                                value = $Asset.$_
                                field_format = $fieldinfo.format
                                field_element = $fieldinfo.type
                            }
                        }
                    }
                }
                if ($customfields.Count -gt 0 -And $customdbfields.Count -gt 0) {
                    $Asset = $Asset | Select *,@{N="custom_fields"; Expression={ $customfields }} -ExcludeProperty $customdbfields
                }
            }
        }
        return $Asset
    }
    End {
    }
}
    
function Get-SnipeItEntityAll {
    <#
        .SYNOPSIS
        Returns all of the given entity types, warning about possible dupes.
        
        .DESCRIPTION
        Returns all of the given entity types, warning about possible dupes. The primary key is run through HtmlDecode to deal with possible HTML entities. Returns cached results depending on age of cache. Valid results will always be in a form of a hashtable indexed by ID cast as a string.
        
        .PARAMETER EntityType
        Required. One of the valid entity types. This is always in the form of their API name (IE, "departments"), with the exception of "assets". Currently supports: "departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets".
        
        .PARAMETER UsersKey
        The key to check for dupes when returning all "users". This is also the key which is run through HtmlDecode and trimmed in advance.

        .PARAMETER ExcludeArchivedAssets
        For assets, exclude archived status type from results.

        .PARAMETER MaxCacheMin
        Number of Minutes from the cached entity's last update to get a fresh copy from the Snipe-It instance. If this is not given, it defaults to the last value set. If no value was ever given, it defaults to 120 minutes (2 hours). A value of 0 will always get a fresh copy each time, through the -NoCache switch can also be given to most parameters.
        
        .PARAMETER RefreshCache
        Always refresh the cache from the Snipe-It instance.

        .PARAMETER ErrorOnDupe
        Throw an error instead of giving a warning on dupes.

        .OUTPUTS
        A hashtable of snipe-it entities indexed by ID cast as a string.
        
        .Notes
        All names / keys are run through HTMLDecode and trimmed.
        
        The dupe-checking is mostly helpful for entities that currently do not constrain uniqueness in names, such as departments, or for the employee_num field in users.

        Possible custom thrown exceptions: [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> $sp_departments = Get-SnipeItEntityAll "departments"
    #>
    param ( 
        [parameter(Mandatory=$true,
                   Position=0)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets")]
        [string]$EntityType,
        
        [parameter(Mandatory=$false)]
        [ValidateSet("username","employee_num")]
        [string]$UsersKey="username",

        [parameter(Mandatory=$false)]
        [switch]$ExcludeArchivedAssets,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [alias('MaxCacheMinutes')]
        [Nullable[int]]$MaxCacheMin,
        
        [parameter(Mandatory=$false)]
        [switch]$RefreshCache,
        
        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe
    )
    # Primary key to double-check for duplicate values
    switch ($EntityType) {
        "users" {
            $primaryKey = $UsersKey
        }
        "assets" {
            $primaryKey = "asset_tag"
        }
        default {
            $primaryKey = "name"
        }
    }
    
    $cache_key = "sp_${EntityType}"
    $getFunc = $null
    switch($EntityType) {
        "departments" {
            $getFunc = "Get-SnipeitDepartment"
            $getParams = @{ All = $true }
        }
        "locations" {
            $getFunc = "Get-SnipeitLocation"
            $getParams = @{ All = $true }
        }
        "companies" {
            $getFunc = "Get-SnipeitCompany"
            $getParams = @{ All = $true }
        }
        "models" {
            $getFunc = "Get-SnipeitModel"
            $getParams = @{ All = $true }
        }
        "manufacturers" {
            $getFunc = "Get-SnipeitManufacturer"
            $getParams = @{ All = $true }
        }
        "categories" {
            $getFunc = "Get-SnipeitCategory"
            $getParams = @{ All = $true }
        }
        "fieldsets" {
            $getFunc = "Get-SnipeitFieldset"
            $getParams = @{}
        }
        "fields" {
            $getFunc = "Get-SnipeitCustomField"
            $getParams = @{}
        }
        "statuslabels" {
            $getFunc = "Get-SnipeitStatus"
            $getParams = @{ All = $true }
        }
        "models" {
            $getFunc = "Get-SnipeitModel"
            $getParams = @{ All = $true }
        }
        "users" {
            $getFunc = "Get-SnipeitUser"
            $getParams = @{ All = $true }
        }
        "assets" {
            $getFunc = "Get-SnipeitAsset"
            $getParams = @{ All = $true }
        }
        "suppliers" {
            $getFunc = "Get-SnipeitSupplier"
            $getParams = @{ All = $true }
        }
        default {
            # Should never get here
            Throw [System.Management.Automation.ValidationMetadataException] "[Get-SnipeItEntityAll] Unsupported EntityType: $EntityType (should never get here?)"
        }
    }
    $sp_entities = $null
    if ($script:_SnipeItCache.Count -gt 0) {
        $sp_entities = $script:_SnipeItCache.$cache_key
    }
    # Refresh the cache if either including archived results or cache is older than MaxCacheMin.
    if ($sp_entities.ht.Count -eq 0 -Or $RefreshCache -Or ($ExcludeArchivedAssets -ne $true -And $sp_entities.excludeArchived -ne $ExcludeArchivedAssets) -Or $sp_entities.maxAgeDate -isnot [DateTime] -Or (Get-Date) -ge $sp_entities.maxAgeDate) {
        Write-Verbose "[Get-SnipeItEntityAll] Collecting all existing [$EntityType] from Snipe-It..."
        # TODO: Suggest update to SnipeitPS suppress these warnings
        $sp_entities = &$getFunc @getParams -WarningAction SilentlyContinue
        if (-Not [string]::IsNullOrWhitespace($sp_entities.StatusCode)) {
            Throw [System.Net.WebException] ("[Get-SnipeItEntityAll] Fatal ERROR fetching all current snipe-it [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $EntityType,$sp_entities.StatusCode,$sp_entities.StatusDescription)
        } else {
            # If assets, include archived in results unless -ExcludeArchived is given.
            if ($EntityType -eq 'assets' -And -Not $ExcludeArchivedAssets) {
                $getParams = @{ 
                    All = $true 
                    status = 'Archived'
                }
                $sp_entities_archived = &$getFunc @getParams -WarningAction SilentlyContinue
                if (-Not [string]::IsNullOrWhitespace($sp_entities_archived.StatusCode)) {
                    Throw [System.Net.WebException] ("[Get-SnipeItEntityAll] Fatal ERROR fetching all archived snipe-it [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $EntityType,$sp_entities_archived.StatusCode,$sp_entities_archived.StatusDescription)
                } else {
                    # Merge results with rest of assets.
                    if ($sp_entities_archived -ne $null) {
                        if ($sp_entities -eq $null) {
                            $sp_entities = $sp_entities_archived
                        } else {
                            if ($sp_entities_archived -isnot [array]) {
                                $sp_entities_archived = @($sp_entities_archived)
                            }
                            if ($sp_entities -isnot [array]) {
                                $sp_entities = @($sp_entities)
                            }
                            $sp_entities += $sp_entities_archived
                        }
                    }
                }    
            }

            # Decode HTMLEntity and trim results
            $sp_entities = $sp_entities | Select @{N=$primaryKey; Expression={([System.Net.WebUtility]::HtmlDecode($_.$primaryKey)).Trim()}},* -ExcludeProperty $primaryKey

            # Check and warn about dupes
            $sp_entities | where {[string]::IsNullOrEmpty($_.$primaryKey) -ne $true} | Group-Object -Property $primaryKey | where {$_.Count -gt 1} | foreach {
                if ($ErrorOnDupe) {
                    Throw [System.Data.DuplicateNameException] ("[Get-SnipeItEntityAll] Got back {0} non-unique [{1}] where [{2}]=[{3}] and -ErrorOnDupe is set" -f $_.Count,$EntityType,$primaryKey,($_.Group.$primaryKey | Select -First 1))
                } else {
                    Write-Warning ("[Get-SnipeItEntityAll] Found {0} non-unique [{1}] where [{2}]=[{3}]" -f $_.Count,$EntityType,$primaryKey,($_.Group.$primaryKey | Select -First 1))
                }
            }

            # Set maximum cache age
            $_maxCacheMin = $MaxCacheMin
            if ($MaxCacheMin -is [int]) {
                $_maxCacheMin = $MaxCacheMin
            } elseif ($sp_entities.maxCacheMin -is [int]) {
                $_maxCacheMin = $sp_entities.maxCacheMin
            } else {
                $_maxCacheMin = 120
            }
                
            # Convert to hash table grouped by ID for easy updating.
            # maxAgeDate is used to know when to refresh the date
            $sp_entities = @{ 
                ht = ($sp_entities | Group-Object -Property id -AsHashTable -AsString) 
                maxCacheMin = $_maxCacheMin
                maxAgeDate = (Get-Date).AddMinutes($_maxCacheMin)
                excludeArchived = $ExcludeArchivedAssets
            }
            
            if ($script:_SnipeItCache.Count -gt 0) {
                $script:_SnipeItCache.$cache_key = $sp_entities
            }
            
            Write-Verbose ("[Get-SnipeItEntityAll] Got back {0} results for [$EntityType]" -f $sp_entities.ht.Count)
        }
    }
    return $sp_entities.ht
}


function Update-SnipeItCache {
    <#
        .SYNOPSIS
        Update the internal cache of Snipe-It entities.
        
        .DESCRIPTION
        Update the internal cache of Snipe-It entities. This function is called internally. Returns $true on success.
        
        .PARAMETER Entity
        Required. Either the entity to update or the ID of the entity to remove.
        
        .PARAMETER EntityType
        Required. One of the valid entity types. This is always in the form of their API name (IE, "departments"), except for "assets". Currently supports "departments", "locations","companies", "manufacturers", "categories", "suppliers", "fieldsets", "fields", "statuslabels", "models", "users", "assets".

        .PARAMETER Remove
        Assume the given entity is an ID and remove it from the cache.

        .OUTPUTS
        True on success, false otherwise.
        
        .Notes
        If the EntityType is "assets", it will call Restore-SnipeItAssetCustomFields to attempt to restore the custom_fields object if it's missing.
        
        .Example
        PS> $success = Update-SnipeItCache $sp_user "users"
    #>
    param (
        [parameter(Mandatory=$true,
                       Position=0,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
        $Entity,

        [parameter(Mandatory=$true, Position=1)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets")]
        [string]$EntityType,
        
        [parameter(Mandatory=$false, Position=2)]
        [switch]$Remove
    )
    Begin {
        try {
            $sp_entities = Get-SnipeItEntityAll $EntityType
        } catch [System.Net.WebException] {
            # Try to continue if we can't get cache.
            Write-Error $_
        }
    }
    Process {
        if ($sp_entities -ne $null) {
            if ($Remove) {
                if ($Entity -is [int] -And $sp_entities.ContainsKey([string]$Entity)) {
                    $sp_entities.Remove([string]$Entity)
                    return $true
                }
            } elseif ([string]::IsNullOrWhitespace($Entity.StatusCode)) {
                if ($Entity.id.Count -gt 1) {
                    Write-Warning "[Update-SnipeItCache] Not updating [$EntityType] cache as we got more than 1 entity"
                } elseif ($Entity.id -is [int]) {
                    # If EntityType is Assets, make sure the custom_fields array is filled out properly
                    if ($EntityType -eq "assets") {
                        $sp_entity = $Entity | Restore-SnipeItAssetCustomFields
                    } else {
                        $sp_entity = $Entity
                    }
                    $sp_entities[[string]$Entity.id] = $sp_entity
                    return $true
                }
            }
        }
        return $false
    }
    End {
    }
}

function Initialize-SnipeItCache {
    <#
        .SYNOPSIS
        Refreshes the given internal Snipe-It caches (default: ALL).
        
        .DESCRIPTION
        Refreshes the given internal Snipe-It caches (default: ALL).
        
        .PARAMETER EntityTypes
        One or more valid entity types. This is always in the form of their API name (IE, "departments"), except for "assets". This defaults to fetching ALL entity types.
        
        Supports: "departments", "locations", "companies", "manufacturers", "categories", "suppliers", "fieldsets", "fields", "statuslabels", "models", "users", "assets"

        .PARAMETER ErrorOnNullEntities
        Throw an error if one or more of the given entities comes back as null, instead of displaying a warning.

        .PARAMETER MaxCacheMin
        Number of Minutes from each cached entity's last update to get a fresh copy from the Snipe-It instance. If this is not given, it defaults to the last value set. If no value was ever given, it defaults to 120 minutes (2 hours). A value of 0 will always get a fresh copy each time, through the -NoCache switch can also be given to most parameters.

        .PARAMETER SleepSecs
        Number of seconds to sleep between each call (Defaukt: 5).

        .PARAMETER UsersKey
        The key to check for dupes when returning all "users". This is also the key which is run through HtmlDecode.

        .OUTPUTS
        None.
        
        .Notes
        Possible custom thrown exceptions: [System.Data.NoNullAllowedException]
        Possible custom thrown exceptions from Get-SnipeItEntityAll: [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Initialize-SnipeItCache @("users", "departments", "companies", "locations")
    #>
    param (
        [parameter(Mandatory=$false,
                       Position=0,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets")]
        [string[]]$EntityTypes=@("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets"),
        
        [parameter(Mandatory=$false)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","users","assets")]
        [string[]]$ErrorOnNullEntities,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [alias('MaxCacheMinutes')]
        [Nullable[int]]$MaxCacheMin,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [alias('SleepSeconds')]
        [int]$SleepSecs=5,

        [parameter(Mandatory=$false)]
        [ValidateSet("username","employee_num")]
        [string]$UsersKey="username"
    )

    foreach ($entitytype in $EntityTypes) {
        $pp = @{}
        if ($Verbose) {
            $pp.Add("Verbose", $true)
        }
        if ($Debug) {
            $pp.Add("Debug", $true)
        }
        if ($entitytype -eq "users") {
            $pp.Add("UsersKey", $UsersKey)
        }
        if ($MaxCacheMin -is [int]) {
            $pp.Add("MaxCacheMin", $MaxCacheMin)
        }
        try {
            $results = Get-SnipeItEntityAll -EntityType $entitytype -RefreshCache @pp
        } catch [System.Net.WebException] {
            # Try to continue if we can't get cache.
            Write-Error $_
        }
        if ($results -eq $null) {
            if ($entitytype -in $ErrorOnNullEntities) {
                Throw [System.Data.NoNullAllowedException] ("[Initialize-SnipeItCache] No valid [$entitytype] was returned and -ErrorOnNullEntities is set (have any of these entity types been created yet?).")
            } else {
                Write-Warning "[Initialize-SnipeItCache] No valid [$entitytype] was returned (have any of these entity types been created yet?)."
            }
        }
        Start-Sleep $SleepSecs
    }
}

function Get-SnipeItApiEntityByName {
    <#
        .SYNOPSIS
        Returns the given entity by name.
        
        .DESCRIPTION
        Returns the given entity by name. This is required until the SnipeItPS module is updated. This function is called internally, there are Get-SnipeIt<Entity>ByName functions for all supported entities (as well as Get-SnipeItAssetEx and Sync-SnipeItAsset for assets).
        
        .PARAMETER Name
        Required. The name of the entity.

        .PARAMETER EntityType
        Required. The type of entity supported by the Snipe-It API. This is always in the form of their API name (IE, "departments"), including "hardware" for assets.
        
        Supported types: "departments", "locations", "companies", "manufacturers", "categories", "suppliers", "fieldsets", "fields", "statuslabels", "models", "hardware".

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it entity object.
        
        .Notes
        Possible custom thrown exceptions: [System.Net.WebException]

        .Example
        PS> Get-SnipeItApiEntityByName "IT" "departments"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$true,
                   Position=1)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","suppliers","fieldsets","fields","statuslabels","models","hardware")]
        [string]$EntityType,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
    }
    Process {
        $sp_entity = $null
        $count_retry = $OnErrorRetry
        while ($count_retry -ge 0) {
            # TODO: Suggest updating SnipeIt API to add name filtering to models, fieldsets, and fields.
            switch($EntityType) {
                "models" {
                    $sp_entity = Get-SnipeItModel -search $name | where {$_.Name -eq $name}
                }
                "fieldsets" {
                    $sp_entity = Get-SnipeItFieldset | where {$_.Name -eq $name}
                }
                "fields" {
                    $sp_entity = Get-SnipeitCustomField | where {$_.Name -eq $name}
                }
                default {
                    # Workaround for SnipeitPS pull request #279
                    $sp_entity = & $_SNIPEITPSFEATURES { param($name,$Api) Invoke-SnipeitMethod -Api "/api/v1/$Api" -Method "GET" -GetParameters @{ name=$name} } $name $EntityType.ToLower()
                }
            }
            if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                $count_retry--
                Write-Warning ("[Get-SnipeItApiEntityByName] ERROR getting snipeit $EntityType by name [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $Name,$sp_entity.StatusCode,$sp_entity.StatusDescription,$count_retry)
            } else {
                # Break out of loop early on anything except "Too Many Requests"
                $count_retry = -1
            }
            # Sleep before next API call
            Start-Sleep -Milliseconds $SleepMS
        }
        if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
            Throw [System.Net.WebException] ("[Get-SnipeItApiEntityByName] Fatal ERROR getting snipeit {0} by name [{1}]! StatusCode: {2}, StatusDescription: {3}" -f $EntityType,$Name,$sp_entity.StatusCode,$sp_entity.StatusDescription)
        }
        return $sp_entity
    }
    End {
    }
}

function Select-SnipeItFilteredEntity {
    <#
        .SYNOPSIS
        Return a set of snipe-it entities filtered by one or more of the given parameters.
        
        .DESCRIPTION
        Return a set of snipe-it entities filtered by one or more of the given parameters. This is used internally when filtering by parameters other than name.
        
        .PARAMETER Entity
        Required. One or more snipe-it entity objects to filter.

        .PARAMETER Params
        Required. A hashtable of parameter values to filter. If the key name is in the form '<parameter>_id', check against '$_.<parameter>.id'. If the key name is in the form '<parameter>_<<name>>', check against '$_.<parameter>.name'.

        .PARAMETER All
        Filter by all parameters given (using -AND), instead of looking for matches one at a time.

        .OUTPUTS
        The filtered set of snipe-it entity objects.

        .Example
        PS> Select-SnipeItFilteredEntity $sp_entities @{ 'model_id' = 1 }
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object[]]$Entity,
        
        [parameter(Mandatory=$true, Position=1)]
        [hashtable]$Params,
        
        [parameter(Mandatory=$false)]
        [switch]$All
    )
    Begin {
    }
    Process {
        $_entity = $Entity
        if ($_entity.Count -gt 1) {
            $filterScript = $null
            foreach($pair in $Params.GetEnumerator()) {
                if ($pair.Name -ne "image" -And $pair.Name -ne "image_delete") {
                    if ($pair.Name -like '*_id') {
                        $name = '{0}.id' -f $pair.Name
                    } elseif ($pair.Name -like '*_<name>') {
                        $name = '{0}.name' -f $pair.Name
                    } else {
                        $name = $pair.Name
                    }
                    if (-Not [string]::IsNullOrEmpty($pair.Value)) {
                        if ($pair.Value -is [string]) {
                            $f = '[System.Net.WebUtility]::HtmlDecode($_.{0}).Trim() -eq "{1}"' -f $name, $pair.Value.Trim()
                        } elseif ($pair.Value -is [bool]) {
                            $f = '$_.{0} -eq ${1}' -f $name, $pair.Value
                        } elseif ($pair.Value -is [int]) {
                            $f = '$_.{0} -eq {1}' -f $name, $pair.Value
                        }
                    }
                    if ($All) {
                        if ([string]::IsNullOrEmpty($filterscript)) {
                            $filterScript = $f
                        } else {
                            $filterScript = "$filterScript -AND $f"
                        }
                    } elseif (-Not [string]::IsNullOrWhitespace($f)) {
                        $_entity = $Entity | Where-Object -FilterScript ([scriptblock]::create($f))
                        if ($_entity.Count -le 1) {
                            break
                        }
                    }
                }
            }
            if (-Not [string]::IsNullOrWhitespace($filterScript)) {
                $_entity = $Entity | Where-Object -FilterScript ([scriptblock]::create($filterScript))
            }
        }
        return $_entity
    }
    End {
    }
}

function Get-SnipeItEntityByName {
    <#
        .SYNOPSIS
        Returns the given entity by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given entity by name, optionally creating if not found (except for "fieldsets", "fields", and "statuslabels").
        
        Note wherever possible you should use the Get-SnipeIt<Entity>ByName functions for each supported entity type instead (and Get-SnipeItAssetEx or Sync-SnipeItAsset for assets).
        
        .PARAMETER Name
        Required. The name of the entity.

        .PARAMETER EntityType
        Required. The type of entity supported by the Snipe-It API. This is always in the form of their API name (IE, "departments").
        
        Currently supports: "departments", "locations", "companies", "manufacturers", "categories", "fieldsets", "fields", "statuslabels", "suppliers".

        .PARAMETER CreateParams
        The parameters to directly pass to the New-SnipeIt<Entity> function. These parameters are not checked for validity except for category_type with "categories", which is required for that entity type. 
        
        These parameters will also be used to further filter the results if more than one entity is returned.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it entity if not found. This is ignored for "fieldsets" and "statuslabels", which are never created.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the entity directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it entity object.
        
        .Notes
        Possible custom thrown exceptions: [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItEntityByName "IT" "departments"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$true,
                   Position=1)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","fieldsets","fields","statuslabels","suppliers")]
        [string]$EntityType,
        
        [parameter(Mandatory=$false)]
        [hashtable]$CreateParams = @{},
        
        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        $passParams = {
            OnErrorRetry = $OnErrorRetry
            SleepMS = $SleepMS
        }
    }
    Process {
        $name = $name.Trim()
        $update_cache = $false
        $sp_entities = $null
        $sp_entity = $null
        $cache_key = "sp_${EntityType}"
        $createFunc = $null
        if (-Not $DontCreateIfNotFound) {
            switch($EntityType) {
                "departments" {
                    $createFunc = "New-SnipeItDepartment"
                }
                "locations" {
                    $createFunc = "New-SnipeItLocation"
                }
                "companies" {
                    $createFunc = "New-SnipeItCompany"
                }
                "manufacturers" {
                    $createFunc = "New-SnipeItManufacturer"
                }
                "categories" {
                    $createFunc = "New-SnipeItCategory"
                    if ($CreateParams -eq $null -Or $CreateParams["category_type"] -notin @("asset", "accessory", "consumable", "component", "license")) {
                        Throw [System.Management.Automation.ValidationMetadataException] ("[Get-SnipeItEntityByName] Missing or invalid 'category_type' required parameter to CreateParams for New-SnipeItCategory: [{0}]" -f $CreateParams["category_type"])
                    }
                }
                "suppliers" {
                    $createFunc = "New-SnipeItSupplier"
                }
                "fieldsets" {
                    # Function does not support creating fieldsets
                }
                "fields" {
                    # Function does not support creating fields
                }
                "statuslabels" {
                    # Function does not support creating statuslabels
                }
                default {
                    # Should never get here
                    Throw [System.Management.Automation.ValidationMetadataException] "[Get-SnipeItEntityByName] Unsupported EntityType: $EntityType (should never get here?)"
                }
            }
        }
        if (-Not [string]::IsNullorEmpty($name)) {
            try {
                $sp_entities = Get-SnipeItEntityAll $EntityType
            } catch [System.Net.WebException] {
                # Try to continue if we can't get cache.
                Write-Error $_
            }
            if ($NoCache -Or $sp_entities.Count -eq 0) {
                $passParams = @{
                    OnErrorRetry=$OnErrorRetry
                    SleepMS=$SleepMS
                }
                # Check live data
                $sp_entity = Get-SnipeItApiEntityByName -Name $name -EntityType $EntityType @passParams
                if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
                    # Return early due to error
                    Throw ("[Get-SnipeItEntityByName] Got back fatal StatusCode [{0}]" -f $sp_entity.StatusCode)
                }
                if ($sp_entity.id -is [int]) {
                    $update_cache = $true
                }
            } else {
                $sp_entity = $sp_entities.Values | where {$_.Name -eq $name}
            }
            # If we have more than one result, and we have additional Create parameters, attempt to filter by those given parameters
            if ($sp_entity.Count -gt 1 -And $CreateParams -ne $null) {
                Write-Verbose ("[Get-SnipeItEntityByName] {0} results returned for [{1}] with name [{2}], attempting to filter by all given parameters..." -f $sp_entity.Count, $EntityType, $name)
                $sp_entity_tmp = Select-SnipeItFilteredEntity $sp_entity $CreateParams -All
                # Check to see if this results in a better match.
                if ($sp_entity_tmp.Count -lt $sp_entity.Count) {
                    if ($sp_entity_tmp.Count -gt 0) {
                        # Best match we're going to find.
                        $sp_entity = $sp_entity_tmp
                    } else {
                        Write-Verbose ("[Get-SnipeItEntityByName] No results returned for [{0}] with name [{1}] filtered by all parameters, attempting to filter by each parameter individually..." -f $EntityType, $name)
                        $sp_entity_tmp = Select-SnipeItFilteredEntity $sp_entity $CreateParams
                        # Check to see if this results in less matches.
                        if ($sp_entity_tmp.Count -lt $sp_entity.Count) {
                            # We found a better match (hopefully an exact one).
                            $sp_entity = $sp_entity_tmp
                        }
                    }
                }
            }
            # Check to see if we still have more than one result
            if ($sp_entity.Count -gt 1) {   
                if ($ErrorOnDupe) {
                    Throw [System.Data.DuplicateNameException] ("[Get-SnipeItEntityByName] {0} results returned for [{1}] with name [{2}] and -ErrorOnDupe is set" -f $sp_entity.Count, $EntityType, $name)
                } else {
                    Write-Warning ("[Get-SnipeItEntityByName] {0} results returned for [{1}] with name [{2}], using first result" -f $sp_entity.Count, $EntityType, $name)
                    $sp_entity = $sp_entity | Select -First 1
                }
            }
            # Create entity if it doesn't exist
            if ($createFunc -ne $null -And $sp_entity.id -isnot [int] -And -Not $DontCreateIfNotFound) {
                Write-Debug("Attempting to create new type of [$EntityType] with name [$name] and additional parameters (if any):" + ($CreateParams | ConvertTo-Json -Depth 10))
                $count_retry = $OnErrorRetry
                while ($count_retry -ge 0) {
                    # TODO: Suggest update to SnipeitPS suppress these warnings
                    $sp_entity = &$createFunc -name $name @CreateParams -WarningAction SilentlyContinue
                    if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                        $count_retry--
                        Write-Warning ("[Get-SnipeItEntityByName] ERROR creating new snipeit {0} with name [{1}]! StatusCode: {2}, StatusDescription: {3}, Retries Left: {4}" -f $EntityType,$Name,$sp_entity.StatusCode,$sp_entity.StatusDescription,$count_retry)
                    } else {
                        if ([string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.id -is [int]) {
                            Write-Verbose ("[Get-SnipeItEntityByName] Created new type of [{0}] in snipe-it with name [{1}] and id [{2}]" -f $EntityType,$sp_entity.name,$sp_entity.id)
                            $update_cache = $true
                        }
                        # Break out of loop early on anything except "Too Many Requests"
                        $count_retry = -1
                    }
                    # Sleep before next API call
                    Start-Sleep -Milliseconds $SleepMS
                }
                if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
                    Throw [System.Net.WebException] ("[Get-SnipeItEntityByName] Fatal ERROR creating new snipeit {0} with name [{1}]! StatusCode: {2}, StatusDescription: {3}" -f $EntityType,$Name,$sp_entity.StatusCode,$sp_entity.StatusDescription)
                }
            }
        }
        # Update cache, if initialized
        if ($update_cache) {
            $success = Update-SnipeItCache $sp_entity $EntityType
        }
        return $sp_entity
    }
    End {
    }
}

function Get-SnipeItCompanyByName {
    <#
        .SYNOPSIS
        Returns the given company by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given company by name, optionally creating if not found.
        
        .PARAMETER Name
        Required. The name of the company.

        .PARAMETER image
        Image file to be passed to New-SnipeItCompany if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it company if not found.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the company directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it company object.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItCompanyByName "Test Company"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound -And -Not [string]::IsNullOrEmpty($image)) {
            $passParams.Add("CreateParams", @{ "image" = $image } )
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "companies" @passParams
    }
    End {
    }
}

function Get-SnipeItDepartmentByName {
    <#
        .SYNOPSIS
        Returns the given department by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given department by name, optionally creating if not found.
        
        Will attempt to filter by all parameters given if more than one result is returned when searching by name.
        
        .PARAMETER Name
        Required. The name of the department.

        .PARAMETER company_id
        Parameter to be passed to New-SnipeItDepartment if created. Also used in filtering if name returns more than one result.

        .PARAMETER location_id
        Parameter to be passed to New-SnipeItDepartment if created. Also used in filtering if name returns more than one result.

        .PARAMETER manager_id
        Parameter to be passed to New-SnipeItDepartment if created. Also used in filtering if name returns more than one result.

        .PARAMETER notes
        Parameter to be passed to New-SnipeItDepartment if created. Also used in filtering if name returns more than one result.

        .PARAMETER image
        Parameter to be passed to New-SnipeItDepartment if created.

        .PARAMETER image_delete
        Parameter to be passed to New-SnipeItDepartment if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it department if not found.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the department directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it department object.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItDepartmentByName "IT" -company_id 1
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$company_id,

        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$location_id,

        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$manager_id,

        [parameter(Mandatory=$false)]
        [string]$notes,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$image_delete=$false,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to Get-SnipeItEntityByName
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound) {
            $createParams = @{}
            foreach ($param in @("company_id","location_id","manager_id","notes","image")) {
                $val = $PSBoundParameters[$param]
                if (-Not [string]::IsNullOrEmpty($val)) {
                    $createParams[$param] = $val
                }
            }
            if ($image_delete) {
                $createParams["image_delete"] = $true
            }
            if ($createParams.Count -gt 0) {
                $passParams.Add("CreateParams", $createParams)
            }
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "departments" @passParams
    }
    End {
    }
}

function Get-SnipeItLocationByName {
        <#
        .SYNOPSIS
        Returns the given location by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given location by name, optionally creating if not found. 
        
        Will attempt to filter by all parameters given if more than one result is returned when searching by name.
        
        .PARAMETER Name
        Required. The name of the location.

        .PARAMETER address
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER address2
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER city
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER state
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER country
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER zip
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER currency
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER parent_id
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER manager_id
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER ldap_ou
        Parameter to be passed to New-SnipeItLocation if created. Also used in filtering if name returns more than one result.

        .PARAMETER image
        Parameter to be passed to New-SnipeItLocation if created.

        .PARAMETER image_delete
        Parameter to be passed to New-SnipeItLocation if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it location if not found.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the location directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it location object.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItLocationByName "New Brunswick" -Country "CA"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$false)]
        [string]$address,

        [parameter(Mandatory=$false)]
        [string]$address2,

        [parameter(Mandatory=$false)]
        [string]$city,

        [parameter(Mandatory=$false)]
        [string]$state,

        [parameter(Mandatory=$false)]
        [string]$country,

        [parameter(Mandatory=$false)]
        [string]$zip,

        [parameter(Mandatory=$false)]
        [string]$currency,

        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$parent_id,

        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$manager_id,

        [parameter(Mandatory=$false)]
        [string]$ldap_ou,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$image_delete=$false,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to Get-SnipeItEntityByName
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound) {
            $createParams = @{}
            foreach ($param in @("address","address2","city","state","country","zip","currency","parent_id","manager_id","ldap_ou","image")) {
                $val = $PSBoundParameters[$param]
                if (-Not [string]::IsNullOrEmpty($val)) {
                    $createParams[$param] = $val
                }
            }
            if ($image_delete) {
                $createParams["image_delete"] = $true
            }
            if ($createParams.Count -gt 0) {
                $passParams.Add("CreateParams", $createParams)
            }
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "locations" @passParams
    }
    End {
    }
}

function Get-SnipeItCategoryByName {
    <#
        .SYNOPSIS
        Returns the given category by name, optionally creating if not found.
        
        Will attempt to filter by all parameters given if more than one result is returned when searching by name.
        
        .DESCRIPTION
        Returns the given category by name, optionally creating if not found.
        
        .PARAMETER Name
        Required. The name of the category.

        .PARAMETER category_type
        The type of category out of ("asset", "accessory", "consumable", "component", "license") to be passed to New-SnipeItCategory if created. Required unless -DontCreateIfNotFound is given. Also used in filtering if name returns more than one result.

        .PARAMETER eula_text
        Parameter to be passed to New-SnipeItCategory if created. Also used in filtering if name returns more than one result.

        .PARAMETER require_acceptance
        Parameter to be passed to New-SnipeItCategory if created. Also used in filtering if name returns more than one result.

        .PARAMETER checkin_email
        Parameter to be passed to New-SnipeItCategory if created. Also used in filtering if name returns more than one result.

        .PARAMETER image
        Parameter to be passed to New-SnipeItCategory if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it category if not found. 

        .PARAMETER NoCache
        Ignore the cache and try to fetch the category directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it category.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItCategoryByName "PC" -category_type "asset"
    #>
    [CmdletBinding(DefaultParameterSetName = 'CreateIfNotFound')]
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [parameter(Mandatory=$true, ParameterSetName='CreateIfNotFound')]
        [parameter(Mandatory=$false, ParameterSetName='DontCreateIfNotFound')]
        [ValidateSet("asset", "accessory", "consumable", "component", "license")]
        [string]$category_type,

        [parameter(Mandatory=$false)]
        [string]$eula_text,

        [parameter(Mandatory=$false)]
        [switch]$use_default_eula,

        [parameter(Mandatory=$false)]
        [switch]$require_acceptance,

        [parameter(Mandatory=$false)]
        [switch]$checkin_email,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$true, ParameterSetName='DontCreateIfNotFound')]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to Get-SnipeItEntityByName
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound) {
            $createParams = @{}
            foreach ($param in @("category_type","eula_text","image")) {
                $val = $PSBoundParameters[$param]
                if (-Not [string]::IsNullOrEmpty($val)) {
                    $createParams[$param] = $val
                }
            }
            foreach ($param in @("use_default_eula","require_acceptance","checkin_email")) {
                if ($PSBoundParameters[$param]) {
                    $createParams[$param] = $true
                }
            }
            if ($createParams.Count -gt 0) {
                $passParams.Add("CreateParams", $createParams)
            }
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "categories" @passParams
    }
    End {
    }
}

function Get-SnipeItManufacturerByName {
    <#
        .SYNOPSIS
        Returns the given manufacturer by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given manufacturer by name, optionally creating if not found.
        
        .PARAMETER Name
        Required. The name of the manufacturer.

        .PARAMETER image
        Parameter to be passed to New-SnipeItManufacturer if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it manufacturer if not found.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the manufacturer directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it manufacturer.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItManufacturerByName "PC" -manufacturer_type "asset"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,
        
        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound -And -Not [string]::IsNullOrEmpty($image)) {
            $passParams.Add("CreateParams", @{ "image" = $image } )
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "manufacturers" @passParams
    }
    End {
    }
}

function Get-SnipeItSupplierByName {
    <#
        .SYNOPSIS
        Returns the given supplier by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given supplier by name, optionally creating if not found.
        
        Will attempt to filter by all parameters given if more than one result is returned when searching by name.
        
        .PARAMETER Name
        Required. The name of the supplier.

        .PARAMETER address
        Address to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER address2
        Address2 to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER city
        City to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER state
        State to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER country
        Country to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER zip
        Zip to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER phone
        Phone number to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER fax
        Fax number to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER email
        Email address to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER contact
        Contact information to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.

        .PARAMETER notes
        Supplier notes to be passed to New-SnipeitSupplier if created. Also used in filtering if name returns more than one result.
        
        .PARAMETER image
        Image file to be passed to New-SnipeItsupplier if created.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it supplier if not found.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the supplier directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it supplier object.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItSupplierByName "Test supplier"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [parameter(Mandatory=$false)]
        [string]$address,

        [parameter(Mandatory=$false)]
        [string]$address2,

        [parameter(Mandatory=$false)]
        [string]$city,

        [parameter(Mandatory=$false)]
        [string]$state,

        [parameter(Mandatory=$false)]
        [string]$country,

        [parameter(Mandatory=$false)]
        [string]$zip,

        [parameter(Mandatory=$false)]
        [string]$phone,

        [parameter(Mandatory=$false)]
        [string]$fax,

        [parameter(Mandatory=$false)]
        [string]$email,

        [parameter(Mandatory=$false)]
        [string]$contact,

        [parameter(Mandatory=$false)]
        [string]$notes,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            DontCreateIfNotFound=$DontCreateIfNotFound
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        if (-Not $DontCreateIfNotFound) {
            parameter(Mandatory=$false)]
            $createParams = @{}
            foreach ($param in @("address","address2","city","state","country","zip","phone","fax","email","contact","notes")) {
                $val = $PSBoundParameters[$param]
                if (-Not [string]::IsNullOrEmpty($val)) {
                    $createParams[$param] = $val
                }
            }
            if ($createParams.Count -gt 0) {
                $passParams.Add("CreateParams", $createParams)
            }
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "companies" @passParams
    }
    End {
    }
}

function Get-SnipeItFieldsetByName {
    <#
        .SYNOPSIS
        Returns the given fieldset by name.
        
        .DESCRIPTION
        Returns the given fieldset by name. This function does NOT create new fieldsets.
        
        .PARAMETER Name
        Required. The name of the fieldset.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the fieldset directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it fieldset.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItFieldsetByName "PC"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "fieldsets" @passParams
    }
    End {
    }
}

function Get-SnipeItCustomFieldByName {
    <#
        .SYNOPSIS
        Returns the given custom field by name.
        
        .DESCRIPTION
        Returns the given custom field by name. This function does NOT create new custom fields.
        
        .PARAMETER Name
        Required. The name of the custom field.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the custom field directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it custom field object.
        
        .Notes
        You can also get a map of custom fields by calling Get-SnipeItCustomFieldMap.
        
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItCustomFieldByName "PC"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "fields" @passParams
    }
    End {
    }
}

function Get-SnipeItStatuslabelByName {
    <#
        .SYNOPSIS
        Returns the given statuslabel by name.
        
        .DESCRIPTION
        Returns the given statuslabel by name. This function does NOT create new status labels.
        
        .PARAMETER Name
        Required. The name of the statuslabel.

        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.

        .PARAMETER NoCache
        Ignore the cache and try to fetch the statuslabel directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it statuslabel.
        
        .Notes
        Possible custom thrown exceptions (from Get-SnipeItEntityByName): [System.Management.Automation.ValidationMetadataException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItStatuslabelByName "Pending"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # parameters to pass to 
        $passParams = @{
            ErrorOnDupe=$ErrorOnDupe
            NoCache=$NoCache
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
    }
    Process {
        return Get-SnipeItEntityByName $Name -EntityType "statuslabels" @passParams
    }
    End {
    }
}

function Get-SnipeItModelByName {
    <#
        .SYNOPSIS
        Returns the given model by name, optionally creating if not found.
        
        .DESCRIPTION
        Returns the given model by name, optionally creating if not found.
        
        .PARAMETER Name
        Required. The name of the model.

        .PARAMETER Manufacturer
        Either the Manufacturer ID or the Manufacturer Name. If given a name that does not exist, create it unless -DontCreateManufacturerIfNotFound is given. Required unless -DontCreateIfNotFound is given.

        .PARAMETER Category
        Either the Category ID or the Category Name. If given a name that does not exist, create it unless -DontCreateCategoryIfNotFound is given. Required unless -DontCreateIfNotFound is given.
        
        .PARAMETER Fieldset
        Either the Fieldset ID or the Fieldset Name to associate with a newly created model.
        
        .PARAMETER model_number
        The model number to associate with the newly created model.
        
        .PARAMETER image
        Parameter to be passed to New-SnipeItModel if created.
        
        .PARAMETER ErrorOnDupe
        Throw [System.Data.DuplicateNameException] on finding a duplicate name instead of giving a warning.
        
        .PARAMETER DontCreateIfNotFound
        Don't create a new snipe-it model if not found. If this switch is not given then -Manufacturer and -Category are required.
        
        .PARAMETER DontCreateManufacturerIfNotFound
        Don't create a new snipe-it manufacturer if not found.
        
        .PARAMETER DontCreateCategoryIfNotFound
        Don't create a new snipe-it category if not found.
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch everything directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it model.
        
        .Notes
        If a new model is to be created, and the manufacturer and category cannot be found or created, [System.Data.ObjectNotFoundException] is thrown.
        Possible thrown custom exceptions: [System.Data.ObjectNotFoundException], [System.Data.DuplicateNameException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItModelByName "Latitude 7490" -Manufacturer "Dell, Inc." -Category "PC" -Fieldset "PC"
    #>
    [CmdletBinding(DefaultParameterSetName = 'CreateIfNotFound')]
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Name,
        
        [parameter(Mandatory=$true, ParameterSetName='CreateIfNotFound')]
        [parameter(Mandatory=$false, ParameterSetName='DontCreateIfNotFound')]
        [string]$Manufacturer,
        
        [parameter(Mandatory=$true, ParameterSetName='CreateIfNotFound')]
        [parameter(Mandatory=$false, ParameterSetName='DontCreateIfNotFound')]
        [string]$Category,
        
        [parameter(Mandatory=$false)]
        [string]$Fieldset,
        
        [parameter(Mandatory=$false)]
        [string]$model_number,

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})]
        [string]$image,

        [parameter(Mandatory=$false)]
        [switch]$ErrorOnDupe,
        
        [parameter(Mandatory=$true, ParameterSetName='DontCreateIfNotFound')]
        [switch]$DontCreateIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateManufacturerIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$DontCreateCategoryIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
    }
    Process {
        $name = $name.Trim()
        $update_cache = $false
        $sp_entities = $null
        $sp_entity = $null
        $cache_key = "sp_models"
        $EntityType = "models"
        if (-Not [string]::IsNullorEmpty($name)) {
            # Pass parameters to other ByName function calls
            $passParams = @{
                OnErrorRetry=$OnErrorRetry
                SleepMS=$SleepMS
            }
            foreach($param in @("Verbose","Debug")) {
                if ($PSBoundParameters[$param]) {
                    $passParams.Add($param, $PSBoundParameters[$param])
                }
            }
            try {
                $sp_entities = Get-SnipeItEntityAll $EntityType
            } catch [System.Net.WebException] {
                # Try to continue if we can't get cache.
                Write-Error $_
            }
            if ($NoCache -Or $sp_entities.Count -eq 0) {
                # Check live data
                $sp_entity = Get-SnipeItApiEntityByName -Name $name -EntityType $EntityType @passParams
                if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
                    Throw [System.Net.WebException] ("[Get-SnipeItModelByName] Got back fatal StatusCode [{0}]" -f $sp_entity.StatusCode)
                }
                if ($sp_entity.id -is [int]) {
                    $update_cache = $true
                }
            } else {
                # Check cache
                $sp_entity = $sp_entities.Values | where {$_.Name -is [string] -And [System.Net.WebUtility]::HtmlDecode($_.Name).Trim() -eq $name}
            }
            # If we have more than one result, and we have additional Create parameters, attempt to filter by those given parameters
            if ($sp_entity.Count -gt 1 -And (-Not [string]::IsNullOrEmpty($Manufacturer) -Or -Not [string]::IsNullOrEmpty($Category) -Or -Not [string]::IsNullOrEmpty($Fieldset) -Or -Not [string]::IsNullOrEmpty($model_number))) {
                Write-Verbose ("[Get-SnipeItModelByName] {0} results returned for [{1}] with name [{2}], attempting to filter by all given parameters..." -f $sp_entity.Count, $EntityType, $name)
                $params = @{}
                if (-Not [string]::IsNullOrEmpty($Manufacturer)) {
                    $id = $Manufacturer -as [int]
                    if ($id -is [int]) {
                        $params['manufacturer_id'] = $id
                    } else {
                        $params['manufacturer_<name>'] = $Manufacturer
                    }
                }
                if (-Not [string]::IsNullOrEmpty($Category)) {
                    $id = $Category -as [int]
                    if ($id -is [int]) {
                        $params['category_id'] = $id
                    } else {
                        $params['category_<name>'] = $Category
                    }
                }
                if (-Not [string]::IsNullOrEmpty($Fieldset)) {
                    $id = $Fieldset -as [int]
                    if ($id -is [int]) {
                        $params['fieldset_id'] = $id
                    } else {
                        $params['fieldset_<name>'] = $Fieldset
                    }
                }
                if (-Not [string]::IsNullOrEmpty($model_number)) {
                    $params['model_number'] = $model_number
                }
                $sp_entity_tmp = Select-SnipeItFilteredEntity $sp_entity $params -All
                # Check to see if this results in a better match.
                if ($sp_entity_tmp.Count -lt $sp_entity.Count) {
                    if ($sp_entity_tmp.Count -gt 0) {
                        # Best match we're going to find.
                        $sp_entity = $sp_entity_tmp
                    } else {
                        Write-Verbose ("[Get-SnipeItEntityByName] No results returned for [{0}] with name [{1}] filtered by all parameters, attempting to filter by each parameter individually..." -f $EntityType, $name)
                        $sp_entity_tmp = Select-SnipeItFilteredEntity $sp_entity $params
                        # Check to see if this results in less matches.
                        if ($sp_entity_tmp.Count -lt $sp_entity.Count) {
                            # We found a better match (hopefully an exact one).
                            $sp_entity = $sp_entity_tmp
                        }
                    }
                }
            }
            # Check to see if we still have more than 1 result.
            if ($sp_entity.id.Count -gt 1) {
                if ($ErrorOnDupe) {
                    Throw [System.Data.DuplicateNameException] ("[Get-SnipeItModelByName] {0} results returned for [{1}] with name [{2}] and -ErrorOnDupe is set" -f $sp_entity.id.Count, $EntityType, $name)
                } else {
                    Write-Warning ("[Get-SnipeItModelByName] {0} results returned for [{1}] with name [{2}], using first result" -f $sp_entity.id.Count, $EntityType, $name)
                    $sp_entity = $sp_entity | Select -First 1
                }
            }
            if ($sp_entity.id -isnot [int] -And -Not $DontCreateIfNotFound) {
                # Additional parameters for other ByName function calls
                foreach($param in @("NoCache","ErrorOnDupe")) {
                    if ($PSBoundParameters[$param]) {
                        $passParams[$param] = $PSBoundParameters[$param]
                    }
                }
                # Required: Get Manufacturer and Category by Name unless a valid ID is given
                $sp_manufacturer_id = $Manufacturer -as [int]
                if ($sp_manufacturer_id -isnot [int]) {
                    if ($DontCreateManufacturerIfNotFound) {
                        $pp = $passParams.Clone()
                        $pp.Add("DontCreateIfNotFound", $true)
                    } else {
                        $pp = $passParams
                    }
                    # Will throw [System.Net.WebException] if there's a problem with the request
                    $sp_manufacturer_id = (Get-SnipeItManufacturerByName $Manufacturer @pp).id
                    if ($sp_manufacturer_id -isnot [int]) {
                        Throw [System.Data.ObjectNotFoundException] "[Get-SnipeItModelByName] Got back invalid Manufacturer ID for name [${Manufacturer}]"
                    }
                }
                $sp_category_id = $Category -as [int]
                if ($sp_category_id -isnot [int]) {
                    $pp = $passParams.Clone()
                    if ($DontCreateCategoryIfNotFound) {
                        $pp.Add("DontCreateIfNotFound", $true)
                    } else {
                        $pp.Add("category_type", "asset")
                    }
                    # Will throw [System.Net.WebException] if there's a problem with the request
                    $sp_category_id = (Get-SnipeItCategoryByName $Category @pp).id
                    if ($sp_category_id -isnot [int]) {
                        Throw [System.Data.ObjectNotFoundException] "[Get-SnipeItModelByName] Got back invalid Category ID for name [${Category}]"
                    }
                }
                # Optional parameters
                $createParams = @{}
                # Set Fieldset on creation
                if (-Not [string]::IsNullOrEmpty($Fieldset)) {
                    $sp_fieldset_id = $Fieldset -as [int]
                    if ($sp_fieldset_id -isnot [int]) {
                        try {
                            $sp_fieldset_id = (Get-SnipeItFieldsetByName $Fieldset).id
                        } catch [System.Net.WebException] {
                            # Try to continue on error.
                            Write-Error $_
                        }
                        if ($sp_fieldset_id -isnot [int]) {
                            Write-Warning "[Get-SnipeItModelByName] Fieldset not found [${Fieldset}]"
                        }
                    }
                    if ($sp_fieldset_id -is [int]) {
                        $createParams.Add("fieldset_id", $sp_fieldset_id)
                    }
                }
                if (-Not [string]::IsNullOrEmpty($model_number)) {
                    $createParams.Add("model_number", $model_number)
                }
                if (-Not [string]::IsNullOrEmpty($image)) {
                    $createParams.Add("image", $image)
                }
                Write-Debug("[Get-SnipeItModelByName] Attempting to create new type of [$EntityType] with name [$name] and additional parameters (if any):" + ($createParams | ConvertTo-Json -Depth 10))
                $count_retry = $OnErrorRetry
                while ($count_retry -ge 0) {
                    $sp_entity = New-SnipeitModel -name $name -category_id $sp_category_id -manufacturer_id $sp_manufacturer_id @createParams
                    if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                        $count_retry--
                        Write-Warning ("[Get-SnipeItModelByName] ERROR creating new snipeit {0} by name [{1}]! StatusCode: {2}, StatusDescription: {3}, Retries Left: {4}" -f $EntityType,$Name,$sp_entity.StatusCode,$sp_entity.StatusDescription,$count_retry)
                    } else {
                        if ([string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.id -is [int]) {
                            Write-Verbose ("[Get-SnipeItModelByName] Created new snipeit {0} [{1}] with id [{2}]" -f $EntityType,$sp_entity.name,$sp_entity.id)
                            $update_cache = $true
                        }
                        # Break out of loop early on anything except "Too Many Requests"
                        $count_retry = -1
                    }
                    # Sleep before next API call
                    Start-Sleep -Milliseconds $SleepMS
                }
                if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
                    Throw [System.Net.WebException] ("[Get-SnipeItModelByName] Fatal ERROR creating new snipeit {0} with name [{1}]! StatusCode: {2}, StatusDescription: {3}" -f $EntityType,$Name,$sp_entity.StatusCode,$sp_entity.StatusDescription)
                }
            }
        }
        # Update cache, if initialized
        if ($update_cache) {
            $success = Update-SnipeItCache $sp_entity $EntityType
        }
        return $sp_entity
    }
    End {
    }
}

function Get-SnipeItEntityByID {
    <#
        .SYNOPSIS
        Returns the given snipe-it entity by ID.
        
        .DESCRIPTION
        Returns the given snipe-it entity by ID, returning from cache if it exists.
        
        .PARAMETER ID
        Required. The ID of the given entity. 

        .PARAMETER EntityType
        Required. The type of entity supported by the Snipe-It API. This is always in the form of their API name (IE, "departments"), except for "assets". 
        
        Supports: "departments", "locations", "companies", "manufacturers", "categories", "fieldsets", "fields", "statuslabels", "suppliers", "models", "assets", "users"
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch everything directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned snipe-it entity.
        
        .Notes
        Possible custom thrown exceptions: [System.Management.Automation.ValidationMetadataException], [System.Net.WebException]

        .Example
        PS> Get-SnipeItEntityByID 13 "users"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$ID,
        
        [parameter(Mandatory=$true,
                   Position=1)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","fieldsets","fields","statuslabels","suppliers","models","assets","users")]
        [string]$EntityType,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
    }
    Process {
        $update_cache = $false
        $sp_entities = $null
        $sp_entity = $null
        $cache_key = "sp_${EntityType}"
        switch($EntityType) {
            "departments" {
                $func = "Get-SnipeItDepartment"
            }
            "locations" {
                $func = "Get-SnipeItLocation"
            }
            "companies" {
                $func = "Get-SnipeItCompany"
            }
            "manufacturers" {
                $func = "Get-SnipeItManufacturer"
            }
            "categories" {
                $func = "Get-SnipeItCategory"
            }
            "fieldsets" {
                $func = "Get-SnipeItFieldset"
            }
            "fields" {
                $func = "Get-SnipeItCustomField"
            }
            "statuslabels" {
                $func = "Get-SnipeItStatus"
            }
            "suppliers" {
                $func = "Get-SnipeItSupplier"
            }
            "models" {
                $func = "Get-SnipeItModel"
            }
            "assets" {
                $func = "Get-SnipeItAsset"
            }
            "users" {
                $func = "Get-SnipeItUser"
            }
            default {
                # Should never get here
                Throw [System.Management.Automation.ValidationMetadataException] "[Get-SnipeItEntityByName] Unsupported EntityType: $EntityType (should never get here?)"
            }
        }
        try {
            $sp_entities = Get-SnipeItEntityAll $EntityType
        } catch [System.Net.WebException] {
            # Try to continue without cache.
            Write-Error $_
        }
        if ($NoCache -Or $sp_entities.Count -eq 0) {
            # Check live data
            $count_retry = $OnErrorRetry
            while ($count_retry -ge 0) {
                # TODO: Suggest update to SnipeItPS to suppress these warnings
                $sp_entity = &$func -id $ID -WarningAction SilentlyContinue
                if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                    $count_retry--
                    Write-Warning ("[Get-SnipeItEntityByName] ERROR getting snipeit {0} by ID [{1}]! StatusCode: {2}, StatusDescription: {3}, Retries Left: {4}" -f $EntityType,$ID,$sp_entity.StatusCode,$sp_entity.StatusDescription,$count_retry)
                } else {
                    if ([string]::IsNullOrWhitespace($sp_entity.StatusCode) -And $sp_entity.id -is [int]) {
                        $update_cache = $true
                    }
                    # Break out of loop early on anything except "Too Many Requests"
                    $count_retry = -1
                }
                # Sleep before next API call
                Start-Sleep -Milliseconds $SleepMS
            }
            if (-Not [string]::IsNullOrWhitespace($sp_entity.StatusCode)) {
                Throw [System.Net.WebException] ("[Get-SnipeItEntityByName] Fatal ERROR getting snipeit {0} with ID [{1}]! StatusCode: {2}, StatusDescription: {3}" -f $EntityType,$ID,$sp_entity.StatusCode,$sp_entity.StatusDescription)
            }
            if ($sp_entity.id -is [int]) {
                $update_cache = $true
            }
        } else {
            $sp_entity = $sp_entities[[string]$ID]
        }
        if ($sp_entity.id.Count -gt 1) {
            # Should never get here
            Write-Warning ("[Get-SnipeItEntityByID] {0} results returned for {1} [{2}], using first result (not sure how this happened?)" -f $sp_entity.id.Count, $EntityType, $ID)
            $sp_entity = $sp_entity | Select -First 1
        }
        # Update cache, if initialized
        if ($update_cache) {
            $success = Update-SnipeItCache $sp_entity $EntityType
        }
        return $sp_entity
    }
    End {
    }
}

function Get-SnipeItAssetEx {
    <#
        .SYNOPSIS
        Returns the given snipe-it asset using the given search criteria.
        
        .DESCRIPTION
        Returns the given snipe-it asset using the given search criteria, returning from cache if it exists.
        
        .PARAMETER AssetTag
        An AssetTag to search for. Either this parameter, Serial, Name, or CustomFieldValue must be given.

        .PARAMETER Serial
        A Serial Number to search for. Either this parameter, AssetTag, Name, or CustomFieldValue must be given.

        .PARAMETER Name
        An Asset Name to search for. Either this parameter, AssetTag, Serial, or CustomFieldValue must be given.
        
        .PARAMETER CustomFieldName
        The display name of the custom field to search for. Must be given if specifying a CustomFieldValue.
        
        .PARAMETER CustomDBFieldName
        The internal database name of the custom field to search for. Must be given if specifying a CustomFieldValue.
        
        .PARAMETER CustomFieldValue
        The value of the custom field to search for. Must be given if specifying a CustomField.
        
        .PARAMETER DontUpdateCacheIfNotFound
        Don't update the cache if the asset is found. Used internally pending updating an asset.
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch the asset directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned asset from snipe-it.
        
        .Notes
        Possible custom thrown exceptions: [System.Net.WebException]
        
        All Names, Custom Field Names, and string Custom Field Values are trimmed and run through HTMLDecode. Asset Tags and Serial # are only trimmed when searching from cache.
        
        Note this function does not check if more than one result / any duplicates are returned.

        .Example
        PS> Get-SnipeItAssetEx -AssetTag "1234"
        
        PS> Get-SnipeItAssetEx -Serial "56789ABC"
        
        PS> Get-SnipeItAssetEx -CustomFieldName "UUID" -CustomDBFieldName "_snipeit_uuid_2" -CustomFieldValue "25b2adae-20b7-11ed-861d-0242ac120002"
    #>
    [CmdletBinding(DefaultParameterSetName='Get by AssetTag')]
    param ( 
        [parameter(Mandatory=$true, ParameterSetName='Get by AssetTag')]
        [string]$AssetTag,
        
        [parameter(Mandatory=$true, ParameterSetName='Get by Serial')]
        [string]$Serial,

        [parameter(Mandatory=$true, ParameterSetName='Get by Asset Name')]
        [string]$Name,
        
        [parameter(Mandatory=$true, ParameterSetName='Get by CustomField')]
        [string]$CustomFieldName,
        
        [parameter(Mandatory=$true, ParameterSetName='Get by CustomField')]
        [string]$CustomDBFieldName,
        
        [parameter(Mandatory=$true, ParameterSetName='Get by CustomField')]
        [string]$CustomFieldValue,
        
        [parameter(Mandatory=$false)]
        [switch]$DontUpdateCacheIfFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )

    try {
        $sp_assets = Get-SnipeItEntityAll "assets"
    } catch [System.Net.WebException] {
        # Try to continue without cache.
        Write-Error $_
    }

    $field = $null
    $customvalue = $CustomFieldValue.Trim()
    if (-Not [string]::IsNullOrEmpty($AssetTag)) {
        $field = "asset_tag"
        $_assettag = $AssetTag
    } elseif (-Not [string]::IsNullOrEmpty($Serial)) {
        $field = "serial"
        $_serial = $Serial
    } elseif (-Not [string]::IsNullOrEmpty($Name)) {
        $field = "name"
        $_name = $Name.Trim()
    } else {
        $field = $CustomFieldName
        $_customvalue = $CustomFieldValue.Trim()
    }
    if ($NoCache -Or $sp_assets.Count -eq 0) {
        # Check live data
        $count_retry = $OnErrorRetry
        while ($count_retry -ge 0) {
            switch ($field) {
                "asset_tag" {
                    $sp_asset = Get-SnipeItAsset -asset_tag $_assettag
                }
                "serial" {
                    $sp_asset = Get-SnipeItAsset -serial $_serial
                }
                "name" {
                    $sp_asset = Get-SnipeItApiEntityByName -Name $_name -EntityType "hardware"
                    if ($sp_asset.id.Count -gt 1) {
                        # Don't think this is needed, just in case...
                        Write-Warning ("[Get-SnipeItAssetEx] Name [$_name] returned {0} results, filtering again based on name.." -f $sp_asset.Count)
                        $sp_asset = $sp_asset | where {$_.Name -is [string] -And [System.Net.WebUtility]::HtmlDecode($_.Name).Trim() -eq $_name}
                    }
                }
                default {
                    # TODO: Suggest update to SnipeItPS to include custom field filtering
                    $params = @{ $CustomDBFieldName = $_customvalue }
                    $sp_asset = & $_SNIPEITPSFEATURES { param($params) Invoke-SnipeitMethod -Api "/api/v1/hardware" -Method "GET" -GetParameters $params } $params
                }
            }
            if ([string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {                  
                $count_retry--
                Write-Warning ("[Get-SnipeItAssetEx] ERROR getting snipeit asset by field [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $field,$sp_asset.StatusCode,$sp_asset.StatusDescription,$count_retry)
            } else {
                if ([string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.id -is [int]) {
                    $update_cache = $true
                }
                # Break out of loop early on anything except "Too Many Requests"
                $count_retry = -1
            }
            # Sleep before next API call
            Start-Sleep -Milliseconds $SleepMS
        }
        if (-Not [string]::IsNullOrWhitespace($sp_asset.StatusCode)) {
            Throw [System.Net.WebException] ("[Get-SnipeItAssetEx] Fatal ERROR getting snipeit asset by field [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $field,$sp_asset.StatusCode,$sp_asset.StatusDescription)
        }
        if ($sp_asset.id -is [int]) {
            $update_cache = $true
        }
    } elseif ($field -eq "asset_tag") {
        $_assettag = $_assettag.Trim()
        $sp_asset = $sp_assets.Values | where {$_.asset_tag -is [string] -And $_.asset_tag.Trim() -eq $_assettag}
    } elseif ($field -eq "serial") {
        $_serial = $_serial.Trim()
        $sp_asset = $sp_assets.Values | where {$_.serial -is [string] -And $_.serial.Trim() -eq $_serial}
    } elseif ($field -eq "name") {
        $sp_asset = $sp_assets.Values | where {$_.Name -is [string] -And [System.Net.WebUtility]::HtmlDecode($_.Name).Trim() -eq $_name}
    } else {
        # Assets may be returned in either form of all custom fields into an array called custom_fields indexed by field name OR placed into the asset itself as DB field names.
        # We use Restore-SnipeItAssetCustomFields when updating cache to make sure it's consistent
        $sp_asset = $sp_assets.Values | where {($_.custom_fields.$CustomFieldName.value -is [string] -And [System.Net.WebUtility]::HtmlDecode($_.custom_fields.$CustomFieldName.value).Trim() -eq $_customvalue) -Or ($_.custom_fields.$CustomFieldName.value -isnot [string] -And $_.custom_fields.$CustomFieldName.value -eq $_customvalue)}
    }
    
    # Update cache if valid result
    if ($update_cache -And -Not $DontUpdateCacheIfFound) {
        $success = Update-SnipeItCache $sp_asset "assets"
    }

    return $sp_asset
}

function Get-SnipeItUserEx {
    <#
        .SYNOPSIS
        Returns the given snipe-it user using the given search criteria.
        
        .DESCRIPTION
        Returns the given snipe-it user using the given search criteria, returning from cache if it exists.
        
        .PARAMETER Username
        A username to search for. Either this parameter or EmployeeNum must be given.

        .PARAMETER EmployeeNum
        An employee number to search for. Either this parameter or Username must be given.
        
        .PARAMETER DontUpdateCacheIfNotFound
        Don't update the cache if the asset is found. Used internally pending updating a user.
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch the user directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The returned user from snipe-it.
        
        .Notes
        Possible custom thrown exceptions: [System.Net.WebException]
        
        Note this function does not check if more than one result / any duplicates are returned.

        .Example
        PS> Get-SnipeItUserEx -Username "bob1"
        
        PS> Get-SnipeItUserEx -EmployeeNum "S-1-5-21-1085031214-1563985344-725345543"
    #>
    [CmdletBinding(DefaultParameterSetName='Get by Username')]
    param ( 
        [parameter(Mandatory=$true,ParameterSetName="Get by Username")]
        [string]$Username,
        
        [parameter(Mandatory=$true, ParameterSetName="Get by EmployeeNum")]
        [string]$EmployeeNum,
        
        [parameter(Mandatory=$false)]
        [switch]$DontUpdateCacheIfFound,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    # Get reference to cache, if it exists
    try {
        if ($EmployeeNum) {
            $sp_users = Get-SnipeItEntityAll "users" -UsersKey "employee_num"
        } else {
            $sp_users = Get-SnipeItEntityAll "users" -UsersKey "username"
        }
    } catch [System.Net.WebException] {
        # Try to continue without cache.
        Write-Error $_
    }
    $sp_user = $null
    $check_username = (-Not [string]::IsNullOrEmpty($Username))
    $check_employeenum = (-Not [string]::IsNullOrEmpty($EmployeeNum))
    if ($NoCache -Or $sp_users.Count -eq 0) {
        # Check live data
        $count_retry = $OnErrorRetry
        while ($count_retry -ge 0) {
            if ($check_employeenum) {
                if ($check_username) {
                    $matchval = $Username
                    # TODO: Suggest update to SnipeItPS to include employee_num filtering
                    $sp_user = & $_SNIPEITPSFEATURES { param($employee_num,$username) Invoke-SnipeitMethod -Api "/api/v1/users" -Method "GET" -GetParameters @{ employee_num=$employee_num; username=$username } } $EmployeeNum $Username
                } else {
                    $matchval = $EmployeeNum
                    # TODO: Suggest update to SnipeItPS to include employee_num filtering
                    $sp_user = & $_SNIPEITPSFEATURES { param($employee_num) Invoke-SnipeitMethod -Api "/api/v1/users" -Method "GET" -GetParameters @{ employee_num=$employee_num } } $EmployeeNum
                }
            } else {
                $matchval = $Username
                $sp_user = Get-SnipeItUser -Username $Username
            }
            if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                $count_retry--
                Write-Warning ("[Get-SnipeItUserEx] ERROR getting snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $matchval,$sp_user.StatusCode,$sp_user.StatusDescription,$count_retry)
            } else {
                if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.id -is [int]) {
                    $update_cache = $true
                }
                # Break out of loop early on anything except "Too Many Requests"
                $count_retry = -1
            }
            # Sleep before next API call
            Start-Sleep -Milliseconds $SleepMS
        }
        # All attempts failed
        if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode)) {
            Throw [System.Net.WebException] ("[Get-SnipeItUserEx] Fatal ERROR getting snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $matchval,$sp_user.StatusCode,$sp_user.StatusDescription)
        }
    } else {
        $sp_user = $sp_users.Values | where {($check_username -And $_.username -eq $Username) -Or ($check_employeenum -And $_.employee_num -eq $EmployeeNum)}
    }
    
    # Update cache, if initialized
    if ($update_cache -And -Not $DontUpdateCacheIfFound) {
        $success = Update-SnipeitCache $sp_user "users"
    }
    return $sp_user
}
            
function Sync-SnipeItUser {
    <#
        .SYNOPSIS
        Syncs the given user with snipe-it.
        
        .DESCRIPTION
        Syncs the given user with snipe-it. The given user object must have valid fields for New-SnipeItUser.
        
        .PARAMETER User
        The user object to sync. "first_name", "last_name", and "username" are required (also "employee_num" if -SyncOnEmployeeNum is given).
        
        Default fields which can be synced: "first_name", "last_name", "username", "password", "employee_num", "email", "phone", "notes", "jobtitle", "activated", "ldap_import", "image", "manager_id", "Manager", "company_id", "Company", "department_id", "Department", "location_id", "Location".
        
        These fields will create a new entry in snipe-it if they don't exist: "company", "department", and "location". See Notes for additional fields which can be given when creating these snipe-it entities.

        .PARAMETER RequiredSyncFields
        Only sync if all the given fields are non-blank. Defaults to requiring a valid department_id. Note that "username", "first_name", and "last_name" must always be non-blank, as well as "employee_num" if -SyncOnEmployeeNum is given.

        .PARAMETER SyncOnEmployeeNum
        Search for users in snipe-it by their employee_num. This is intended to be used with the SID value from AD, but can be used with any unique value. The employee_num field must be non-blank for this to work. If the employee_num field is blank in snipe-it, double-check the username for a match. If we return more than one user this way, throw [System.Data.DuplicateNameException].
        
        .PARAMETER OnlyUpdateBlankFields
        Only update the given fields if they are blank in snipe-it. Note that "employee_num" is invalid for this option if -SyncOnEmployeeNum is given.
        
        .PARAMETER DefaultNotes
        Default notes to write on new user creation or update if blank, only if the "notes" field is blank for the given user. Default: "User synced by API Script".
        
        .PARAMETER OverwriteNotes
        Always overwrite the notes field if matching the given notes. Defaults to "<DELETED_FROM_AD>", in case a user was given this note from Remove-SnipeItInactiveUsers.
        
        .PARAMETER DontCreateIfNotFound
        Don't create the user if not found, only update existing users.
        
        .PARAMETER DontCreateCompanyIfNotFound
        Don't create any new companies if given but not found.
        
        .PARAMETER DontCreateDepartmentIfNotFound
        Don't create any new departments if given but not found.
        
        .PARAMETER DontCreateLocationIfNotFound
        Don't create any new locations if given but not found.
        
        .PARAMETER DebugOutputCreateOnly
        Only output debug information when creating new users.

        .PARAMETER Trim
        Trims all strings instead of just proper names and username (and employee_num if -SyncOnEmployeeNum is set).
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch the user directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        Returns the processed asset.
        
        .Notes
        When creating companies, departments, and locations it will also sync these fields: "company_image", "department_manager_id", "department_location_id", "department_image", "department_notes", "location_address", "location_address2", "location_city", "location_country", "location_currency", "location_state", "location_zip", "location_image", "location_ldap_ou", "location_manager_id", "location_parent_id". If more than one result is given by name it will also use these fields to filter the results (except for "image").
        
        Managers are synced recursively to make sure they exist first. Although the function double-checks if a user is a manager to themselves, if you have a circle in your manager assignments this will result in infinite recursion and a likely crash.
        
        If a duplicate if found from either username (should never happen) or employee_num, [System.Data.DuplicateNameException] is thrown.
        
        Possible custom thrown exceptions: [System.Management.Automation.ValidationMetadataException], [System.Net.WebException], [System.Data.DuplicateNameException]
        
        There are a lot of different possible ways exceptions can be thrown. If you're not using a try/catch block you may need to use -ErrorAction Continue.
        
        .Example
        PS> Sync-SnipeItUser -User $User -SyncOnEmployeeNyum
        
        PS> $formatted_users | Sync-SnipeItUser -SyncOnEmployeeNyum
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({-Not [string]::IsNullOrWhitespace($_.username) -And -Not [string]::IsNullOrEmpty($_.first_name) -And -Not [string]::IsNullOrEmpty($_.last_name)})]
        [object]$User,
        
        # Do not create/update user when select properties are blank.
        # Username must always be non-blank (and employee_num must be non-blank IF SyncOnEmployeeNum is set).
        # first_name and last_name are also always required.
        [parameter(Mandatory=$false)]
        [ValidateSet("employee_num", "email", "phone", "notes", "jobtitle", "activated", "ldap_import", "image", "company_id", "department_id", "location_id", "manager_id")]
        [AllowEmptyCollection()]
        [string[]]$RequiredSyncFields=@("department_id"),
        
        # Assumes EmployeeNum is unique.
        [parameter(Mandatory=$false)]
        [switch]$SyncOnEmployeeNum,
        
        # Dont overwrite existing values when updating
        # "employee_num" must always be updated if $SyncOnEmployeeNum is set
        [parameter(Mandatory=$false)]
        [ValidateSet("employee_num", "email", "phone", "notes", "jobtitle", "image", "groups", "company_id", "department_id", "location_id", "manager_id")]
        [ValidateScript({-Not $SyncOnEmployeeNum -Or $_ -notcontains "employee_num"})]
        [AllowEmptyCollection()]
        [string[]]$OnlyUpdateBlankFields,
        
        [parameter(Mandatory=$false)]
        [string]$DefaultNotes = "User synced by API Script",
        
        # Ignore "OnlyUpdateBlankFields" for Notes if matching.
        # This is used with Remove-SnipeItInactiveUsers for flagging users deleted from AD.
        [parameter(Mandatory=$false)]
        [string]$OverwriteNotes = "<DELETED_FROM_AD>",
        
        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateCompanyIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateDepartmentIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateLocationIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$Trim,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # Fields accepted by the Add and Set functions that do not fall under $_SPECIALFIELDS.
        $_BUILTINFIELDS = @("first_name", "last_name", "username", "password", "employee_num", "email", "phone", "notes", "jobtitle", "activated", "ldap_import", "image", "groups")
        # Give an error if these fields are missing.
        $_ALWAYSREQUIREDFIELDS = @("first_name", "last_name", "username")
        # Don't give an error if these fields are missing.
        $_NONREQUIREDFIELDS = @("password", "email", "phone", "notes", "jobtitle", "activated", "ldap_import", "image", "groups")

        # Create a hash table for faster lookups.
        $_onlyUpdateBlankFieldsMap = @{}
        foreach ($field in $OnlyUpdateBlankFields) {
            $_onlyUpdateBlankFieldsMap[$field] = $true
        }
    }
    Process {
        $update_cache = $false
        $sp_user = $null
        $sp_users = $null
        $user_values = @{}
        
        # Additional parameters for other entity function calls
        $passParams = @{ 
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        foreach($param in @("NoCache","Verbose")) {
            if ($PSBoundParameters[$param]) {
                $passParams[$param] = $PSBoundParameters[$param]
            }
        }
        if (-Not $DebugOutputCreateOnly -And $Debug) {
            $passParams["Debug"] = $true
        }

        # Check to see if we're missing any absolutely required fields
        $props_missing = @()
        
        $field = "employee_num"
        if (-Not [string]::IsNullOrWhitespace($User.$field)) {
            if ($User.$field -is [string] -And ($Trim -Or $SyncOnEmployeeNum)) {
                $user_values[$field] = $User.$field.Trim()
            } else {
                $user_values[$field] = $User.$field
            }
        } elseif ($SyncOnEmployeeNum) {
            Throw [System.Management.Automation.ValidationMetadataException] ("[Sync-SnipeItUser] Cannot sync user [{0}] due to missing/blank field [employee_num] with -SyncOnEmployeeNum set" -f $_.username)
        } else {
            $props_missing += @($field)
        }
        
        # username, first_name, last_name
        # These should already be validated in parameter validation.
        foreach ($field in $_ALWAYSREQUIREDFIELDS) {
            if ($User.$field -is [string] -And ($Trim -Or $field -eq 'username')) {
                $user_values[$field] = $User.$field.Trim()
            } else {
                $user_values[$field] = $User.$field
            }
        }
        
        Write-Verbose ("[Sync-SnipeItUser] Processing user [{0}] with employee_num [{1}] " -f $user_values["username"],$user_values["employee_num"])
    
        # Non-required fields (except groups)
        foreach ($field in $_NONREQUIREDFIELDS) {
            if ($field -ne "groups") {
                if (-Not [string]::IsNullOrEmpty($User.$field)) {
                    if ($Trim -And $User.$field -is [string]) {
                        $user_values[$field] = $User.$field.Trim()
                    } else {
                        $user_values[$field] = $User.$field
                    }
                } else {
                    $props_missing += @($field)
                }
            }
        }
        # Set groups, if given.
        # Note this requires SnipeItPS 1.10.225 or newer
        if ($User.groups -is [int] -Or ($User.groups -is [array] -And $User.groups.Count -gt 0)) {
            $user_values["groups"] = $User.groups
        } else {
            $props_missing += @("groups")
        }
        
        # Set default notes, if given.
        if ([string]::IsNullOrEmpty($User.notes)) {
            $user_values["notes"] = $DefaultNotes
        }
        
        # Snipe-It Entity fields
        # Get company reference
        $company_id = $null
        if ($User.company_id -ne $null) {
            $company_id = $User.company_id -as [int]
        }
        if ($company_id -isnot [int]) {
            if (-Not [string]::IsNullOrWhitespace($User.company)) {
                $pp = $passParams.Clone()
                if ($DontCreateCompanyIfNotFound) {
                    $pp.Add("DontCreateIfNotFound", $true)
                } elseif (-Not [string]::IsNullOrEmpty($User.company_image)) {
                    $pp.Add("image", $User.company_image)
                }
                try {
                    $company_id = (Get-SnipeItCompanyByName $User.company @pp).id
                } catch [System.Net.WebException] {
                    Write-Error $_
                }
            }
        }
        if ($company_id -is [int]) {
            $user_values["company_id"] = $company_id
        } else {
            $props_missing += @("company","company_id")
        }
        
        # Get department reference
        $department_id = $null
        if ($User.department_id -ne $null) {
            $department_id = $User.department_id -as [int]
        }
        if ($department_id -isnot [int]) {
            if (-Not [string]::IsNullOrWhitespace($User.department)) {
                # Add company_id to create parameters, if it exists and is needed.
                $pp = $passParams.Clone()
                if ($DontCreateDepartmentIfNotFound) {
                    $pp.Add("DontCreateIfNotFound", $true)
                } elseif ($user_values["company_id"] -is [int]) {
                    $pp.Add('company_id', $user_values["company_id"])
                }
                # user field name, New-SnipeItDepartment parameter
                foreach ($field in @(@("department_image","image"),@("department_manager_id","manager_id"),@("department_location_id","location_id"),@("department_notes","notes"))) {
                    $val = $User.($field[0])
                    if (-Not [string]::IsNullOrEmpty($val)) {
                        if ($Trim -And $val -is [string]) {
                            $val = $val.Trim()
                        }
                        $pp.Add($field[1], $val)
                    }
                }
                try {
                    $department_id = (Get-SnipeItDepartmentByName $User.department @pp).id
                } catch [System.Net.WebException] {
                    Write-Error $_
                }
            }
        }
        if ($department_id -is [int]) {
            $user_values["department_id"] = $department_id
        } else {
            $props_missing += @("department","department_id")
        }
        
        # Get Location reference
        $location_id = $null
        if ($User.location_id -ne $null) {
            $location_id = $User.location_id -as [int]
        }
        if ($location_id -isnot [int]) {
            if (-Not [string]::IsNullOrWhitespace($User.location)) {
                $pp = $passParams.Clone()
                if ($DontCreateLocationIfNotFound) {
                    $pp.Add("DontCreateIfNotFound", $true)
                } else {
                    # user field name, New-SnipeItLocation parameter
                    foreach ($field in @(@("location_address","address"),@("location_address2","address2"),@("location_city","city"),@("location_state","state"),@("location_country","country"),@("location_zip","zip"),@("location_currency","currency"),@("location_parent_id","parent_id"),@("location_manager_id","manager_id"),@("location_ldap_ou","ldap_ou"),@("location_image","image"))) {
                        $val = $User.($field[0])
                        if (-Not [string]::IsNullOrEmpty($val)) {
                            if ($Trim -And $val -is [string]) {
                                $val = $val.Trim()
                            }
                            $pp.Add($field[1], $val)
                        }
                    }
                }
                $location_id = (Get-SnipeItLocationByName $User.location @pp).id
            }
        }
        if ($location_id -is [int]) {
            $user_values["location_id"] = $location_id
        } else {
            $props_missing += @("location","location_id")
        }

        # Get updated snipeit manager reference (Recursive), if it exists
        $manager_id = $null
        if ($User.manager_id -ne $null) {
            $manager_id = $User.manager_id -as [int]
        }
        if ($manager_id -isnot [int]) {
            # If we have all required parameters, recursively sync reference
            $manager = $User.manager
            if ($manager -ne $null -And -Not [string]::IsNullOrWhitespace($manager.username) -And -Not [string]::IsNullOrEmpty($manager.first_name) -And -Not [string]::IsNullOrEmpty($manager.last_name)) {
                if ($user_values["username"] -eq $manager.username.Trim()) {
                    Write-Warning("[Sync-SnipeItUser] User appears to have self set as manager (username=[{0}]), skipping setting manager to try and avoid recursive loop" -f $manager.username)
                } else {
                    Write-Verbose("[Sync-SnipeItUser] User has manager set with username [{0}] and employee_num [{1}]" -f $manager.username,$manager.employee_num)
                    # Get the updated manager reference, creating if necessary.
                    $pp = $passParams.Clone()
                    $pp.Add("RequiredSyncFields", $RequiredSyncFields)
                    $pp.Add("SyncOnEmployeeNum", $SyncOnEmployeeNum)
                    if ($OverwriteNotes -is [string]) {
                        $pp.Add("OverwriteNotes", $OverwriteNotes)
                    }
                    $manager_id = (Sync-SnipeItUser $manager @pp).id
                }
            }
        }
        if ($manager_id -is [int]) {
            if (-Not $DebugOutputCreateOnly) {
                Write-Debug ("[Sync-SnipeItUser] Found snipe-it user ID [{0}] from user manager=[{1}]" -f $manager_id,$User.manager.username)
            }
            $user_values["manager_id"] = $manager_id
        } else {
            $props_missing += @("manager","manager_id")
        }   
        
        # Grab cache, if it exists
        try {
            if ($SyncOnEmployeeNum) {
                $sp_users = Get-SnipeItEntityAll "users" -UsersKey "employee_num"
            } else {
                $sp_users = Get-SnipeItEntityAll "users" -UsersKey "username"
            }
        } catch [System.Net.WebException] {
            # Try to continue without cache.
            Write-Error $_
        }
        
        $sp_user = $null
        $sp_user2 = $null
        if ($NoCache -Or $sp_users.Count -eq 0) {
            # Check live data
            # First, check for matching username
            $sp_user = Get-SnipeItUserEx -Username $user_values["username"] -NoCache -OnErrorRetry $OnErrorRetry -SleepMS $SleepMS -DontUpdateCacheIfFound
            if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.Count -gt 1) {
                # Should never get here
                Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user by username - {0} results returned for user [{1}] with SID [{2}] " -f $sp_user.Count, $user_values["username"], $user_values["employee_num"])
            }
            if ($SyncOnEmployeeNum) {
                $sp_user2 = Get-SnipeItUserEx -EmployeeNum $user_values["employee_num"] -NoCache -OnErrorRetry $OnErrorRetry -SleepMS $SleepMS -DontUpdateCacheIfFound
                if ([string]::IsNullOrWhitespace($sp_user2.StatusCode)) {
                    if ($sp_user2.id.Count -gt 1) {
                        Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user by employee_num - {0} results returned for user [{1}] with SID [{2}] " -f $sp_user2.id.Count, $user_values["username"], $user_values["employee_num"])
                    }
                    if ($sp_user2.id -is [int]) {
                        if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.id -is [int]) {
                            if ([string]::IsNullOrWhitespace($sp_user.employee_num)) {
                                Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user [{0}] with SID [{1}] due to existing user matching username returned with blank employee_num" -f $user_values["username"],$user_values["employee_num"])
                                return -5
                            } elseif ([System.Net.WebUtility]::HtmlDecode($sp_user.employee_num).Trim() -ne $user_values["employee_num"]) {
                                Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user [{0}] with SID [{1}] due to existing user matching username with different employee_num" -f $user_values["username"],$user_values["employee_num"])
                            }
                        }
                        $sp_user = $sp_user2
                    }
                }
            }           
        } elseif ($SyncOnEmployeeNum) {
            # Also check if matching username so we can update employee_num if missing
            $sp_user = $sp_users.Values | where {[System.Net.WebUtility]::HtmlDecode($_.employee_num).Trim() -eq $user_values["employee_num"] -Or ($_.username -eq $user_values["username"] -And [string]::IsNullOrWhitespace($_.employee_num))}
        } else {
            $sp_user = $sp_users.Values | where {$_.username -eq $user_values["username"].Trim()}
        }
        # Check to see if we have any (fatal) collisions
        if ($sp_user.id.Count -gt 1) {
            if ($SyncOnEployeeNum) {
                $dupes = $sp_user | Group-Object -Property employee_num -NoElement | where {$_.Count -gt 1}
                if ($dupes.Count -gt 0) {
                    Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user [{0}] with employee_num [{1}] due to employee_num matching {2} users and -SyncOnEmployeeNum set" -f $user_values["username"],$user_values["employee_num"],$_.Count)
                }
                $dupes = $sp_user | where {[string]::IsNullOrWhitespace($_.employee_num) -eq $true}
                if ($dupes.Count -gt 0) {
                    Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user [{0}] with SID [{1}] due to existing user matching username returned with blank employee_num" -f $user_values["username"],$user_values["employee_num"])
                }
            }
            # Shouldn't ever get here
            Throw [System.Data.DuplicateNameException] ("[Sync-SnipeItUser] Cannot update user - {0} results returned for user [{1}] with employee_num [{2}] (shouldn't ever get here)" -f $sp_user.Count, $user_values["username"], $user_values["employee_num"])
        }
        if ($sp_user.id -isnot [int] -And -Not $DontCreateIfNotFound) {
            # Check to see if we need to stop due to missing fields in the user object.
            $requiredfields = $RequiredSyncFields
            if ($requiredfields -ne $null -And $requiredfields -isnot [array]) {
                $requiredfields = @($requiredfields)
            }
            $intersect = $requiredfields | where {$props_missing -contains $_}
            if ($intersect.Count -gt 0) {
                Write-Warning ("[Sync-SnipeItUser] Skipping creating new snipe-it user [{0}] due to blank/missing properties: {1}" -f $sp_user.username,($intersect -join ", "))
                return
            }
            # Add in relevant fields to the New-SnipeItUser call
            $createParams = @{}
            # Check to see if we have any snipe-it entities to add
            foreach($field in @("company_id","department_id","location_id","manager_id")) {
                if ($user_values[$field] -is [int]) {
                    $createParams.Add($field, $user_values[$field])
                }
            }
            
            # All other fields except groups.
            foreach($field in $_BUILTINFIELDS) {
                if ($field -ne "groups") {
                    $val = $user_values[$field]
                    if ($val -is [bool] -Or -Not [string]::IsNullOrEmpty($val)) {
                        $createParams.Add($field, $val)
                    }
                }
            }
            # Add groups, if they exist.
            if ($user_values["groups"] -ne $null) {
                $createParams.Add("groups", $user_values["groups"])
            }
            # Generate password if needed.
            if ([string]::IsNullOrEmpty($createParams["password"])) {
                Add-Type -AssemblyName 'System.Web'
                $createParams["password"] = [System.Web.Security.Membership]::GeneratePassword(30, 4)
            }
            
            # DEBUG
            if (-Not [string]::IsNullOrEmpty($createParams["email"])) {
                Write-Verbose("[Sync-SnipeItUser] Email parameter has been set for creating user [{0}]" -f $user_values["username"])
            }
            Write-Debug("[Sync-SnipeItUser] Create new user parameters: " + ($createParams | ConvertTo-Json -Depth 3))
            
            $count_retry = $OnErrorRetry
            while ($count_retry -ge 0) {
                $sp_user = New-SnipeitUser @createParams
                if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                    $count_retry--
                    Write-Warning ("[Sync-SnipeItUser] ERROR creating snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $user_values["username"],$sp_user.StatusCode,$sp_user.StatusDescription,$count_retry)
                } else {
                    if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.id -is [int]) {
                        Write-Verbose ("[Sync-SnipeItUser] Created new snipe-it user [{0}] with employee_num [{1}]" -f $sp_user.username,$sp_user.employee_num)
                        $update_cache = $true
                    }
                    # Break out of loop early on anything except "Too Many Requests"
                    $count_retry = -1
                }
                # Sleep before next API call
                Start-Sleep -Milliseconds $SleepMS
            }
            # All attempts failed
            if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode)) {
                Throw [System.Net.WebException] ("[Sync-SnipeItUser] Fatal ERROR creating snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $user_values["username"],$sp_user.StatusCode,$sp_user.StatusDescription)
            }
            if ($sp_user.id -is [int]) {
                Write-Verbose ("[Sync-SnipeItUser] Created new snipe-it user with ID: [{0}], username: [{1}], employee_num: [{2}]" -f $sp_user.id,$sp_user.username,$sp_user.employee_num)
            }
        } else {
            # Check for updates
            $updateParams = @{}
            # Allow updating username on SID match or blank employee_num on username match
            # But only if we're NOT syncing on username
            if ($SyncOnEmployeeNum) {
                if ($sp_user.username -ne $user_values["username"]) {
                    $updateParams.Add('username', $user_values["username"])
                } elseif ($user_values["employee_num"] -ne $sp_user.employee_num) {
                    $updateParams.Add('employee_num', $user_values["employee_num"])
                }
            } elseif (-Not [string]::IsNullOrWhitespace($user_values["employee_num"]) -And ([string]::IsNullOrWhitespace($sp_user.employee_num) -Or ($user_values["employee_num"] -ne $sp_user.employee_num -And -Not $_onlyUpdateBlankFieldsMap["employee_num"]))) {
                $updateParams.Add('employee_num', $user_values["employee_num"])
            }
            
            # Check to see if we have any snipe-it entities to add
            foreach($field in @(("company_id","company"),@("department_id","department"),@("location_id","location"),@("manager_id","manager"))) {
                $val = $user_values[$field[0]]
                $spval = $sp_user.($field[1]).id
                if ($val -is [int] -And ($spval -isnot [int] -Or ($spval -ne $val -And -Not $_onlyUpdateBlankFieldsMap[$field[0]]))) {
                    $updateParams.Add($field[0], $val)
                }
            }
            
            # All other fields except notes and groups.
            foreach($field in $_BUILTINFIELDS) {
                if ($field -ne "notes" -And $field -ne "groups" -And -Not $updateParams.ContainsKey($field)) {
                    $val = $user_values[$field]
                    $spval = $sp_user.$field
                    if ($val -is [string]) {
                        $spval = [System.Net.WebUtility]::HtmlDecode($spval)
                        if ($Trim) {
                            $val = $val.Trim()
                            $spval = $spval.Trim()
                        }
                        if (-Not [string]::IsNullOrEmpty($val) -And ([string]::IsNullOrEmpty($spval) -Or ($spval -ne $val -And -Not $_onlyUpdateBlankFieldsMap[$field]))) {
                            $updateParams.Add($field, $val)
                        }
                    } elseif ($val -is [bool] -And $val -ne $sp_user.$field) {
                        $updateParams.Add($field, $val)
                    }
                }
            }
            # Check to see if we're updating the Notes field.
            $sp_user_notes = [System.Net.WebUtility]::HtmlDecode($sp_user.notes)
            if ($Trim) {
                $sp_user_notes.Trim()
            }
            if ($user_values["notes"] -ne $null -And ([string]::IsNullOrEmpty($sp_user_notes) -Or ($sp_user_notes -ne $user_values["notes"] -And (-Not $_onlyUpdateBlankFieldsMap[$field] -Or ([string]::IsNullOrEmpty($OverwriteNotes) -Or $sp_user_notes -eq $OverwriteNotes))))) {
                $updateParams.Add('notes',$user_values["notes"])
            }
            # Check to see if we need to update groups.
            if ($user_values["groups"] -ne $null) {
                # Convert to sorted string to compare (basically what Compare-Object does)
                if ($sp_user.groups.rows.id.Count -eq 0 -Or (-Not $_onlyUpdateBlankFieldsMap['groups'] -And (($user_values['groups'] | Sort) -join ',') -ne (($sp_user.groups.rows.id | Sort) -join ','))) { 
                    $updateParams.Add("groups", $user_values["groups"])
                }
            }
            # If we have update parameters, call Set-SnipeitUser.
            if ($updateParams.Count -gt 0) {
                Write-Verbose ("[Sync-SnipeItUser] User [{0}] (employee_num: [{1}]) requires updates for fields: {2}" -f $sp_user.username,$sp_user.employee_num,($updateParams.Keys -join ", "))
                # Check to see if we need to stop due to missing required fields in the user object.
                $intersect = $RequiredSyncFields | where {$props_missing -contains $_}
                if ($intersect.Count -gt 0) {
                    Write-Warning ("[Sync-SnipeItUser] Skipping updating user [{0}] due to blank/missing properties: {1}" -f $sp_user.username,($intersect -join ", "))
                    return $sp_user
                }
                
                # DEBUG
                if (-Not [string]::IsNullOrEmpty($updateParams["email"])) {
                    Write-Verbose("[Sync-SnipeItUser] Email parameter has been set for updating user [{0}]" -f $user_values["username"])
                }
                if (-Not $DebugOutputCreateOnly) {
                    Write-Debug("[Sync-SnipeItUser] Update user parameters: " + ($updateParams | ConvertTo-Json -Depth 3))
                }
                
                $count_retry = $OnErrorRetry
                while ($count_retry -ge 0) {
                    $sp_user = Set-SnipeitUser -id $sp_user.id @UpdateParams
                    if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                        $count_retry--
                        Write-Warning ("[Sync-SnipeItUser] ERROR updating snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $user_values["username"],$sp_user.StatusCode,$sp_user.StatusDescription,$count_retry)
                    } else {
                        if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.id -is [int]) {
                            # Check to see if user was actually updated.
                            $updated_at = $sp_user.updated_at
                            if ($updated_at.datetime -ne $null) {
                                $updated_at = $updated_at.datetime
                            } elseif ($updated_at.date -ne $null) {
                                $updated_at = $updated_at.date
                            }
                            $updated_at = $updated_at -as [DateTime]
                            if ($updated_at -isnot [DateTime] -Or $updated_at -lt (Get-Date).AddMinutes(-15)) {
                                Write-Warning("[Sync-SnipeItUser] Returned user has updated date too far in the past or otherwise invalid, user may not have updated correctly (ID: [{0}], username: [{1}], employee_num: [{2}])" -f $sp_user.id, $sp_user.username, $sp_user.employee_num)
                            } else {
                                Write-Verbose ("[Sync-SnipeItUser] Updated user (ID: [{0}], username: [{1}], employee_num: [{2}])" -f $sp_user.id, $sp_user.username, $sp_user.employee_num)
                            }
                            $update_cache = $true
                        }
                        # Break out of loop early on anything except "Too Many Requests"
                        $count_retry = -1
                    }
                    # Sleep before next API call
                    Start-Sleep -Milliseconds $SleepMS
                }
                # All attempts failed
                if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode)) {
                    Throw [System.Net.WebException] ("[Sync-SnipeItUser] Fatal ERROR updating snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $user_values["username"],$sp_user.StatusCode,$sp_user.StatusDescription)
                }
            }
        }
        # Update cache with new entry, if valid
        if ($update_cache) {
            $success = Update-SnipeItCache $sp_user "users"
        }
        return $sp_user
    }
    End {
    }
}

function Remove-SnipeItInactiveUsers {
    <#
        .SYNOPSIS
        Remove, flag, or report on snipe-it users which no longer exist in the given list of users.
        
        .DESCRIPTION
        Remove, flag, or report on snipe-it users which no longer exist in the given list of users.
        
        .PARAMETER CompareUsers
        Required. The list of users to compare. These objects must at least contain the "username" fields. It's expected you pass your list of formatted users previously given to Sync-SnipeItUser.
        
        .PARAMETER CompareEmployeeNum
        Match users on employee_num instead of username.

        .PARAMETER AlsoCompareUsername
        Where employee_num is blank, also compare usernames. This is only valid if -CompareEmployeeNum is given.
        
        .PARAMETER OnlyIfLdapImport
        Only process snipe-it users where the ldap_import flag is set to $true. Useful if you have previously set this on all imported users and aren't comparing unique employee numbers, or are also comparing usernames.
        
        .PARAMETER DontDelete
        Don't delete any users, only flag them.
        
        .PARAMETER OnlyReport
        Don't touch users at all, just return the inactive users.
        
        .PARAMETER Notes
        Flag inactive users with the given notes. Defaults to "<DELETED_FROM_AD>".
        
        .PARAMETER RefreshCache
        Refresh the cache of snipe-it users before processing.
        
        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        Returns all inactive snipe-it users found.
        
        .Notes
        Possible custom thrown exceptions: [System.Net.WebException]
        
        It's recommended when deleting users to set the ldap_import flag on all your users and use -OnlyIfLdapImport if comparing by username or giving the -AlsoCompareUsername switch. This way the function shouldn't ever delete special users created by functions like Sync-SnipeItDeptUsers.
        
        .Example
        PS> $inactive_users = Remove-SnipeItInactiveUsers -CompareUsers $formatted_users -CompareEmployeeNum -OnlyReport
    #>
    [CmdletBinding(DefaultParameterSetName='Default')]
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object[]]$CompareUsers,
        
        # Match users on employee_num instead of username.
        # Also skip snipe-it users with blank employee_num.
        [parameter(Mandatory=$true, ParameterSetName='CompareEmployeeNum')]
        [switch]$CompareEmployeeNum,
        
        # Also match username if blank employee_num.
        [parameter(Mandatory=$false, ParameterSetName='CompareEmployeeNum')]
        [switch]$AlsoCompareUsername,
        
        # Only remove if the LdapImport flag is set.
        [parameter(Mandatory=$false)]
        [switch]$OnlyIfLdapImport,

        [parameter(Mandatory=$false)]
        [switch]$DontDelete,

        # Only return users that would be purged, don't touch them in snipe-it.
        [parameter(Mandatory=$false)]
        [switch]$OnlyReport,
        
        # Add notes if we're not deleting. Overwrites what's already written.
        [parameter(Mandatory=$false)]
        [string]$Notes = '<DELETED_FROM_AD>',
        
        [parameter(Mandatory=$false)]
        [switch]$RefreshCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
    }
    Process {
        $update_cache = $false
        $inactive_users = $null
        # Don't wrap this call in try/catch, since we need it -- let it error.
        if ($CompareEmployeeNum) {
            $sp_users = Get-SnipeItEntityAll "users" -UsersKey "employee_num" -RefreshCache:$RefreshCache
        } else {
            $sp_users = Get-SnipeItEntityAll "users" -UsersKey "username" -RefreshCache:$RefreshCache
        }
        
        Write-Verbose("[Remove-SnipeItInactiveUsers] Looping through {0} total users in snipe-it with settings [-CompareEmployeeNum:$CompareEmployeeNum, -AlsoCompareUsername:$AlsoCompareUsername, -OnlyIfLdapImport:$OnlyIfLdapImport]..." -f $sp_users.Count)
        if ($CompareEmployeeNum) {
            # Match by employee_num if not blank, otherwise by username
            $inactive_users = $sp_users.Values | where {(-Not $OnlyIfLdapImport -Or $_.ldap_import -eq $true) -And ((-Not [string]::IsNullOrWhitespace($_.employee_num) -And $CompareUsers.employee_num -notcontains [System.Net.WebUtility]::HtmlDecode($_.employee_num).Trim()) -Or ($AlsoCompareUsername -And [string]::IsNullOrWhitespace($_.employee_num) -And -Not [string]::IsNullOrWhitespace($_.username) -And $CompareUsers.username -notcontains $_.username))}
        } else {
            # Only compare username
            $inactive_users = $sp_users.Values | where {(-Not $OnlyIfLdapImport -Or $_.ldap_import -eq $true) -And -Not [string]::IsNullOrWhitespace($_.username) -And $CompareUsers.username -notcontains $_.username}
        }
        # Convert to array
        $inactive_users = $inactive_users | foreach { $_ }
        if ($inactive_users -isnot [array] -And $inactive_users -ne $null) {
            $inactive_users = @($inactive_users)
        }
        Write-Verbose("[Remove-SnipeItInactiveUsers] Found {0} total snipe-it users that no longer exist in synced users, and of those {1} can be deleted" -f $inactive_users.Count, ($inactive_users | where {$_.available_actions.delete -eq $true}).Count)
        if (-Not $OnlyReport) {
            $do_set_notes = (-Not [string]::IsNullOrEmpty($Notes))
            # Remove or flag users
            $purged_user_count = 0
            foreach($sp_user in $inactive_users) {
                $update_cache = $false
                $purge_user = $null
                if ($sp_user.id -is [int]) {
                    if (-Not $DontDelete -And $sp_user.available_actions.delete -eq $true) {
                        Write-Verbose("[Remove-SnipeItInactiveUsers] Preparing to delete inactive snipe-it user with ID [{0}], username [{1}], employee_num [{2}]" -f $sp_user.id,$sp_user.username,$sp_user.employee_num)
                        $count_retry = $OnErrorRetry
                        while ($count_retry -ge 0) {
                            $result = Remove-SnipeitUser -id $sp_user.id
                            if (-Not [string]::IsNullOrWhitespace($result.StatusCode) -And $result.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                                $count_retry--
                                Write-Warning ("[Remove-SnipeItInactiveUsers] ERROR removing snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $sp_user.id,$result.StatusCode,$result.StatusDescription,$count_retry)
                            } else {
                                if ([string]::IsNullOrWhitespace($result.StatusCode) -Or $result.StatusCode -eq 200 -Or $result.StatusCode -eq "OK") {
                                    Write-Verbose ("[Remove-SnipeItInactiveUsers] DELETED inactive snipe-it user with ID [{0}], username [{1}], employee_num [{2}]" -f $sp_user.id,$sp_user.username,$sp_user.employee_num)
                                    $purge_user = $sp_user.id
                                }
                                # Break out of loop early on anything except "Too Many Requests"
                                $count_retry = -1
                            }
                            # Sleep before next API call
                            Start-Sleep -Milliseconds $SleepMS
                        }
                        # All attempts failed
                        if (-Not [string]::IsNullOrWhitespace($result.StatusCode)) {
                            Throw [System.Net.WebException] ("[Remove-SnipeItInactiveUsers] Fatal ERROR removing snipeit user [ID: {0}]! StatusCode: {1}, StatusDescription: {2}" -f $sp_user.id,$result.StatusCode,$result.StatusDescription)
                        }
                        # Remove from cache
                        if ($purge_user -is [int]) {
                            $success = Update-SnipeItCache $purge_user "users" -Remove
                            $purged_user_count += 1
                        }
                    } else {
                        if ($sp_user.available_actions.delete -eq $true) {
                            $user_deletable = "deletable"
                        } else {
                            $user_deletable = "non-deletable"
                        }
                        # If not purging, remove login access and optionally add notes
                        $updateParams = @{}
                        if ($do_set_notes -And ($sp_user.notes -isnot [string] -Or $Notes.Trim() -ne [System.Net.WebUtility]::HtmlDecode($sp_user.notes).Trim())) {
                            $updateParams.Add("notes", $Notes)
                        }
                        if ($sp_user.activated -eq $true) {
                            $updateParams.Add("activated", $false)
                        }
                        $sp_user_updated = $null
                        if ($updateParams.Count -gt 0) {
                            $count_retry = $OnErrorRetry
                            while ($count_retry -ge 0) {
                                $sp_user_updated = Set-SnipeitUser -id $sp_user.id @updateParams
                                if (-Not [string]::IsNullOrWhitespace($sp_user_updated.StatusCode) -And $sp_user_updated.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                                    $count_retry--
                                    Write-Warning ("[Remove-SnipeItInactiveUsers] ERROR updating inactive snipeit user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $sp_user_updated.id,$sp_user_updated.StatusCode,$sp_user_updated.StatusDescription,$count_retry)
                                } else {
                                    if ([string]::IsNullOrWhitespace($sp_user_updated.StatusCode) -And $sp_user_updated.id -is [int]) {
                                        # Check to see if user was actually updated.
                                        $updated_at = $sp_user_updated.updated_at
                                        if ($updated_at.datetime -ne $null) {
                                            $updated_at = $updated_at.datetime
                                        } elseif ($updated_at.date -ne $null) {
                                            $updated_at = $updated_at.date
                                        }
                                        $updated_at = $updated_at -as [DateTime]
                                        if ($updated_at -isnot [DateTime] -Or $updated_at -lt (Get-Date).AddMinutes(-15)) {
                                            Write-Warning("[Remove-SnipeItInactiveUsers] Returned user has updated date too far in the past or otherwise invalid, user may not have updated correctly (ID: [{0}], username: [{1}], employee_num: [{2}]) for fields: {3}" -f $sp_user_updated.id, $sp_user_updated.username, $sp_user_updated.employee_num, ($updateParams.Keys -join ", "))
                                        } else {
                                            Write-Verbose ("[Remove-SnipeItInactiveUsers] Updated inactive snipe-it user with ID [{0}], username [{1}], employee_num [{2}] for fields: {3}" -f $sp_user_updated.id,$sp_user_updated.username,$sp_user_updated.employee_num,($updateParams.Keys -join ", "))
                                        }
                                        $update_cache = $true
                                    }
                                    # Break out of loop early on anything except "Too Many Requests"
                                    $count_retry = -1
                                }
                                # Sleep before next API call
                                Start-Sleep -Milliseconds $SleepMS
                            }
                            # All attempts failed
                            if (-Not [string]::IsNullOrWhitespace($result.StatusCode)) {
                                Throw [System.Net.WebException] ("[Remove-SnipeItInactiveUsers] Fatal ERROR updating inactive snipeit user [ID: {0}]! StatusCode: {1}, StatusDescription: {2}" -f $sp_user.id,$result.StatusCode,$result.StatusDescription)
                            }   
                        } else {
                            Write-Verbose ("[Remove-SnipeItInactiveUsers] Nothing to update for [$user_deletable] inactive snipe-it user with ID [{0}], username [{1}], employee_num [{2}]" -f $sp_user.id,$sp_user.username,$sp_user.employee_num)
                        }
                        if ($update_cache) {
                            $success = Update-SnipeItCache $sp_user_updated "users"
                        }
                    }   
                }
            }
            if (-Not $DontDelete) {
                Write-Verbose("[Remove-SnipeItInactiveUsers] Deleted [$purged_user_count] out of [{0}] inactive users from snipe-it." -f $inactive_users.Count)
            }
        }
        
        return $inactive_users
    }
    End {
    }
}

function Sync-SnipeItDeptUsers {
    <#
        .SYNOPSIS
        Create special users for each department found in snipe-it.
        
        .DESCRIPTION
        Create special users for each department found in snipe-it. Useful for assigning assets to departments as of Snipe-It version 6.x.
        
        Note this function will only create users, it will not update them.
        
        .PARAMETER Prefix
        The prefix to add in front of the username. Giving "" will result in the departments just having an underscore in front of it. Defaults to "_dept".
        
        .PARAMETER Lastname
        The last name of the departmental user, where the first name is the department name. Defaults to "_Department_".
        
        .PARAMETER SyncCompany
        When creating the departmental user, assign the department's company to it, if it exists.

        .PARAMETER SyncLocation
        When creating the departmental user, assign the department's location to it, if it exists. This will make the location show on any assigned assets.
        
        .PARAMETER Notes
        Notes used when creating departmental users. Defaults to "Created by API Script for Assigning Assets to Department".
        
        .PARAMETER SkipEmptyDepartment
        Skip creating users for empty departments. Useful if you accidentially have duplicate departments.
        
        .PARAMETER NoCache
        Ignore cache and fetch everything straight from snipe-it. This effectively refreshes the cache.
        
        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        None.
        
        .Notes
        Possible custom thrown exceptions: [System.Net.WebException]
        
        .Example
        PS> Sync-SnipeItDeptUsers -SyncCompany -SyncLocation -SkipEmptyDepartment
    #>
    param (     
        [parameter(Mandatory=$false)]
        [string]$Prefix = "_dept",
        
        [parameter(Mandatory=$false)]
        [string]$Lastname = "_Department_",
        
        [parameter(Mandatory=$false)]
        [switch]$SyncLocation,
        
        [parameter(Mandatory=$false)]
        [switch]$SyncCompany,
        
        [parameter(Mandatory=$false)]
        [string]$Notes = "Created by API Script for Assigning Assets to Department",
        
        # Don't create if the department has no users.
        [parameter(Mandatory=$false)]
        [switch]$SkipEmptyDepartment,

        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    
    # Get collection of snipeit departments and users from cache, if it exists
    $passParams = @{
        OnErrorRetry=$OnErrorRetry
        SleepMS=$SleepMS
    }
    foreach($param in @("Verbose","Debug","NoCache")) {
        if ($PSBoundParameters[$param]) {
            $passParams.Add($param, $PSBoundParameters[$param])
        }
    }
    
    $pp = @{}
    if ($NoCache) {
        $pp.Add("RefreshCache", $true)
    }
    # Don't wrap this call in try/catch, since we need it -- let it error.
    $sp_depts = Get-SnipeItEntityAll "departments" @pp
    $sp_users = Get-SnipeItEntityAll "users" -UsersKey "username" @pp
        
    # For generating random passwords
    Add-Type -AssemblyName 'System.Web' 
    # Create suffixes and usernames and check to see if they already exist
    # Also check for dupes using Group-Object
    $sp_depts.Values | where {-Not $SkipEmptyDepartment -Or $_.users_count -gt 0} | Select id,company,location,@{N="Name"; Expression={$_.Name.Trim()}} -ExcludeProperty Name | Select name,id,company,location,@{N="Suffix"; Expression={($_.Name -replace '[^A-Za-z0-9]','').ToLower()}} | where {[string]::IsNullOrEmpty($_.Suffix) -ne $true} | Select name,id,company,location,@{N="Username"; Expression={ $Prefix + '_' + $_.Suffix }} | Group-Object -Property "Username" | foreach {
        $sp_dept = $_.Group
        if ($_.Count -gt 1) {
            Write-Warning ("[Sync-SnipeItDeptUsers] Departmental username [{0}] matches {2} departments, skipping" -f $_.Name, $_.Count)
        } else {
            $name = $sp_dept.Name
            $username = $sp_dept.Username
            $id = $sp_dept.id
            Write-Verbose ("[Sync-SnipeItDeptUsers] Processing [{0}] department with departmental username [{1}]" -f $name, $username)
            $update_cache = $false
            if ($sp_dept.id -is [int]) {
                $sp_user = Get-SnipeItUserEx -Username $username @passParams
                # Username not found, so create it
                if ($sp_user.id -isnot [int]) { 
                    $count_retry = $OnErrorRetry
                    while ($count_retry -ge 0) {
                        # Optionally assign location and company IDs
                        $createParams = @{}
                        if ($SyncLocation -And $sp_dept.location.id -is [int]) {
                            $createParams.Add("location_id", $sp_dept.location.id)
                        }
                        if ($SyncCompany -And $sp_dept.company.id -is [int]) {
                            $createParams.Add("location_id", $sp_dept.company.id)
                        }
                        if ($Notes -ne $null) {
                            $createParams.Add("notes",$Notes)
                        }
                        $sp_user = New-SnipeitUser -first_name $name -last_name $Lastname -username $username -activated $false -department_id $id -password ([System.Web.Security.Membership]::GeneratePassword(30, 4)) @createParams
                        if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                            $count_retry--
                            Write-Warning ("[Sync-SnipeItDeptUsers] ERROR creating snipeit department user [{0}]! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $username,$sp_user.StatusCode,$sp_user.StatusDescription,$count_retry)
                        } else {
                            if ([string]::IsNullOrWhitespace($sp_user.StatusCode) -And $sp_user.id -is [int]) {
                                Write-Verbose ("[Sync-SnipeItDeptUsers] Created new snipe-it department user with ID: [{0}], username: [{1}] for department [{2}]" -f $sp_user.id,$sp_user.username,$name)
                                $update_cache = $true
                            }
                            # Break out of loop early on anything except "Too Many Requests"
                            $count_retry = -1
                        }
                        # Sleep before next API call
                        Start-Sleep -Milliseconds $SleepMS
                    }
                    # All attempts failed
                    if (-Not [string]::IsNullOrWhitespace($sp_user.StatusCode)) {
                        Throw [System.Net.WebException] ("[Sync-SnipeItDeptUsers] ERROR creating snipeit departmental user [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $username,$sp_user.StatusCode,$sp_user.StatusDescription)
                    }
                }
            }
            # Update cache with new entry, if valid
            if ($update_cache) {
                $success = Update-SnipeItCache $sp_user "users"
            }
        }
    }
}

<#
Asset Object returned from Update may have custom fields inserted with database names:

id                                         : <int>
name                                       : <string>
asset_tag                                  : <string>
model_id                                   : <int>
serial                                     : <string>
purchase_date                              : <DATETIME>
purchase_cost                              : <string>
order_number                               : <string>
assigned_to                                : <int>
notes                                      : <string>
image                                      : ?
user_id                                    : ?
created_at                                 : <DATETIME>
updated_at                                 : <DATETIME>
physical                                   : <int>
deleted_at                                 : <DATETIME>
status_id                                  : <int>
archived                                   : <0 or 1>
warranty_months                            : <int>
depreciate                                 : ?
supplier_id                                : <int>
requestable                                : <0 or 1>
rtd_location_id                            : <int>
accepted                                   : <0 or 1>
last_checkout                              : <DATETIME>
expected_checkin                           : <DATETIME>
company_id                                 : <int>
assigned_type                              : App\Models\User
last_audit_date                            : <DATETIME>
next_audit_date                            : <DATETIME>
location_id                                : <int>
checkin_counter                            : <int>
checkout_counter                           : <int>
requests_counter                           : <int>
model                                      : @{id=<int>; name=<string>; model_number=<string>; manufacturer_id=<int>; category_id=<int>; created_at=<DATETIME>; 
                                             updated_at=<DATETIME>; depreciation_id=; eol=; image=; deprecated_mac_address=<int>; fieldset_id=<int>; notes=; requestable=<0 or 1>; 
                                             fieldset=}
#>
function Sync-SnipeItAsset {
    <#
        .SYNOPSIS
        Syncs the given asset with snipe-it.
        
        .DESCRIPTION
        Syncs the given asset with snipe-it, searching within cache if initialized.
        
        .PARAMETER Asset
        Required. The asset object to sync.
        
        The following fields will automatically resolve to their IDs or be created if not found: manufacturer, category, model (only created if manufacturer and category are valid), supplier, company, and location. See Notes for more detail.
        
        You can give either assigned_to or assigned_id for a user id assignment, but not both.
        
        .PARAMETER SyncFields
        One or more fields to sync. These may be built-in from the Snipe-It API or custom field names. If not given the function tries to compute them from the object property members of type NoteProperty.
        
        .PARAMETER UniqueIDField
        A field which maps to a unique field on the asset, used for displaying in logs. (Default: "Serial")
        
        .PARAMETER SyncOnFieldMap
        A map of fields contained within -SyncFields to try and match and what fields to sync if the given field matches. If given a value of $true, sync all fields defined. This allows you to sync only partial information if you're missing something like Serial Number. This accepts an ordered hashtable, where it will be matched in the order of the keys. Note 'id' and 'asset_tag' will always attempted to be matched first, if they exist.
        
        Example: [ordered]@{ "Serial" = $true; "AD SID" = @("Name", "LastLogonTime", "LastLogonUser", "IPAddress") }
        
        .PARAMETER RequiredCreateFields
        One or more fields which must be non-blank before the asset is created. Defaults to @("Name","Serial").
        
        .PARAMETER OnlyUpdateBlankFields
        Only update the given fields if they are blank in snipe-it, instead of overwriting existing values.
        
        .PARAMETER DefaultModel
        A default model name to lookup if the given asset's model and model_id fields are both blank or not found. Note this will never overwrite data already in Snipe-It.
        
        .PARAMETER DefaultCreateNotes
        Default notes to use when creating if the asset's notes field is blank. Defaults to "Created by API Script".
        
        .PARAMETER DefaultCreateStatus
        Default status or status ID when creating assets. This must equate to a valid status. Defaults to 2.
        
        .PARAMETER UpdateArchivedStatus
        Change archived assets to the given status if found when syncing, otherwise display a warning.

        .PARAMETER DontCreateIfNotFound
        Don't create any assets if not found, only update existing assets.
        
        .PARAMETER DontCreateCompanyIfNotFound
        Don't create any new companies if given but not found.
        
        .PARAMETER DontCreateLocationIfNotFound
        Don't create any new locations if given but not found.
        
        .PARAMETER DontCreateSupplierIfNotFound
        Don't create any new suppliers if given but not found.
        
        .PARAMETER DebugOutputCreateOnly
        Only give debug output when creating new assets.
        
        .PARAMETER Trim
        Trim all strings, instead of just proper names.
        
        .PARAMETER NoCache
        Ignore the cache and try to fetch the user directly from Snipe-It.

        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).

        .OUTPUTS
        The processed snipe-it asset.
        
        .Notes
        Returns the processed asset from snipe-it (whether or not it was updated).
        
        [System.Data.DuplicateNameException] is thrown on duplicates.
        
        If a required entity is not found and cannot be created (statuslabel, model, etc.), [System.Data.ObjectNotFoundException] is thrown.
        
        Possible custom thrown exceptions: [System.Net.WebException], [System.Data.ObjectNotFoundException], [System.Data.DuplicateNameException]

        If an asset custom value is a hashtable, assume the field is a checkbox and compute the hashtable values as: $true (add if not present), $false (remove if present), or missing/$null (ignore existing value if present). An example would be where the existing value is "Foo, Bar" and the new values is @{ "Foo"=$false }, the result would be "Bar".
        
        When creating models it will also use values from these fields: "manufacturer, "manufacturer_id", "category", "category_id", "fieldset", "fieldset_id", "model_number", "model_image".
        
        When creating companies it will also use values from these fields: "company_image"
        
        When creating locations it will also use values from these fields: "location_address", "location_address2", "location_city", "location_country", "location_currency", "location_state", "location_zip", "location_image", "location_ldap_ou", "location_manager_id", "location_parent_id"
        
        When creating suppliers it will also use values from these fields: "supplier_address", "supplier_address2", "supplier_city", "supplier_state", "supplier_country", "supplier_zip", "supplier_phone", "supplier_fax", "supplier_email", "supplier_contact", "supplier_notes", "supplier_image"
        
        There are a lot of different ways exceptions can be thrown. If you're not using a try/catch block you may find use in -ErrorAction Continue.
        
        .Example
        PS> Sync-SnipeItAsset -Asset $Asset -SyncFields @("Name","Serial")
        
        PS> $formatted_assets | Sync-SnipeItAsset -SyncFields @("Name","Serial")
    #>
    param (
        # If asset value is a hashtable, it is assumed to be a checkbox field.
        # Hashtable values can be either missing/null, $false (remove if present), or $true (add if not present).
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object]$Asset,
        
        [parameter(Mandatory=$false)]
        [string[]]$SyncFields,
        
        [parameter(Mandatory=$false)]
        [System.Collections.IDictionary]$SyncOnFieldMap=@{ "Serial" = $true },
        
        [parameter(Mandatory=$false)]
        [string]$UniqueIDField = "Serial",
        
        [parameter(Mandatory=$false)]
        [string[]]$RequiredCreateFields=@("Name","Serial"),
        
        [parameter(Mandatory=$false)]
        [AllowEmptyCollection()]
        [string[]]$OnlyUpdateBlankFields,
        
        [parameter(Mandatory=$false)]
        [string]$DefaultModel,
        
        [parameter(Mandatory=$false)]
        [string]$DefaultCreateNotes = "Created by API Script",
        
        [parameter(Mandatory=$false)]
        [string]$DefaultCreateStatus = "2",
        
        [parameter(Mandatory=$false)]
        [string]$UpdateArchivedStatus,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateCompanyIfNotFound,

        [parameter(Mandatory=$false)]
        [switch]$DontCreateLocationIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$DontCreateSupplierIfNotFound,
        
        [parameter(Mandatory=$false)]
        [switch]$DebugOutputCreateOnly,

        [parameter(Mandatory=$false)]
        [switch]$Trim,
        
        [parameter(Mandatory=$false)]
        [switch]$NoCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    Begin {
        # Anything other than the fields below are considered custom fields
        $_UPDATEFIELDS = @("asset_tag","serial","name","notes","order_number","warranty_months","purchase_cost","purchase_date","requestable","archived","url","image","image_delete")
        $_CREATEFIELDS = @("asset_tag","serial","name","notes","order_number","warranty_months","purchase_cost","purchase_date","requestable","url","checkout_to_type")
        $_SPECIALFIELDS = @("id","customfields","assigned_to","assigned_id","status","status_id","company","location_id","Location","rtd_location_id","model","model_id","modelnumber","manufacturer","manufacturer_id","category","category_id","fieldset","fieldset_id","location_address","location_address2","location_city","location_state","location_country","location_zip","location_currency","location_parent_id","location_manager_id","location_ldap_ou","location_image","supplier","supplier_id","supplier_address","supplier_address2","supplier_city","supplier_state","supplier_country","supplier_zip","supplier_phone","supplier_fax","supplier_email","supplier_contact","supplier_notes","supplier_image")
        
        $_ALLBUILTINFIELDS = ($_UPDATEFIELDS + $_CREATEFIELDS + $_SPECIALFIELDS) | Select -Unique

        # Get list of fields to Sync if not in parameters
        if ($SyncFields.Count -gt 0) {
            $_syncFields = $SyncFields
        } elseif ($Asset -is [System.Collection.IDictionary] -Or $Asset -is [hashtable]) {
            $_syncFields = $Asset.Keys
        } else {
            $_syncFields = $Asset | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name | Out-String -Stream
        }
            
        $fieldMap = Get-SnipeItCustomFieldMap
        if ($fieldMap.Count -eq 0) {
            if (($_syncFields | where {$_ALLBUILTINFIELDS -notcontains $_}).Count -gt 0) {
                Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Asset has {0} custom fields to sync but field map is empty or null")
            }
        }
        
        # Convert $SyncOnFieldMap into an ordered hashtable
        if ($SyncOnFieldMap -is [System.Collections.Specialized.OrderedDictionary]) {
            $_syncOnFieldMap = $SyncOnFieldMap
        } else {
            $_syncOnFieldMap = [ordered]@{}
            foreach($item in $SyncOnFieldMap.GetEnumerator()) {
                $_syncOnFieldMap[$item.Name] = $item.Value
            }
        }
        # Add 'id' and 'asset_tag', if they don't already exist
        if (-Not $_syncOnFieldMap.Contains('id')) {
            $_syncOnFieldMap.Insert(0, 'asset_tag', $true)
        }
        if (-Not $_syncOnFieldMap.Contains('id')) {
            $_syncOnFieldMap.Insert(0, 'id', $true)
        }
        
        # Create a hash table for faster lookups.
        $_onlyUpdateBlankFieldsMap = @{}
        foreach ($field in $OnlyUpdateBlankFields) {
            $_onlyUpdateBlankFieldsMap[$field] = $true
        }
    }
    Process {
        $UniqueID = $Asset.$UniqueIDField
        $sp_asset = $null
        $update_cache = $false
        
        # Parameters to pass to other function calls
        $passParams = @{ 
            OnErrorRetry=$OnErrorRetry
            SleepMS=$SleepMS
        }
        foreach($param in @("NoCache","Verbose")) {
            if ($PSBoundParameters[$param]) {
                $passParams[$param] = $PSBoundParameters[$param]
            }
        }
        if (-Not $DebugOutputCreateOnly -And $Debug) {
            $passParams["Debug"] = $true
        }
        
        # Lookup Model, in case we don't have a valid Fieldset
        # Start with Model ID, if given
        $model = $Asset.model
        $sp_model = $null
        $model_id = $null
        if ($Asset.model_id -ne $null) {
            $model_id = $Asset.model_id -as [int]
        }
        if ($model_id -isnot [int]) {
            if ($Asset.model_id -ne $null) {
                Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid model_id [{0}]" -f $Asset.model_id)
            }
        } else {
            $sp_model = Get-SnipeItEntityByID $model_id "models" @passParams
        }
        if ($sp_model.id -isnot [int] -And -Not [string]::IsNullOrWhitespace($model)) {
            # Make create parameters in case we need them.
            $createParams = $passParams.Clone()
            # Add Category and Manufacturer if they both exist, otherwise don't create model if either doesn't exist
            If (-Not [string]::IsNullOrWhitespace($Asset.Category) -And -Not [string]::IsNullOrWhitespace($Asset.Manufacturer)) {
                $createParams.Add("Category", $Asset.Category)
                $createParams.Add("Manufacturer", $Asset.Manufacturer)
            } else {
                # Category and manaufacturer are required, but ignored with -DontCreateIfNotFound
                $createParams.Add("DontCreateIfNotFound", $true)
                if (-Not $DebugOutputCreateOnly) {
                    Write-Debug "[Sync-SnipeItAsset] [$UniqueID] Cannot create Model [$Model] when missing Category or Manufacturer"
                }
            }
            # Add Fieldset and ModelNumber, if they exist
            if ($Asset.fieldset_id -is [int]) {
                $createParams.Add("Fieldset", $Asset.fieldset_id)
            } elseif (-Not [string]::IsNullOrWhitespace($Asset.fieldset)) {
                $createParams.Add("Fieldset", $Asset.fieldset)
            }
            if (-Not [string]::IsNullOrEmpty($Asset.ModelNumber)) {
                $createParams.Add("ModelNumber", $Asset.ModelNumber)
            }
            # Will create if not exist.
            $sp_model = Get-SnipeItModelByName $model @createParams
        }
        # If we don't have a valid model, check if we have a DefaultModel to use
        $using_default_model = $false
        if ($sp_model.id -isnot [int] -And -Not [string]::IsNullOrWhitespace($DefaultModel)) {
            # Category and manaufacturer are required, but ignored with -DontCreateIfNotFound
            $sp_model = Get-SnipeItModelByName $DefaultModel @passParams -DontCreateIfNotFound
            if ($sp_model.id -isnot [int]) {
                Write-Warning "[Sync-SnipeItAsset] [$UniqueID] Model [$model] and default model [$DefaultModel] both returned no valid ID"
            } else {
                $using_default_model = $true
                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Model [$model] not found, using default model [$DefaultModel] with model_id [{0}]" -f $sp_model.id)
            }
        }

        # Check to see if we have any matches. Automatically checks cache if available.
        $matchfield = $null
        foreach($key in $_syncOnFieldMap.Keys) {
            if ($sp_asset.id -isnot [int]) {
                switch($key) {
                    "id" {
                        if ($Asset.id -ne $null) {
                            $field = ($Asset.id -as [int])
                            if ($field -is [int]) {
                                $sp_asset = Get-SnipeItEntityByID $Asset.id "assets" @passParams
                                if ($sp_asset.id.Count -gt 1) {
                                    # Should never get here
                                    if (-Not $DebugOutputCreateOnly) {
                                        Write-Debug ("[Sync-SnipeItAsset] [$UniqueID] Got back {0} results searching by [{1}]=[{2}], discarding result" -f $sp_asset.id.Count, $key, $Asset.id)
                                    }
                                } elseif ($sp_asset.id -is [int]) {
                                    $matchfield = $key
                                    break
                                }
                            }
                        }
                    }
                    "asset_tag" {
                        if (-Not [string]::IsNullOrWhitespace($Asset.asset_tag)) {
                            $sp_asset = Get-SnipeItAssetEx -AssetTag $Asset.asset_tag @passParams
                            if ($sp_asset.id.Count -gt 1) {
                                if (-Not $DebugOutputCreateOnly) {
                                    Write-Debug ("[Sync-SnipeItAsset] [$UniqueID] Got back {0} results searching by [{1}]=[{2}], discarding result" -f $sp_asset.id.Count, $key, $Asset.asset_tag)
                                }
                            } elseif ($sp_asset.id -is [int]) {
                                $matchfield = $key
                                break
                            }
                        }
                    }
                    "serial" {
                        if (-Not [string]::IsNullOrWhitespace($Asset.serial)) {
                            $sp_asset = Get-SnipeItAssetEx -Serial $Asset.serial @passParams
                            if ($sp_asset.id.Count -gt 1) {
                                if (-Not $DebugOutputCreateOnly) {
                                    Write-Debug ("[Sync-SnipeItAsset] [$UniqueID] Got back {0} results searching by [{1}]=[{2}], discarding result" -f $sp_asset.id.Count, $key, $Asset.serial)
                                }
                            } elseif ($sp_asset.id -is [int]) {
                                $matchfield = $key
                                break
                            }
                        }
                    }
                    "name" {
                        if (-Not [string]::IsNullOrWhitespace($Asset.name)) {
                            $sp_asset = Get-SnipeItAssetEx -Name $Asset.name @passParams
                            if ($sp_asset.id.Count -gt 1) {
                                if (-Not $DebugOutputCreateOnly) {
                                    Write-Debug ("[Sync-SnipeItAsset] [$UniqueID] Got back {0} results searching by [{1}]=[{2}], discarding result" -f $sp_asset.id.Count, $key, $Asset.name)
                                }
                            } elseif ($sp_asset.id -is [int]) {
                                $matchfield = $key
                                break
                            }
                        }
                    }
                    default {
                        # Any custom field
                        $field = $key
                        if (-Not [string]::IsNullOrEmpty($field)) {
                            $val = $Asset.$field
                            if (-Not [string]::IsNullOrWhitespace($val)) {
                                $dbfield = ($fieldMap[$field]).db_column_name
                                if (-Not [string]::IsNullOrEmpty($dbfield)) {
                                    $sp_asset = Get-SnipeItAssetEx -CustomFieldName $field -CustomDBFieldName $dbfield -CustomFieldValue $val @passParams
                                    if ($sp_asset.id.Count -gt 1) {
                                        if (-Not $DebugOutputCreateOnly) {
                                            Write-Debug ("[Sync-SnipeItAsset] [$UniqueID] Got back {0} results searching by [{1}]=[{2}], discarding result" -f $sp_asset.id.Count, $key, $val)
                                        }
                                    } elseif ($sp_asset.id -is [int]) {
                                        $matchfield = $key
                                        break
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        # Check to make sure we have a valid status ID for creating or updating
        $status_id = $Asset.status_id
        if ($status_id -ne $null) {
            $status_id = ($status_id -as [int])
            if ($status_id -isnot [int]) {
                Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid status_id [{0}]" -f $Asset.status_id)
            }
        }
        if ($status_id -isnot [int] -And -Not [string]::IsNullOrWhitespace($Asset.status)) {
            $status_id = (Get-SnipeItStatusLabelByName $Asset.status @passParams).id
            if ($status_id -isnot [int]) {
                Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid status [{0}]" -f $Asset.status)
            }
        }
        
        # Only perform the following checks if we have a match or a valid model.
        if ($matchfield -Or $sp_model.id -is [int]) {
            # get company ID, if it exists
            $company_id = $Asset.company_id
            if ($company_id -ne $null) {
                $company_id = ($company_id -as [int])
                if ($company_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid company_id [{0}]" -f $Asset.company_id)
                }
            }
            if ($company_id -isnot [int] -And -Not [string]::IsNullOrWhitespace($Asset.company)) {
                if ($DontCreateCompanyIfNotFound) {
                    $pp = $passParams.Clone()
                    $pp.Add("DontCreateIfNotFound", $true)
                } else {
                    $pp = $passParams
                }
                $company_id = (Get-SnipeItCompanyByName $Asset.company @pp).id
                if ($company_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid company [{0}]" -f $Asset.company)
                }
            }

            # get location ID, if it exists
            $location_id = $Asset.location_id
            if ($location_id -ne $null) {
                $location_id = ($location_id -as [int])
                if ($location_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid location_id [{0}]" -f $Asset.location_id)
                }
            }
            if ($location_id -isnot [int] -And -Not [string]::IsNullOrWhitespace($Asset.location)) {
                $pp = $passParams.Clone()
                if ($DontCreateLocationIfNotFound) {
                    $pp.Add("DontCreateIfNotFound", $true)
                } else {
                    # asset field name, New-SnipeItLocation parameter
                    foreach ($field in @(@("location_address","address"),@("location_address2","address2"),@("location_city","city"),@("location_state","state"),@("location_country","country"),@("location_zip","zip"),@("location_currency","currency"),@("location_parent_id","parent_id"),@("location_manager_id","manager_id"),@("location_ldap_ou","ldap_ou"),@("location_image","image"))) {
                        $val = $Asset.($field[0])
                        if (-Not [string]::IsNullOrEmpty($val)) {
                            $pp.Add($field[1], $val)
                        }
                    }
                }
                $location_id = (Get-SnipeItLocationByName $Asset.location @pp).id
                if ($location_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid location [{0}]" -f $Asset.location)
                }
            }
            
            # get supplier ID, if it exists
            $supplier_id = $Asset.supplier_id
            if ($supplier_id -ne $null) {
                $supplier_id = ($supplier_id -as [int])
                if ($supplier_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid supplier_id [{0}]" -f $Asset.supplier_id)
                }
            }
            if ($supplier_id -isnot [int] -And -Not [string]::IsNullOrWhitespace($Asset.supplier)) {
                $pp = $passParams.Clone()
                if ($DontCreateSupplierIfNotFound) {
                    $pp.Add("DontCreateIfNotFound", $true)
                } else {
                    # asset field name, New-SnipeItSupplier parameter
                    foreach ($field in @(@("supplier_address","address"),@("supplier_address2","address2"),@("supplier_city","city"),@("supplier_state","state"),@("supplier_country","country"),@("supplier_zip","zip"),@("supplier_contact","contact"),@("supplier_email","email"),@("supplier_phone","phone"),@("supplier_fax","fax"),@("supplier_image","image"))) {
                        $val = $Asset.($field[0])
                        if (-Not [string]::IsNullOrEmpty($val)) {
                            $pp.Add($field[1], $val)
                        }
                    }
                }
                $supplier_id = (Get-SnipeItsupplierByName $Asset.supplier @pp).id
                if ($supplier_id -isnot [int]) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid Supplier [{0}]" -f $Asset.supplier)
                }
            }
            
            # Add assigned_id / assigned_to, if it exists
            $assigned_id = $null
            if ($Asset.assigned_id -is [int]) {
                if ($Asset.assigned_to -is [int]) {
                    Write-Warning("[Sync-SnipeItAsset] Cannot sync assigned_id / assigned_to when both are valid numbers")
                } else {
                    $assigned_id = $Asset.assigned_id
                }
            }
            if ($assigned_id -isnot [int] -And $Asset.assigned_to -is [int]) {
                $assigned_id = $Asset.assigned_to
            }
        }
            
        if (-Not $matchfield) {
            if ($DontCreateIfNotFound) {
                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Could not find asset by any fields and -DontCreateIfNotFound is set: " + ($_syncOnFieldMap.Keys -join ", "))
            } else {
                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Could not find asset by any fields: " + ($_syncOnFieldMap.Keys -join ", "))
                
                # Asset not found, so add it.
                
                # First check to make sure we have a valid model ID.
                if ($sp_model.id -isnot [int]) {
                    Write-Warning "[Sync-SnipeItAsset] [$UniqueID] Cannot create asset due to invalid or missing model/model ID"
                    # return early
                    return
                }
                # Also check we have a valid status_id to use.
                if ($status_id -isnot [int]) {
                    $status_id = ($DefaultCreateStatus -as [int])
                    if ($status_id -isnot [int]) {
                        $status_id = (Get-SnipeItStatusLabelByName $DefaultCreateStatus @passParams).id
                        if ($status_id -isnot [int]) {
                            Throw [System.Data.ObjectNotFoundException] "[Sync-SnipeItAsset] [$UniqueID] Cannot create asset due to no valid [status_id] defined and -DefaultCreateStatus having invalid or missing status [$DefaultCreateStatus]"
                        }
                    }
                }
            
                # Fill out create parameters.
                $createParams = @{}
                $createCustomFields = @{}
                $createFields = @()
                $createAllowed = $false
                $unknownFields = @()
                foreach ($field in $_syncFields) {
                    if (-Not [string]::IsNullOrEmpty($field)) {
                        $val = $Asset.$field
                        if ($val -is [hashtable] -Or -Not [string]::IsNullOrEmpty($val)) {
                            if ($Trim -And $val -is [string]) {
                                $val = $val.Trim()
                            }
                            # Check if this is a builtin or custom field
                            if ($field -in $_CREATEFIELDS) {
                                $createParams.Add($field, $val)
                                $createFields += @($field)
                            } else {
                                # Custom fields are not added directly to create params
                                $dbfield = ($fieldMap[$field]).db_column_name
                                if ([string]::IsNullOrEmpty($dbfield)) {
                                    if ($field -notin $_ALLBUILTINFIELDS) {
                                        $unknownFields += @($field)
                                    }
                                } else {
                                    # Assumed to be a checkbox field if hashtable.
                                    if ($val -is [hashtable]) {
                                        $createCustomFields.Add($dbfield, ($val.GetEnumerator() | where {$_.Value -eq $true} | Select -ExpandProperty Name) -join ", ")
                                    } else {
                                        $createCustomFields.Add($dbfield, $val)
                                    }
                                    $createFields += @($field)
                                }
                            }
                        }
                    }
                }
                # Check to see if we're missing any required fields
                $diff = $RequiredCreateFields | where {$createFields -notcontains $_}
                if ($diff.Count -gt 0) {
                    Write-Verbose("[Sync-SnipeItAsset] [$UniqueID] Not creating new asset due to missing fields: " + ($diff -join ", "))
                } elseif ($createParams.Count -eq 0 -And $createCustomFields.Count -eq 0) {
                    Write-Verbose("[Sync-SnipeItAsset] [$UniqueID] Asset creation would result in no valid fields, skipping")
                } else {
                    # Add custom fields
                    if ($createCustomFields.Count -gt 0) {
                        $createParams.Add("customfields", $createCustomFields)
                    }
                    # Add Model ID
                    if (-Not $createParams.ContainsKey("model_id")) {
                        $createParams.Add("model_id", $sp_model.id)
                        $createFields += @("model_id")
                    }
                    # Add Status ID
                    if (-Not $createParams.ContainsKey("status_id")) {
                        $createParams.Add("status_id", $status_id)
                        $createFields += @("status_id")
                    }
                    # Add Company ID, if valid
                    if ($company_id -is [int] -And -Not $createParams.ContainsKey("company_id")) {
                        $createParams.Add("company_id", $company_id)
                        $createFields += @("company_id")
                    }
                    # Add Location ID, if valid
                    if ($location_id -is [int] -And -Not $createParams.ContainsKey("rtd_location_id")) {
                        $createParams.Add("rtd_location_id", $location_id)
                        $createFields += @("location_id")
                    }
                    # Add Supplier ID, if valid
                    if ($supplier_id -is [int] -And -Not $createParams.ContainsKey("supplier_id")) {
                        $createParams.Add("supplier_id", $supplier_id)
                        $createFields += @("supplier_id")
                    }
                    # Add assigned_id, if it exists.
                    if ($assigned_id -is [int] -And -Not $createParams.ContainsKey("assigned_id")) {
                        $createParams.Add("assigned_id", $assigned_id)
                        $createFields += @("assigned_id")
                    }
                    # Add Default Create Notes
                    if (-Not $createParams.ContainsKey("notes")) {
                        $createParams.Add("notes", $DefaultCreateNotes)
                        $createFields += @("notes")
                    }
                    # Do we have any unknown fields from $_syncFields?
                    if ($unknownFields.Count -gt 0) {
                        Write-Warning("[Sync-SnipeItAsset] [$UniqueID] Cannot find the following fields in the current fieldset: " + ($unknownFields -join ", "))
                    }

                    Write-Debug("[Sync-SnipeItAsset] [$UniqueID] Creating new asset with parameters: " + ($createParams | ConvertTo-Json -Depth 10))
                    
                    $count_retry = $OnErrorRetry
                    while ($count_retry -ge 0) {
                        $sp_asset = New-SnipeitAsset @createParams
                        if (-Not [string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                            $count_retry--
                            Write-Warning ("[Sync-SnipeItAsset] [{0}] ERROR creating snipeit asset! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $UniqueID,$sp_asset.StatusCode,$sp_asset.StatusDescription,$count_retry)
                        } else {
                            if ([string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.id -is [int]) {
                                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Created new snipe-it asset (ID: {0}) with fields: {1}" -f $sp_asset.id,($createFields -join ", "))
                                $update_cache = $true
                            }
                            # Break out of loop early on anything except "Too Many Requests"
                            $count_retry = -1
                        }
                        # Sleep before next API call
                        Start-Sleep -Milliseconds $SleepMS
                    }
                    # All attempts failed
                    if (-Not [string]::IsNullOrWhitespace($sp_asset.StatusCode)) {
                        Throw [System.Net.WebException] ("[Sync-SnipeItAsset] Fatal ERROR creating snipeit asset [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $UniqueID,$sp_asset.StatusCode,$sp_asset.StatusDescription)
                    }
                }
            }
        } else {
            # Asset found, check to see what we can update
            # "id" and "asset_tag" matches sync ALL fields by default.
            if ($_syncOnFieldMap[$matchfield] -is [bool] -And $_syncOnFieldMap[$matchfield]) {
                # Select all fields if true
                $updateFields = $_syncFields
            } else {
                $updateFields = $_syncOnFieldMap[$matchfield]
            }
            # Constrain update fields to only valid ones
            $updateFields = $updateFields | where {$_UPDATEFIELDS -contains $_ -Or $fieldMap.Keys -contains $_}
            $updateParams = @{}
            $updateCustomFields = @{}
            $fieldsToUpdate = @()
            $unknownFields = @()
            foreach ($field in $updateFields) {
                if (-Not [string]::IsNullOrEmpty($field)) {
                    $val = $Asset.$field
                    if ($Trim -And $val -is [string]) {
                        $val = $val.Trim()
                    }
                    if ($val -is [hashtable] -Or -Not [string]::IsNullOrEmpty($val)) {
                        if ($val -isnot [hashtable] -And $field -in $_UPDATEFIELDS) {
                            # Updating a built-in field
                            $spval = $sp_asset.$field
                            if ([string]::IsNullOrEmpty($spval) -Or -Not $_onlyUpdateBlankFieldsMap[$field]) {
                                # Check to see if it's in date format and if so, cast to compare.
                                if ($spval.date -ne $null -Or $spval.datetime -ne $null) {
                                    $spval_dt = $sp_asset.$field.formatted -as [DateTime]
                                    $val_dt = $val -as [DateTime]
                                    if ($val_dt -eq $spval_dt) {
                                         # No update needed
                                         $val = $null
                                     }
                                } else {
                                    if ($spval -is [string]) {
                                        $spval = [System.Net.WebUtility]::HtmlDecode($spval)
                                        if ($Trim) {
                                            $spval = $spval.Trim()
                                        }
                                    }
                                    if ($val -eq $spval) {
                                        # No update needed
                                        $val = $null
                                    }
                                }
                                if ($val -ne $null) {
                                    $updateParams.Add($field, $val)
                                    $fieldsToUpdate += @($field)
                                }
                            }
                        } else {
                            # Updating a custom field
                            $dbfield = ($fieldMap[$field]).db_column_name
                            if ([string]::IsNullOrEmpty($dbfield)) {
                                if ($field -notin $_ALLBUILTINFIELDS) {
                                    $unknownFields += @($field)
                                }
                            } else {
                                $spval = $sp_asset.custom_fields.$field.value
                                $spval_format = $sp_asset.custom_fields.$field.field_format
                                if ([string]::IsNullOrEmpty($spval) -Or -Not $_onlyUpdateBlankFieldsMap[$field]) {
                                    if ($Trim -And $spval -is [string]) {
                                        $spval = $spval.Trim()
                                    }
                                    # Assumed to be a checkbox field if hashtable.
                                    if ($val -is [hashtable]) {
                                        # Get difference of current values.
                                        # Hashtable values can be either $true (add if not present), $false (remove if present), or missing/$null (ignore if present).
                                        # First, convert the existing string data to an array.
                                        $spval = ([System.Net.WebUtility]::HtmlDecode($spval) -split ", ")
                                        # Do an effective outer join of the two arrays on value -ne $false.
                                        $val = ($val.Keys + $spval) | Select -Unique | where {-Not [string]::IsNullOrWhitespace($_) -And $val[$_] -ne $false}
                                        if ($val.Count -eq 0) {
                                            # No update needed
                                            $val = $null
                                        } else {
                                            # Turn back into sorted string data for comparison
                                            $val = ($val | Sort) -join ", "
                                            $spval = ($spval | Sort) -join ", "
                                            if ($val -eq $spval -Or ([string]::IsNullOrWhitespace($spval) -And [string]::IsNullOrWhitespace($val))) {
                                                # No update needed.
                                                $val = $null
                                            }
                                        }
                                    # All other values.
                                    } else {
                                        # Check if it's a date format and cast to DateTime to compare.
                                        if ($spval_format -eq "DATE") {
                                            $spval_dt = $spval -as [DateTime]
                                            $val_dt = $val -as [DateTime]
                                            if ($val_dt -eq $spval_dt) {
                                                # No update needed.
                                                $val = $null
                                            }
                                        } else {
                                            if ($spval -is [string]) {
                                                $spval = [System.Net.WebUtility]::HtmlDecode($spval)
                                            }
                                            if ($val -eq $spval) {
                                                # No update needed.
                                                $val = $null
                                            }
                                        }
                                    }
                                    if ($val -ne $null) {
                                        $updateCustomFields.Add($dbfield, $val)
                                        $fieldsToUpdate += @($field)
                                    }
                                }
                            }
                        }
                    }
                }
            }
            # Do we have to update the rtd_location_id (default location id)?
            if ($location_id -is [int] -And ($sp_asset.rtd_location.id -isnot [int] -Or ($sp_asset.rtd_location.id -ne $location_id -And -Not $_onlyUpdateBlankFieldsMap['location'] -And -Not $_onlyUpdateBlankFieldsMap['location_id'] -And -Not $_onlyUpdateBlankFieldsMap['rtd_location_id'])))  {
                $updateParams.Add("rtd_location_id", $location_id)
                $fieldsToUpdate += @("rtd_location_id")
            }
            # Do we have to update the company_id?
            if ($company_id -is [int] -And ($sp_asset.company.id -isnot [int] -Or ($sp_asset.company.id -ne $company_id -And -Not $_onlyUpdateBlankFieldsMap['company'] -And -Not $_onlyUpdateBlankFieldsMap['company_id']))) {
                $updateParams.Add("company_id", $company_id)
                $fieldsToUpdate += @("company_id")
            }
            # Do we have need to update the status ID?
            if ($status_id -is [int] -And $sp_asset.status_label.id -ne $status_id) {
                $updateParams.Add("status_id", $status_id)
                $fieldsToUpdate += @("status_id")
            } elseif ($sp_asset.status_label.status_type -eq 'archived') {
                # For archived assets, check if $UpdateArchivedStatus is set
                if (-Not [string]::IsNullOrWhitespace($UpdateArchivedStatus)) { 
                    $status_id = ($UpdateArchivedStatus -as [int])
                    if ($status_id -isnot [int]) {
                        $status_id = (Get-SnipeItStatusLabelByName $UpdateArchivedStatus @passParams).id
                        if ($status_id -isnot [int]) {
                            Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Invalid status given to -UpdateArchivedStatus [{0}]" -f $UpdateArchivedStatus)
                        } elseif ($sp_asset.status_label.id -ne $status_id) {
                            $updateParams.Add("status_id", $status_id)
                            $fieldsToUpdate += @("status_id")
                        }
                    }
                }
            }
            # Do we have need to update the model ID?
            # Only update when not using the default model.
            if ($sp_model.id -is [int] -And $sp_asset.model.id -ne $sp_model.id -And -Not $using_default_model) {
                $updateParams.Add("model_id", $sp_model.id)
                $fieldsToUpdate += @("model_id")
            }
            # Do we need to update the assigned_to ID?
            if ($assigned_id -is [int] -And -Not $updateParams.ContainsKey("assigned_id")) {
                $updateParams.Add("assigned_to", $assigned_id)
                $fieldsToUpdate += @("assigned_to")
            }
            
            # Add custom fields to update parameters
            if ($updateCustomFields.Count -gt 0) {
                $updateParams.Add("customfields", $updateCustomFields)
            }
            # Do we have any unknown fields from $_syncFields?
            if ($unknownFields.Count -gt 0) {
                Write-Warning("[Sync-SnipeItAsset] [$UniqueID] Cannot find the following fields in the current fieldset: " + ($unknownFields -join ", "))
            }
            # Only update if we have something to update
            if ($updateParams.Count -le 0) {
                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Nothing to update for snipe-it asset ID [{0}] (matched by: {1})" -f $sp_asset.id, $matchfield)
            } else {
                # Give a warning if we're updating archived assets.
                if ($sp_asset.status_label.status_type -eq 'archived' -And -Not $updateParams.ContainsKey('status_id')) {
                    Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Updating archived asset ID [{0}] without updating status" -f $sp_asset.id)
                }
                if (-Not $DebugOutputCreateOnly) {
                    Write-Debug("[Sync-SnipeItAsset] [$UniqueID] Updating snipe-it asset ID [{0}] with parameters: {1}" -f $sp_asset.id,($updateParams | ConvertTo-Json -Depth 10))
                }
                $count_retry = $OnErrorRetry
                while ($count_retry -ge 0) {
                    $sp_asset = Set-SnipeitAsset -id $sp_asset.id @updateParams
                    if (-Not [string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                        $count_retry--
                        Write-Warning ("[Sync-SnipeItAsset] [{0}] ERROR updating snipeit asset! StatusCode: {1}, StatusDescription: {2}, Retries Left: {3}" -f $UniqueID,$sp_asset.StatusCode,$sp_asset.StatusDescription,$count_retry)
                    } else {
                        if ([string]::IsNullOrWhitespace($sp_asset.StatusCode) -And $sp_asset.id -is [int]) {
                            # Check to see if asset was actually updated.
                            $updated_at = $sp_asset.updated_at
                            if ($updated_at.datetime -ne $null) {
                                $updated_at = $updated_at.datetime
                            } elseif ($updated_at.date -ne $null) {
                                $updated_at = $updated_at.date
                            }
                            $updated_at = $updated_at -as [DateTime]
                            if ($updated_at -isnot [DateTime] -Or $updated_at -lt (Get-Date).AddMinutes(-15)) {
                                Write-Warning ("[Sync-SnipeItAsset] [$UniqueID] Returned asset with ID [{0}] (matched by: {1}) has updated date is too far in the past or otherwise invalid, may not updated correctly for fields: {2}" -f $sp_asset.id,$matchfield,($fieldsUpdated -join ", "))
                            } else {
                                Write-Verbose ("[Sync-SnipeItAsset] [$UniqueID] Updated snipe-it asset ID [{0}] (matched by: {1}) for fields: {2}" -f $sp_asset.id,$matchfield,($fieldsToUpdate -join ", "))
                            }
                            $update_cache = $true
                            # Refetch asset from server due to https://github.com/snipe/snipe-it/issues/11725
                            $sp_asset_refetched = Get-SnipeItEntityByID $sp_asset.id "assets" -NoCache
                            if ([string]::IsNullOrWhitespace($sp_asset_refetched.StatusCode) -And $sp_asset_refetched.id -is [int]) {
                                $sp_asset = $sp_asset_refetched
                            }
                        }
                        # Break out of loop early on anything except "Too Many Requests"
                        $count_retry = -1
                    }
                    # Sleep before next API call
                    Start-Sleep -Milliseconds $SleepMS
                }
                # All attempts failed
                if (-Not [string]::IsNullOrWhitespace($sp_asset.StatusCode)) {
                    Throw [System.Net.WebException] ("Fatal ERROR updating snipeit asset [{0}]! StatusCode: {1}, StatusDescription: {2}" -f $UniqueID,$sp_asset.StatusCode,$sp_asset.StatusDescription)
                }
            }
        }
        
        # Add to cache
        if ($update_cache) {
            $sp_asset = $sp_asset | Restore-SnipeItAssetCustomFields
            $success = Update-SnipeItCache $sp_asset "assets"
        }
        
        return $sp_asset
    }
    End {
    }
}

function Format-SnipeItEntity {
    <#
        .SYNOPSIS
        Formats a snipe-it entity/object for output to file, like CSV.
        
        .DESCRIPTION
        Formats a snipe-it entity/object for output to file, like CSV.
        
        .PARAMETER Entity
        Required. The snipe-it entity/object to process.

        .OUTPUTS
        A snipe-it object with the result formatted for output to file.
        
        .Example
        PS> $Users | Format-SnipeItEntity | Export-CSV -NoTypeInformation "users.csv"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object]$Entity
    )
    Begin {
        function _formatFunc($val) {
            if ($val.name -is [string]) {
                [System.Net.WebUtility]::HtmlDecode($val.name)
            } elseif ($val.formatted -is [string]) {
                $val.formatted
            } elseif ($val -is [string]) {
                [System.Net.WebUtility]::HtmlDecode($val)
            } elseif ($val -ne $null) {
                ($val | ConvertTo-Json -Depth 10) -replace "[\t\n\r]",' '
            }
        }
    }
    Process {
        $existingMemberNames = $Entity | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name | Out-String -Stream
        $selectArray = $existingMemberNames | foreach {
            @{N=[string]$_; Expression=[Scriptblock]::Create("_formatFunc(`$_.'$_')") }
        }
        # Convert into array if needed
        if ($existingMemberNames -isnot [array]) {
            $existingMemberNames = @($existingMemberNames)
        }
        return $Entity | Select $selectArray -ExcludeProperty $existingMemberNames
    }
    End {
    }
}

function Format-SnipeItAsset {
    <#
        .SYNOPSIS
        Formats a snipe-it asset for output to file, like CSV.
        
        .DESCRIPTION
        Formats a snipe-it asset for output to file, like CSV.
        
        .PARAMETER Asset
        Required. The asset object to process.
        
        .PARAMETER AddDepartment
        Add the user's department to the output, if it exists.
        
        .OUTPUTS
        An asset object with the result formatted for output to file.
        
        .Example
        PS> $Assets | Format-SnipeItAsset -AddDepartment | Export-CSV -NoTypeInformation "assets.csv"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object]$Asset,
        
        [parameter(Mandatory=$false)]
        [switch]$AddDepartment
    )
    Begin {
        function _formatFunc($val) {
            if ($val.username -is [string]) {
                "{0} ({1})" -f [System.Net.WebUtility]::HtmlDecode($val.name), $val.username
            } elseif ($val.name -is [string]) {
                if ($val.status_meta -is [string]) {
                    "{0} ({1})" -f [System.Net.WebUtility]::HtmlDecode($val.name), $val.status_meta
                } else {
                    [System.Net.WebUtility]::HtmlDecode($val.name)
                }
            } elseif ($val.formatted -is [string]) {
                $val.formatted
            } elseif ($val -is [string]) {
                [System.Net.WebUtility]::HtmlDecode($val)
            } elseif ($val -ne $null) {
                ($val | ConvertTo-Json -Depth 10) -replace "[\t\n\r]",' '
            }
        }
    }
    Process {
        $existingMemberNames = $Asset | Get-Member -MemberType NoteProperty | where {$_.Name -ne "custom_fields"} | Select -ExpandProperty Name | Out-String -Stream
        $selectArray = $existingMemberNames | foreach {
            @{N=[string]$_; Expression=[Scriptblock]::Create("_formatFunc(`$_.'$_')") }
        }
        # Add Department column if -AddDepartment is set (will be null if not assigned to a user)
        if ($AddDepartment) {
            $department = $null
            if ($Asset.assigned_to.id -is [int] -And $Asset.assigned_to.type -eq "user") {
                $department = (Get-SnipeItEntityByID $Asset.assigned_to.id "users").department.name
            }
            $selectArray += @{N="Department"; Expression={ [System.Net.WebUtility]::HtmlDecode($department) }}
        }
        # Add custom field names back into the main object
        if ($Asset.custom_fields -is [PSObject]) {
            $customfields = $Asset.custom_fields | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name | foreach {
                $skip_processing = $null
                if ($existingMemberNames -contains $_) {
                    $name = "CustomField_" + $_
                    if ($existingMemberNames -contains $name) {
                        $skip_processing = $true
                    }
                } else {
                    $name = [string]$_
                }
                if (-Not $skip_processing) {
                    @{N=$name; Expression=[Scriptblock]::Create("`$_.custom_fields.'$_'.value") }
                }
            }
            if ($customfields.Count -gt 0) {
                $selectArray += $customfields
            }
        }
        # Convert into array if needed
        if ($existingMemberNames -isnot [array]) {
            $existingMemberNames = @($existingMemberNames)
        }
        $existingMemberNames += "custom_fields"
        return $Asset | Select $selectArray -ExcludeProperty $existingMemberNames
    }
    End {
    }
}

function Remove-SnipeItInactiveEntity {
    <#
        .SYNOPSIS
        Removes all unassigned snipe-it entities of the given type(s).
        
        .DESCRIPTION
        Removes all unassigned snipe-it entities of the given type(s).
        
        .PARAMETER EntityTypes
        Required. One or more of the types of entities supported by the Snipe-It API. This is always in the form of their API name (IE, "departments").
        
        .PARAMETER ExcludeNames
        Exclude the given names (exact match).
        
        .PARAMETER ExcludeNamePattern
        Exclude names matching the given pattern.
        
        .PARAMETER OnErrorRetry
        The number of times to retry if we get certain error codes like "Too Many Requests" (default: 3). Give 0 to never retry.

        .PARAMETER SleepMS
        The number of milliseconds to sleep after each API call (default: 1000ms).
        
        .OUTPUTS
        None.
        
        .Example
        PS> Remove-SnipeItInactiveEntity "departments","companies","locations"
    #>
    param (
        [parameter(Mandatory=$true,
                   Position=0)]
        [ValidateSet("departments","locations","companies","manufacturers","categories","models","fields","suppliers")]
        [string[]]$EntityTypes,
        
        [parameter(Mandatory=$false)]
        [string[]]$ExcludeNames,
        
        [parameter(Mandatory=$false)]
        [string]$ExcludeNamePattern,
        
        [parameter(Mandatory=$false)]
        [switch]$RefreshCache,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$OnErrorRetry=3,
        
        [parameter(Mandatory=$false)]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$SleepMS=1000
    )
    
    $passParams = @{}
    foreach ($param in @("RefreshCache","Debug","Verbose")) {
        if ($PSBoundParameters[$param]) {
            $passParams[$param] = $true
        }
    }
            
    Write-Debug("[Remove-SnipeItInactiveEntity] ExcludeNames: [{0}], ExcludeNamePattern: [{1}]" -f ($ExcludeNames -join ', '), $ExcludeNamePattern)
    foreach ($entityType in $EntityTypes) {
        $removeFunc = $null
        switch($entityType) {
            "departments" {
                $removeFunc = "Remove-SnipeitDepartment"
            }
            "locations" {
                $removeFunc = "Remove-SnipeitLocation"
            }
            "companies" {
                $removeFunc = "Remove-SnipeitCompany"
            }
            "manufacturers" {
                $removeFunc = "Remove-SnipeitManufacturer"
            }
            "categories" {
                $removeFunc = "Remove-SnipeitCategory"
            }
            "models" {
                $removeFunc = "Remove-SnipeitModel"
            }
            "fields" {
                $removeFunc = "Remove-SnipeitCustomField"
            }
            "suppliers" {
                $removeFunc = "Remove-SnipeitSupplier"
            }
            default {
                Throw [System.Management.Automation.ValidationMetadataException] "[Remove-SnipeItInactiveEntity] Unsupported EntityType: $entityType (should never get here?)"
            }
        }
        $sp_entities = (Get-SnipeItEntityAll $entityType @passParams).Values 
        if ($sp_entities.id.Count -gt 0) {
            $deletable_entities = $sp_entities | where {$_.available_actions.delete -eq $true -And ([string]::IsNullOrEmpty($_.name) -Or (([string]::IsNullOrEmpty($ExcludeNames) -Or $ExcludeNames -notcontains $_.name) -And ([string]::IsNullOrEmpty($ExcludeNamePattern) -Or $_name -notmatch $ExcludeNamePattern)))}
            if ($deletable_entities.id.Count -eq 0) {
                Write-Verbose("[Remove-SnipeItInactiveEntity] No matching inactive entities found for [$entityType]")
            } elseif (-Not [string]::IsNullOrEmpty($removeFunc)) {
                Write-Verbose("[Remove-SnipeItInactiveEntity] Found [{0}] matching deletable inactive [{1}] in snipe-it, removing..." -f $deletable_entities.id.Count, $entityType)
                foreach ($entity in $deletable_entities) {
                    if ($entity.id -is [int]) {
                        $count_retry = $OnErrorRetry
                        while ($count_retry -ge 0) {
                            # TODO: Suggest update to SnipeitPS suppress these warnings
                            $result = &$removeFunc -id $entity.id -WarningAction SilentlyContinue
                            if (-Not [string]::IsNullOrWhitespace($result.StatusCode) -And $result.StatusCode -in $SNIPEIT_RETRY_ON_STATUS_CODES) {
                                $count_retry--
                                Write-Warning ("[Remove-SnipeItInactiveEntity] ERROR removing snipeit inactive [{0}] with name [{1}], id {2}! StatusCode: {3}, StatusDescription: {4}, Retries Left: {5}" -f $entityType,$entity.name,$entity.id,$result.StatusCode,$result.StatusDescription,$count_retry)
                            } else {
                                if ([string]::IsNullOrWhitespace($result.StatusCode) -Or $result.StatusCode -eq 200 -Or $result.StatusCode -eq 'OK') {
                                    Write-Verbose ("[Remove-SnipeItInactiveEntity] Removed inactive [{0}] from snipe-it with name [{1}], id {2}" -f $entityType,$entity.name,$entity.id)
                                    $update_cache = $true
                                }
                                # Break out of loop early on anything except "Too Many Requests"
                                $count_retry = -1
                            }
                            # Sleep before next API call
                            Start-Sleep -Milliseconds $SleepMS
                        }
                        # Throw exception on consistent errors
                        if (-Not $update_cache) {
                            Throw [System.Net.WebException] ("[Remove-SnipeItInactiveEntity] Fatal ERROR removing inactive [{0}] from snipe-it where name=[{1}] and id={2}! StatusCode: {3}, StatusDescription: {4}" -f $entityType,$entity.name,$entity.id,$result.StatusCode,$result.StatusDescription)
                        } else {
                            $success = Update-SnipeItCache $entity.id $entityType -Remove
                        }
                    }
                }
            }
        }
    }
}
