#requires -version 3
<#
    Get Director data from a Delivery controller

    Modification History:

    @guyrleech  04/03/20  Added Citrix Cloud capability
    @guyrleech  05/03/20  Made -join run recursive. Fixed bug with date ranges
#>

<#
.SYNOPSIS

Send queries to a Citrix Delivery Controller or Citrix Cloud and present the results back as PowerShell objects

.PARAMETER ddc

The Delivery Controller to query

.PARAMETER customerId

The Citrix Cloud customerid to query

.PARAMETER authToken

The Citrix cloud authentication token to use. If not specified will prompt for credentials

.PARAMETER credential

A credential object to use to connect to the specific Delivery Controller

.PARAMETER query

The item to query. If not specified, all services/queries available will be returned

.PARAMETER join

Where there are id's in the retrieved query, look up the objects for the id's and substitute them in the output

.PARAMETER last

Only retrieve items such as sessions or connections which have been created in the last specified period

.PARAMETER username

The username to use when querying the Delivery Controller. If not specified the user running the script will be used

.PARAMETER password

The password for the account used to query the Delivery Controller. If the %RandomKey% environment variable is set, its contents will be used as the password

.PARAMETER protocol

The protocol to use to query the Delivery Controller

.PARAMETER oDataVersion

The version of OData to use in the query. If not specified then the script will work out which is the latest available and use that

.EXAMPLE

'.\Get Citrix OData data.ps1' -ddc ctxddc01 

Send a web request to the Delivery Controller ctxddc01 and retrieve the list of all available services

.EXAMPLE

'.\Get Citrix OData data.ps1' -ddc ctxddc01 -query Users

Send a web request to the Delivery Controller ctxddc01 to retrieve the list of all users

.EXAMPLE

'.\Get Citrix OData data.ps1' -ddc ctxddc01 -query connections -join -last 7d

Send a web request to the Delivery Controller ctxddc01 to retrieve the list of all connections created within the last 7 days and cross reference any id's returned

.EXAMPLE

'.\Get Citrix OData data.ps1' -customerid yourcloudid -query connections -join -last 7d

Prompt for credentials for the Citrix Cloud customer with id yourcloudid, send a web request to the Citrix Cloud to retrieve the list of all connections created within the last 7 days and cross reference any id's returned

.NOTES

https://developer-docs.citrix.com/projects/monitor-service-odata-api/en/latest/

If an auth token is not passed, the Citrix Remote PowerShell SDK must be available in order to get an auth token - https://www.citrix.com/downloads/citrix-cloud/product-software/xenapp-and-xendesktop-service.html

#>

Param
(
    [Parameter(ParameterSetName='ddc',Mandatory=$true)]
    [string]$ddc ,
    [Parameter(ParameterSetName='cloud',Mandatory=$true)]
    [string]$customerid ,
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$authtoken ,
    [string]$query ,
    [switch]$join ,
    [string]$last ,
    [datetime]$from ,
    [datetime]$to ,
    [System.Management.Automation.PSCredential]$credential ,
    [string]$username , 
    [string]$password ,
    [Parameter(ParameterSetName='ddc',Mandatory=$false)]
    [ValidateSet('http','https')]
    [string]$protocol = 'http' ,
    [int]$oDataVersion = 4
)

## map tables to the date stamp we will filter on
[hashtable]$dateFields = @{
     'Session' = 'StartDate'
     'Connection' = 'BrokeringDate'
     'ConnectionFailureLog' = 'FailureDate'
}

## Modified from code at https://jasonconger.com/2013/10/11/using-powershell-to-retrieve-citrix-monitor-data-via-odata/
Function Invoke-ODataTransform
{
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        $records
    )

    Begin
    {
        $propertyNames = $null

        [int]$timeOffset = if( (Get-Date).IsDaylightSavingTime() ) { 1 } else { 0 }
    }

    Process
    {
        if( $records -is [array] )
        {
            if( ! $propertyNames )
            {
                $properties = ($records | Select -First 1).content.properties
                if( $properties )
                {
                    $propertyNames = $properties | Get-Member -MemberType Properties | Select -ExpandProperty name
                }
                else
                {
                    // v4+
                    $propertyNames = 'NA' -as [string]
                }
            }
            if( $propertyNames -is [string] )
            {
                $records | Select -ExpandProperty value
            }
            else
            {
                ForEach( $record in $records )
                {
                    $h = @{ 'ID' = $record.ID }
                    $properties = $record.content.properties

                    ForEach( $propertyName in $propertyNames )
                    {
                        $targetProperty = $properties.$propertyName
                        if($targetProperty -is [Xml.XmlElement])
                        {
                            try
                            {
                                $h.$propertyName = $targetProperty.'#text'
                                ## see if we need to adjust for daylight savings
                                if( $timeOffset -and ! [string]::IsNullOrEmpty( $h.$propertyName ) -and $targetProperty.type -match 'DateTime' )
                                {
                                    $h.$propertyName = (Get-Date -Date $h.$propertyName).AddHours( $timeOffset )
                                }
                            }
                            catch
                            {
                                ##$_
                            }
                        }
                        else
                        {
                            $h.$propertyName = $targetProperty
                        }
                    }

                    [PSCustomObject]$h
                }
            }
        }
        elseif( $records -and $records.PSObject.Properties[ 'value' ] ) ##JSON
        {
            $records.value
        }
    }
}

Function Get-DateRanges
{
    Param
    (
        [string]$query ,
        $from ,
        $to ,
        [switch]$selective ,
        [int]$oDataVersion
    )
    
    $field = $dateFields[ ($query -replace 's$' , '') ]
    if( ! $field )
    {
        if( $selective )
        {
            return $null ## only want specific ones
        }
        $field = 'CreatedDate'
    }
    if( $oDataVersion -ge 4 )
    {
        if( $from )
        {
            "()?`$filter=$field ge $(Get-Date -date $from -format s)Z"
        }
        if( $to )
        {
            "and $field le $(Get-Date -date $to -format s)Z"
        }
    }
    else
    {
        if( $from )
        {
            "()?`$filter=$field ge datetime'$(Get-Date -date $from -format s)'"
        }
        if( $to )
        {
            "and $field le datetime'$(Get-Date -date $to -format s)'"
        }
    }
}

Function Resolve-CrossReferences
{
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        $properties ,
        [switch]$cloud
    )
    
    Process
    {
        $properties | Where-Object { ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(SessionKey)$' ) -and ! [string]::IsNullOrEmpty( $Matches[1] ) }  | Select-Object -Property Name | ForEach-Object `
        {
            [string]$id = $Matches[1]
            [bool]$current = $false
            if( $id -match '^Current(.*)$' )
            {
                $current = $true
                $id = $Matches[1]
            }
            elseif( $id -eq 'SessionKey' )
            {
                $id = 'Session'
            }

            if( ! $tables[ $id ] -and ! $alreadyFetched[ $id ] )
            {
                if( $cloud )
                {
                    $params.uri = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}s" -f $protocol , $customerid , $version ,  $id ) ## + (Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion)
                }
                else
                {
                    $params.uri = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}s" -f $protocol , $ddc , $version , $id ) ## + (Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion)
                }

                ## save looking up again, especially if it errors as we are not looking up anything valid
                $alreadyFetched.Add( $id , $id )

                [hashtable]$table = @{}
                try
                {
                    Invoke-RestMethod @params | Invoke-ODataTransform | ForEach-Object `
                    {
                        ## add to hash table keyed on its id
                        ## ToDo we need to go recursive to see if any of these have Ids that we need to resolve without going infintely recursive
                        $object = $_
                        [string]$thisId = $null
                        [string]$keyName = $null

                        if( $object.PSObject.Properties[ 'id' ] )
                        {
                            $thisId = $object.Id
                            $keyName = 'id'
                        }
                        elseif( $object.PSObject.Properties[ 'SessionKey' ] )
                        {
                            $thisId = $object.SessionKey
                            $keyname = 'SessionKey'
                        }

                        if( $thisId )
                        {
                            [string]$key = $(if( $thisId -match '\(guid''(.*)''\)$' )
                                {
                                    $Matches[ 1 ]
                                }
                                else
                                {
                                    $thisId
                                })
                            $object.PSObject.properties.remove( $key )
                            $table.Add( $key , $object )
                        }

                        ## Look at other properties to figure if it too is an id and grab that table too if we don't have it already
                        ForEach( $property in $object.PSObject.Properties )
                        {
                            if( $property.MemberType -eq 'NoteProperty' -and $property.Name -ne $keyName -and $property.Name -ne 'sid' -and $property.Name -match '(.*)Id$' )
                            {
                                $property | Resolve-CrossReferences -cloud:$cloud
                            }
                        }
                    }
                    if( $table.Count )
                    {
                        Write-Verbose -Message "Adding table $id with $($table.Count) entries"
                        $tables.Add( $id , $table )
                    }
                }
                catch
                {
                    $nop = $null
                }
            }
        }
    }
}

Function Resolve-NestedProperties
{
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        $properties ,
        $previousProperties
    )
    
    Process
    {
        $properties | Where-Object { $_.Name -ne 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and ! [string]::IsNullOrEmpty( $Matches[1] ) } | ForEach-Object `
        {
            $property = $_
                
            if( ! [string]::IsNullOrEmpty( ( $id = ( $Matches[1] -replace '^Current' , '')) ))
            {
                if ( $table = $tables[ $id ] )
                {
                    if( $property.Value -and ( $item = $table[ ($property.Value -as [string]) ]))
                    {
                        $datum.PSObject.properties.remove( $property )
                        $item.PSObject.Properties | ForEach-Object `
                        {
                            [pscustomobject]@{ "$id.$($_.Name)" = $_.Value }
                            if( $_.Name -ne $property.Name -and ( ! $previousProperties -or ! ( $previousProperties | Where-Object Name -eq $_.Name ))) ## don't lookup self or a key if it was one we previously looked up
                            {
                                Resolve-NestedProperties -properties $_ -previousProperties $properties
                            }
                        }
                    }
                }
            }
        }
    }
}

[hashtable]$params = @{ 'ErrorAction' = 'SilentlyContinue' }
[hashtable]$alreadyFetched = @{}

if( $PSBoundParameters[ 'username' ] )
{
    if( ! $PSBoundParameters[ 'password' ] )
    {
        $password = $env:randomkey
    }

    if( ! [string]::IsNullOrEmpty( $password ) )
    {
        $credential = New-Object System.Management.Automation.PSCredential( $username , ( ConvertTo-SecureString -AsPlainText -String $password -Force ) )
    }
    else
    {
        Throw "Must specify password when using -username either via -password or %RandomKey%"
    }
}

if( $credential )
{
    $params.Add( 'Credential' , $credential )
}
 else
{
    $params.Add( 'UseDefaultCredentials' , $true )
}

## used to try and figure out the highest supported oData version but proved problematic
[int]$highestVersion = $oDataVersion ## if( $oDataVersion -le 0 ) { 10 } else { -1 }
$fatalException = $null
[int]$version = $oDataVersion

$startDate = $null

if( ! [string]::IsNullOrEmpty( $last ) )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Write-Error "Unknown multiplier `"$($last[-1])`"" ; Exit }
    }
    $endDate = [datetime]::Now
    if( $last.Length -le 1 )
    {
        $from = $endDate.AddSeconds( -$multiplier )
    }
    else
    {
        $from = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [int] ) * $multiplier ) )
    }
}

$services = $null
## queries are case sensitive so help people who don't know this but don't do it for everything as would break items like DesktopGroups
if( $query -cmatch '^[a-z]' )
{
    $TextInfo = (Get-Culture).TextInfo
    $query = $TextInfo.ToTitleCase( $query ).ToString()
}

if( $PsCmdlet.ParameterSetName -eq 'cloud' )
{
    if( ! $PSBoundParameters[ 'authtoken' ] )
    {
        Add-PSSnapin -Name Citrix.Sdk.Proxy.*
        if( ! ( Get-Command -Name Get-XDAuthentication -ErrorAction SilentlyContinue ) )
        {
            Throw "Unable to find the Get-XDAuthentication cmdlet - is the Virtual Apps and Desktops Remote PowerShell SDK installed ?"
        }
        Get-XDAuthentication -CustomerId $customerid
        if( ! $? )
        {
            Throw "Failed to get authentication token for Cloud customer id $customerid"
        }
        $authtoken = $GLOBAL:XDAuthToken
    }
    $params.Add( 'Headers' , @{ 'Customer' = $customerid ; 'Authorization' = $authtoken } )
    $protocol = 'https'
}

[bool]$cloud = $false

[array]$data = @( do
{
    if( $oDataVersion -le 0 )
    {
        ## Figure out what the latest OData version supported is. Could get via remoting but remoting may not be enabled
        if( $highestVersion -le 0 )
        {
            break
        }
        $version = $highestVersion--
    }
    
    if( $PsCmdlet.ParameterSetName -eq 'cloud' )
    {
        $params[ 'Uri' ] = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $customerid , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
        $cloud = $true
    }
    else
    {
        $params[ 'Uri' ] = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $ddc , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
    }

    Write-Verbose "URL : $($params.Uri)"

    try
    {
        if( ! $PSBoundParameters[ 'query' ] )
        {
            $services = Invoke-RestMethod @params
        }
        else
        {
            Invoke-RestMethod @params | Invoke-ODataTransform
        }
        $fatalException = $null
        break ## since call succeeded so that we don't report for lower versions
    }
    catch
    {
        $fatalException = $_
    }
} while ( $highestVersion -gt 0 ) )

if( $fatalException )
{
    Throw $fatalException
}

if( $services )
{
    if( $services.PSObject.Properties[ 'service' ] )
    {
        $services.service.workspace.collection | Select-Object -Property 'title' | Sort-Object -Property 'title'
    }
    else
    {
        $services.value | Sort-Object -Property 'name'
    }
}
elseif( $data -and $data.Count )
{
    if( $PSBoundParameters[ 'join' ] )
    {
        [hashtable]$tables = @{}

        ## now figure out what other tables we need in order to satisfy these ids (not interested in id on it's own)
        $data[0].PSObject.Properties | Resolve-CrossReferences -cloud:$cloud

        [int]$originalPropertyCount = $data[0].PSObject.Properties.GetEnumerator()|Measure-Object |Select-Object -ExpandProperty Count
        [int]$finalPropertyCount = -1

        ## now we need to add these cross referenced items
        ForEach( $datum in $data )
        {
            $datum.PSObject.Properties | Where-Object { $_.Name -ne 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and ! [string]::IsNullOrEmpty( $Matches[1] ) } | ForEach-Object `
            {
                $property = $_
                Resolve-NestedProperties $property | ForEach-Object `
                {
                    $_.PSObject.Properties | Where-Object MemberType -eq 'NoteProperty' | ForEach-Object `
                    {
                        Add-Member -InputObject $datum -MemberType NoteProperty -Name $_.Name -Value $_.Value
                    }
                }
            }

            if( $finalPropertyCount -lt 0 )
            {
                $finalPropertyCount = $datum.PSObject.Properties.GetEnumerator()|Measure-Object |Select-Object -ExpandProperty Count
                Write-Verbose -Message "Expanded from $originalPropertyCount properties to $finalPropertyCount"
            }

            $datum
        }
    }
}
else
{
    Write-Warning "No data returned"
}
