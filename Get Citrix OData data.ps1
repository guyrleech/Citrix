#requires -version 3

<#
    Get Director data from a Delivery controller or Citrix Cloud

    Modification History:

    @guyrleech  04/03/20  Added Citrix Cloud capability
    @guyrleech  05/03/20  Made -join run recursive. Fixed bug with date ranges
    @guyrleech  30/03/22  Added -AllowUnencryptedAuthentication for pwsh 7 Invoke-RestMethod
    @guyrleech  25/04/22  Fixed issue where only first 100 items being returned because of Citrix API changes
    @guyrleech  07/06/22  Added -profilename for Citrix Cloud
    @guyrleech  10/10/22  Added progress indicator
    @guyrleech  15/06/23  Added output formats, exclusion of ids, output file wttih %variables%
    @guyrleech  20/06/23  Fixed issues when running against XD 7.6
#>

<#
.SYNOPSIS

Send queries to a Citrix Delivery Controller or Citrix Cloud and present the results back as PowerShell objects

.PARAMETER ddc

The Delivery Controller to query

.PARAMETER customerId

The Citrix Cloud customerid to query

.PARAMETER profilename

The name of the Citrix Cloud credentials profile (as returned by Get-XDCredentials -ListProfiles

.PARAMETER authToken

The Citrix cloud authentication token to use. If not specified will prompt for credentials

.PARAMETER AllowUnencryptedAuthentication

PowerShell 7.x errors with "The cmdlet cannot protect plain text secrets sent over unencrypted connections" if Invoke-RestMethod is called via http so specify this to override the behaviour

.PARAMETER outputfile

Name and path to write output to. If it exists, use -overwrite to overwrite.
Pseudo environment variables like %day% and %month% can be used in the folder and/or file name and will be created as necessary

.PARAMETER overwrite

If specified as "yes", any existing output file will be overwritten otherwise the script will fail if the outoput file already exists

.PARAMETER format

The output format to use. If not specified it will be determined from the output file extension

.PARAMETER outputEncoding

The output encoding to use

.PARAMETER credential

A credential object to use to connect to the specific Delivery Controller

.PARAMETER noids

Do not include any id in the output.
Use with -join to resolve ids to entity names

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

.PARAMETER progressEveryPercent

Show progress at every this percentage of completion when joining rows

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

'.\Get Citrix OData data.ps1' -ddc ctxddc01 -query connections -join yes -last 7d

Send a web request to the Delivery Controller ctxddc01 to retrieve the list of all connections created within the last 7 days and cross reference any id's returned

.EXAMPLE

'.\Get Citrix OData data.ps1' -ddc ctxddc01 -query connections -join yes -last 7d -outputFile c:\logs\%year%\%monthname%\citrix.odata.%hour.%minute%.%second%.csv -noids yes

Send a web request to the Delivery Controller ctxddc01 to retrieve the list of all connections created within the last 7 days and cross reference any id's returned
but do not include the ids in the csv output which is written to a file in the c:\logs folder with any missing path elements being created as required

.EXAMPLE

'.\Get Citrix OData data.ps1' -customerid yourcloudid -query connections -join yes -last 7d

Prompt for credentials for the Citrix Cloud customer with id yourcloudid, send a web request to the Citrix Cloud to retrieve the list of all connections created within the last 7 days and cross reference any id's returned

.NOTES

https://developer-docs.citrix.com/projects/monitor-service-odata-api/en/latest/

If an auth token is not passed, the Citrix Remote PowerShell SDK must be available in order to get an auth token - https://www.citrix.com/downloads/citrix-cloud/product-software/xenapp-and-xendesktop-service.html

#>

Param
(
    [Parameter(ParameterSetName='ddc',Mandatory=$true)]
    [string]$ddc ,
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$customerid ,
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$profileName ,
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$authtoken ,
    [string]$query ,
    [ValidateSet('Yes','No')]
    [string]$join = 'no' ,
    [ValidateSet('Yes','No')]
    [string]$noId = 'no' ,
    [datetime]$from ,
    [datetime]$to ,
    [string]$last ,
    [ValidateSet('csv','json','xml','object','txt')]
    [string]$format = 'csv' ,
    [ValidateSet( 'String' , 'Unicode' , 'Byte' , 'BigEndianUnicode' , 'UTF8' , 'UTF7' , 'UTF32' , 'Ascii' , 'Default' , 'Oem' , 'BigEndianUTF32' )]
    [string]$outputEncoding = 'UTF8' ,
    [ValidateSet('Yes','No')]
    [string]$overWrite = 'No' ,
    [string]$includePropertyRegex ,
    [string]$excludePropertyRegex ,
    [string]$csvOutputDelimiter = ',' , ## Dutch use semicolon!
    [string]$outputFile , 
    [System.Management.Automation.PSCredential]$credential ,
    [string]$username , 
    [string]$password ,
    [switch]$noQueryCaseChange ,
    [Parameter(ParameterSetName='ddc',Mandatory=$false)]
    [ValidateSet('http','https')]
    [string]$protocol = 'http' ,
    [string]$baseCloudURL =  'https://api-us.cloud.com/monitorodata' ,
    [int]$oDataVersion = 4 ,
    [int]$retryMilliseconds = 1000 ,
    [int]$progressEveryPercent = 10 ,
    [switch]$AllowUnencryptedAuthentication
)

if( $PSBoundParameters[ 'from' ] -and $PSBoundParameters[ 'to' ] -and $from -gt $to )
{
    Throw "Start date $(Get-Date -Date $from -Format G) is after end date $(Get-Date -Date $to -Format G)"
}

## map tables to the date stamp we will filter on
[hashtable]$dateFields = @{
     'Session' = 'StartDate'
     'Connection' = 'BrokeringDate'
     'ConnectionFailureLog' = 'FailureDate'
}
#region Functions

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
        if( $records -is [array] -or $records -is [Xml.XmlElement] )
        {
            if( -Not $propertyNames )
            {
                $properties = ($records | Select-Object -First 1).content.properties
                if( $properties )
                {
                    $propertyNames = $properties | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                }
                else
                {
                    // v4+
                    $propertyNames = 'NA' -as [string]
                }
            }
            if( $propertyNames -is [string] )
            {
                $records | Select-Object -ExpandProperty value
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
    if( -Not $field )
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
            "()?`$filter=$field ge $($from.ToString( 's' ))Z"
        }
        if( $to )
        {
            "and $field le $($to.ToString('s'))Z"
        }
    }
    else
    {
        if( $from )
        {
            "()?`$filter=$field ge datetime'$($from.ToString( 's' ))'"
        }
        if( $to )
        {
            "and $field le datetime'$($to.ToString( 's' ))'"
        }
    }
}

Function Resolve-CrossReferences
{
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        $properties ,
        $include ,
        $exclude ,
        [switch]$cloud
    )
    
    Process
    {
        $properties.Where( { ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(SessionKey)$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } ) | Select-Object -Property Name | . { Process `
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
            
            if( -not [string]::IsNullOrEmpty( $include ) )
            {
                if( $id -notmatch $include )
                {
                    try
                    {
                        $alreadyFetched.Add( $id , $id )
                    }
                    catch
                    {
                    }
                }
                ## else included
            }
            elseif( -not [string]::IsNullOrEmpty( $exclude ) )
            {
                if( $id -match $exclude )
                {
                    try
                    {
                        $alreadyFetched.Add( $id , $id )
                    }
                    catch
                    {
                    }
                }
                ## else not excluded
            }

            if( -Not $tables[ $id ] -and -Not $alreadyFetched[ $id ] )
            {
                if( $cloud )
                {
                    $params[ 'Uri' ] = "$baseCloudURL/$($id)s" ## + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
     
                    ##$params.uri = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}s" -f $protocol , $customerid , $version ,  $id ) ## + (Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion)
                }
                else
                {
                    $params.uri = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}s" -f $protocol , $ddc , $version , $id ) ## + (Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion)
                }

                ## save looking up again, especially if it errors as we are not looking up anything valid
                $alreadyFetched.Add( $id , $id )

                [hashtable]$table = @{}
                [string]$lasturi = $params.uri

                try
                {
                    ## have to deal with Citrix returning data in batches of 100 (or whatever they choose)
                    $queryResults = New-Object -TypeName System.Collections.Generic.List[object]

                    do
                    {
                        $resultsPage = $null
                        
                        try
                        {
                            $resultsPage = Invoke-RestMethod @params

                            if( $null -ne $resultsPage )
                            {
                                $queryResults += $resultsPage

                                ## https://support.citrix.com/article/CTX312284
                                if( $resultsPage.PSObject.Properties['@odata.nextLink' ] -and -not [string]::IsNullOrEmpty( $resultsPage.'@odata.nextLink' ) )
                                {
                                    $params.uri = $resultsPage.'@odata.nextLink'
                                    ## prevent infinite loop if something goes wrong
                                    if( $params.uri -ne $lasturi )
                                    {
                                        Write-Verbose -Message "More data available, fetching from $($params.uri)"
                                        $lasturi = $params.uri
                                    }
                                    else
                                    {
                                        Write-Warning -Message "Next link $lasturi is the same as the previous one so aborting loop"
                                        break
                                    }
                                }
                                else ## no further results available so quit loop
                                {
                                    break
                                }
                            }
                        }
                        catch
                        {
                            $fatalException = $_
                            Write-Verbose -Message "Resolve-CrossReferences exception: $($params.uri) : $fatalException"

                            if( $cloud )
                            {
                                if( $fatalException.Exception.Response.StatusCode -eq 429 ) ##  Too Many Requests
                                {
                                    Write-Verbose -Message "$(Get-Date -Format G) : too many requests error so will retry after $($retryMilliseconds)ms"
                                    Start-Sleep -Milliseconds $retryMilliseconds
                                    $resultsPage = 'Try again' ## just causes do while not to exit
                                }
                                ## else might be accidental bad request so ignore but bail out of loop since little point repeating   
                            }
                        }
                    } while( $resultsPage )

                    $queryResults | Invoke-ODataTransform | . { Process `
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
                            if( $property.MemberType -ieq 'NoteProperty' -and $property.Name -ine $keyName -and $property.Name -ine 'sid' -and $property.Name -match '(.*)Id$' )
                            {
                                $property | Resolve-CrossReferences -cloud:$cloud -include $include -exclude $exclude
                            }
                        }
                    }}
                    if( $table.Count )
                    {
                        Write-Verbose -Message "Adding table $id with $($table.Count) entries"
                        $tables.Add( $id , $table )
                    }
                }
                catch
                {
                }
            }
        }
    }}
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
        ## $properties | Where-Object { $_.Name -ne 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and ! [string]::IsNullOrEmpty( $Matches[1] ) } | ForEach-Object `
        $properties.Where( { $_.Name -ine 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } ).ForEach( `
        {
            $property = $_
                
            if( -Not [string]::IsNullOrEmpty( ( $id = ( $Matches[1] -replace '^Current' )) ))
            {
                if ( $table = $tables[ $id ] )
                {
                    if( $property.Value -and ( $item = $table[ ($property.Value -as [string]) ]))
                    {
                        $datum.PSObject.properties.remove( $property )
                        $item.PSObject.Properties | ForEach-Object `
                        {
                            [pscustomobject]@{ "$id.$($_.Name)" = $_.Value }
                            if( $_.Name -ine $property.Name -and ( -Not $previousProperties -or -Not ( $previousProperties.Where( { $_.Name -eq $_.Name } )))) ## don't lookup self or a key if it was one we previously looked up
                            {
                                Resolve-NestedProperties -properties $_ -previousProperties $properties
                            }
                        }
                    }
                }
            }
        })
    }
}

Function Out-PassThru
{
    Process
    {
        $_
    }
}

#endregion Functions

if( -Not [string]::IsNullOrEmpty( $outputFile ) )
{
    ## format not specified so get from output file extension
    if( -not $PSBoundParameters[ 'format' ] )
    {
        try
        {
            $format = $outputFile -replace '^.*\.(\w+)$' , '$1'
        }
        catch
        {
            Throw "Cannot determine a supported output format from output file extension on $outputFile"
        }
    }

    if( $outputFile.IndexOf( '%' ) -ne $outputFile.LastIndexOf( '%' ) )
    {
        $now = [datetime]::Now
        $outputFile = $outputFile -replace '%year%' , $now.ToString( 'yyyy' ) -replace '%month%' , $now.ToString( 'MM' ) -replace '%day%' , $now.ToString( 'dd') -replace '%monthname%' , $now.ToString( 'MMMM' ) -replace '%dayname%' , $now.ToString( 'dddd') `
            -replace '%hours?%' , $now.ToString( 'HH')  -replace'%minutes?%' , $now.ToString( 'mm') -replace '%seconds?%' , $now.ToString( 'ss') -replace '%query%' , $query
        [string]$logFolder = Split-Path -Path $outputFile -Parent
        if( -Not ( Test-Path -Path $logFolder -PathType Container ))
        {
            if( -Not( New-Item -Path $logFolder -ItemType Directory -Force ) )
            {
                Write-Warning -Message "Failed to create log folder $logFolder"
            }
        }
    }
    
    if( (Test-Path -Path $outputFile) -and $overwrite -ine 'yes' )
    {
        Throw "Cannot proceeed as output file `"$outputFile`" already exists and -overwrite not used"
    }
}

[hashtable]$outputProcessors = @{
    'csv' =    @{ Command = 'ConvertTo-csv'  ; Arguments = @{ 'NoTypeInformation' = $true ; 'Delimiter' = $csvOutputDelimiter } }
    'json' =   @{ Command = 'ConvertTo-Json' ; Arguments = @{ 'Depth' = 10 } }
    'xml' =    @{ Command = 'ConvertTo-XML'  ; Arguments = @{ 'Depth' = 10 } }
    'txt' =    @{ Command = 'Out-String'     ; Arguments = @{ } }
    'object' = @{ Command = 'Out-PassThru'   ; Arguments = @{ } }
}

if( $outputProcessor = $outputProcessors[ $format ] )
{
    $outputCommand = $outputProcessor.Command
    $outputArguments = $outputProcessor.Arguments
}
else
{
    Throw "Unsupported output format $format"
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

if( $AllowUnencryptedAuthentication )
{
    if( $PSVersionTable.PSVersion.Major -le 5 )
    {
        Write-Warning -Message "-AllowUnencryptedAuthentication not supported but also not required for this version of PowerShell"
    }
    else
    {
        $params.Add( 'AllowUnencryptedAuthentication' , $true )
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
if( $query -cmatch '^[a-z]' -and -Not $noQueryCaseChange )
{
    $TextInfo = (Get-Culture).TextInfo
    $query = $TextInfo.ToTitleCase( $query ).ToString()
    ## TODO need to ensure $ keywords are lower case
    ##$query = $query -replace '$
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
        if( $PSBoundParameters[ 'profileName' ] )
        {
            Get-XDAuthentication -ProfileName $profileName
            if( [string]::IsNullOrEmpty( $customerid ) )
            {
                $customerid = (Get-XDCredentials -ProfileName $profileName).Credentials.CustomerId
                if( [string]::IsNullOrEmpty( $customerid ) )
                {
                    Throw "Failed to get customer id from profile $profileName"
                }
            }
        }
        else
        {
            Get-XDAuthentication -CustomerId $customerid
        }
        if( ! $? )
        {
            Throw "Failed to get authentication token for Cloud customer id $customerid"
        }
        $authtoken = $GLOBAL:XDAuthToken
    }
    $params.Add( 'Headers' , @{ 'Citrix-CustomerId' = $customerid ; 'Authorization' = $authtoken } )
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
        ##$params[ 'Uri' ] = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $customerid , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
        $params[ 'Uri' ] = "$baseCloudURL/$query" + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
        $cloud = $true
    }
    else
    {
        $params[ 'Uri' ] = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $ddc , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $version)
    }

    Write-Verbose "URL : $($params.Uri)"

    try
    {
        [int]$results = 0
        [int]$requests = 0
        [string]$lasturi = $params.uri

        do
        {
            $requests++
            $resultsPage = $null
            $resultsPage = Invoke-RestMethod @params

            if( $null -ne $resultsPage )
            {
                $results += ( $resultsPage | Select-Object -ExpandProperty Value | Measure-Object).Count
                if( [string]::IsNullOrEmpty( $query ) )
                {
                    $resultsPage
                }
                else
                {
                    $resultsPage | Invoke-ODataTransform
                }
                ## https://support.citrix.com/article/CTX312284
                if( $resultsPage.PSObject.Properties['@odata.nextLink' ] -and -not [string]::IsNullOrEmpty( $resultsPage.'@odata.nextLink' ) )
                {
                    $params.uri = $resultsPage.'@odata.nextLink'
                    ## prevent infinite loop if something goes wrong
                    if( $params.uri -ne $lasturi )
                    {
                        Write-Verbose -Message "More data available, fetching from $($params.uri)"
                        $lasturi = $params.uri
                    }
                    else
                    {
                        Write-Warning -Message "Next link $lasturi is the same as the previous one so aborting loop"
                        break
                    }
                }
                else ## no further results available so quit loop
                {
                    break
                }
            }
        } while( $resultsPage )
        Write-Verbose -Message "Got $results query results in total across $requests requests"
        
        $fatalException = $null
        break ## since call(s) succeeded so that we don't report for lower versions
    }
    catch
    {
        $fatalException = $_
        if( $cloud )
        {
            if( $fatalException.Exception.Response.StatusCode -eq 429 ) ##  Too Many Requests
            {
                Write-Verbose -Message "$(Get-Date -Format G) : too many requests error so will retry after $($retryMilliseconds)ms"
                Start-Sleep -Milliseconds $retryMilliseconds
            }
            else ## something unrecoverable so exit loop
            {
                $highestVersion = -1
            }
        }
        else
        {
            $version = --$highestVersion
        }
    }
} while ( $highestVersion -gt 0 ) )

if( $fatalException )
{
    Throw $fatalException
}

if( [string]::IsNullOrEmpty( $query ) )
{
    $services = $data
}

if( $services )
{
    if( $services.PSObject.Properties[ 'service' ] -or ( $services | Get-Member -MemberType Property -Name service -ErrorAction SilentlyContinue ))
    {
        $services.service.workspace.collection | Select-Object -Property 'title' | Sort-Object -Property 'title'
    }
    else
    {
        $services | Select-Object -expandproperty 'value' | Sort-Object -Property 'name'
    }
}
elseif( $data -and $data.Count )
{
    [array]$results = @( if( $join -ieq 'yes' )
    {
        [string]$activity = "Joining $($data.Count) result rows"
        Write-Verbose -Message "$(Get-Date -Format G): $activity"
        [hashtable]$tables = @{}

        ## now figure out what other tables we need in order to satisfy these ids (not interested in id on it's own)
        $data[0].PSObject.Properties | Resolve-CrossReferences -cloud:$cloud -Include $includePropertyRegex -Exclude $excludePropertyRegex 

        [int]$originalPropertyCount = $data[0].PSObject.Properties.GetEnumerator() | Measure-Object | Select-Object -ExpandProperty Count
        [int]$finalPropertyCount = -1
        [int]$counter = 0
        [int]$lastPercentCompete = -1

        ## now we need to add these cross referenced items
        ForEach( $datum in $data )
        {
            $counter++
            if( $progressEveryPercent -gt 0 )
            {
                [int]$percentComplete = ($counter / $data.Count) * 100
                if( $percentComplete -ne $lastPercentCompete -and $percentComplete % $progressEveryPercent -eq 0 )
                {
                    Write-Progress -Activity $activity -Status "$counter processed" -PercentComplete $percentComplete
                }
                $lastPercentCompete = $percentComplete ## avoid updating very next item because of rounding calculation
            }

            $datum.PSObject.Properties.Where( { $_.Name -ne 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } ) | . { Process `
            {
                $property = $_
                Resolve-NestedProperties -properties $property | . { Process `
                {
                    $_.PSObject.Properties.Where( { $_.MemberType -eq 'NoteProperty' } ) | . { Process `
                    {
                        if( $noid -ieq 'no' -or $_.Name -notmatch 'id$' )
                        {
                            Add-Member -InputObject $datum -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
                        }
                    }}
                }}
            }}

            if( $finalPropertyCount -lt 0 )
            {
                $finalPropertyCount = $datum.PSObject.Properties.GetEnumerator() | Measure-Object | Select-Object -ExpandProperty Count
                Write-Verbose -Message "Expanded from $originalPropertyCount properties to $finalPropertyCount"
            }

            $datum
        }
        
        if( $progressEveryPercent -gt 0 )
        {
            Write-Progress -Completed -Activity $activity
        }
    }
    else
    {
        $data
    })

    Write-Verbose -Message "Got $($results.Count) results"

    
    $finalOutput = $results | . $outputCommand @outputArguments ## output to stdout

    if( [string]::IsNullOrEmpty( $outputFile ) )
    {
        $finalOutput
    }
    else
    {
        $written = $null
        $finalOutput | Set-Content -Path $outputFile -Encoding $outputEncoding
        if( $? )
        {
            Write-Verbose -Message "Wrote $($finalOutput.Length) items to `"$outputFile`""
        }
    }
}
else
{
    Write-Warning "No data returned"
}

# SIG # Begin signature block
# MIIjcAYJKoZIhvcNAQcCoIIjYTCCI10CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUx0AmkkDAz8jBPUla0J5GYh4m
# yvuggh2OMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJ
# KoZIhvcNAQEMBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQg
# QXNzdXJlZCBJRCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1
# OVowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBS
# b290IEc0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+Rd
# SjwwIjBpM+zCpyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20d
# q7J58soR0uRf1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7f
# gvMHhOZ0O21x4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRA
# X7F6Zu53yEioZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raR
# mECQecN4x7axxLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzU
# vK4bA3VdeGbZOjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2
# mHY9WV1CdoeJl2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkr
# fsCUtNJhbesz2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaA
# sPvoZKYz0YkH4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxf
# jT/JvNNBERJb5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEe
# xcCPorF+CiaZ9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQF
# MAMBAf8wHQYDVR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaA
# FEXroq/0ksuCMS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcB
# AQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggr
# BgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEUm9vdENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQK
# MAgwBgYEVR0gADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3
# v1cHvZqsoYcs7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy
# 3iS8UgPITtAq3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cn
# RNTnf+hZqPC/Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3
# WlxUjG/voVA9/HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2
# zm8jLfR+cWojayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCC
# Bq4wggSWoAMCAQICEAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0
# MB4XDTIyMDMyMzAwMDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVz
# dGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD
# 0Z5Mom2gsMyD+Vr2EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39
# Q7SE2hHxc7Gz7iuAhIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decf
# BmWNlCnT2exp39mQh0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RU
# CyFobjchu0CsX7LeSn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+x
# tVhNef7Xj3OTrCw54qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OA
# e3VuJyWQmDo4EbP29p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRA
# KKtzQ87fSqEcazjFKfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++b
# Pf4OuGQq+nUoJEHtQr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+
# OcD5sjClTNfpmEpYPtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2Tj
# Y+Cm4T72wnSyPx4JduyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZ
# DNIztM2xAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQW
# BBS6FtltTYUvcyl2mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/
# 57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYI
# KwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9j
# cmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1Ud
# IAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEA
# fVmOwJO2b5ipRCIBfmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnB
# zx0H6T5gyNgL5Vxb122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXO
# lWk/R3f7cnQU1/+rT4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBw
# CnzvqLx1T7pa96kQsl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q
# 6/aesXmZgaNWhqsKRcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJ
# uXdmkfFynOlLAlKnN36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEh
# QNC3EyTN3B14OuSereU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo4
# 6Zzh3SP9HSjTx/no8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3
# v5gA3yAWTyf7YGcWoWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHz
# V9m8BPqC3jLfBInwAM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZV
# VCsfgPrA8g4r5db7qS9EFUrnEw4d2zc4GqEr9u3WfPwwggbAMIIEqKADAgECAhAM
# TWlyS5T6PCpKPSkHgD1aMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcw
# FQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3Rl
# ZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjIwOTIxMDAw
# MDAwWhcNMzMxMTIxMjM1OTU5WjBGMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGln
# aUNlcnQxJDAiBgNVBAMTG0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDIyIC0gMjCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAM/spSY6xqnya7uNwQ2a26HoFIV0
# MxomrNAcVR4eNm28klUMYfSdCXc9FZYIL2tkpP0GgxbXkZI4HDEClvtysZc6Va8z
# 7GGK6aYo25BjXL2JU+A6LYyHQq4mpOS7eHi5ehbhVsbAumRTuyoW51BIu4hpDIjG
# 8b7gL307scpTjUCDHufLckkoHkyAHoVW54Xt8mG8qjoHffarbuVm3eJc9S/tjdRN
# lYRo44DLannR0hCRRinrPibytIzNTLlmyLuqUDgN5YyUXRlav/V7QG5vFqianJVH
# hoV5PgxeZowaCiS+nKrSnLb3T254xCg/oxwPUAY3ugjZNaa1Htp4WB056PhMkRCW
# fk3h3cKtpX74LRsf7CtGGKMZ9jn39cFPcS6JAxGiS7uYv/pP5Hs27wZE5FX/Nurl
# fDHn88JSxOYWe1p+pSVz28BqmSEtY+VZ9U0vkB8nt9KrFOU4ZodRCGv7U0M50GT6
# Vs/g9ArmFG1keLuY/ZTDcyHzL8IuINeBrNPxB9ThvdldS24xlCmL5kGkZZTAWOXl
# LimQprdhZPrZIGwYUWC6poEPCSVT8b876asHDmoHOWIZydaFfxPZjXnPYsXs4Xu5
# zGcTB5rBeO3GiMiwbjJ5xwtZg43G7vUsfHuOy2SJ8bHEuOdTXl9V0n0ZKVkDTvpd
# 6kVzHIR+187i1Dp3AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0T
# AQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+e
# yG8wHQYDVR0OBBYEFGKK3tBh/I8xFO2XC809KpQU31KcMFoGA1UdHwRTMFEwT6BN
# oEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJT
# QTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGA
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUH
# MAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQAD
# ggIBAFWqKhrzRvN4Vzcw/HXjT9aFI/H8+ZU5myXm93KKmMN31GT8Ffs2wklRLHiI
# Y1UJRjkA/GnUypsp+6M/wMkAmxMdsJiJ3HjyzXyFzVOdr2LiYWajFCpFh0qYQitQ
# /Bu1nggwCfrkLdcJiXn5CeaIzn0buGqim8FTYAnoo7id160fHLjsmEHw9g6A++T/
# 350Qp+sAul9Kjxo6UrTqvwlJFTU2WZoPVNKyG39+XgmtdlSKdG3K0gVnK3br/5iy
# JpU4GYhEFOUKWaJr5yI+RCHSPxzAm+18SLLYkgyRTzxmlK9dAlPrnuKe5NMfhgFk
# nADC6Vp0dQ094XmIvxwBl8kZI4DXNlpflhaxYwzGRkA7zl011Fk+Q5oYrsPJy8P7
# mxNfarXH4PMFw1nfJ2Ir3kHJU7n/NBBn9iYymHv+XEKUgZSCnawKi8ZLFUrTmJBF
# YDOA4CPe+AOk9kVH5c64A0JH6EE2cXet/aLol3ROLtoeHYxayB6a1cLwxiKoT5u9
# 2ByaUcQvmvZfpyeXupYuhVfAYOd4Vn9q78KVmksRAsiCnMkaBXy6cbVOepls9Oie
# 1FqYyJ+/jbsYXEP10Cro4mLueATbvdH7WwqocH7wl4R44wgDXUcsY6glOJcB0j86
# 2uXl9uab3H4szP8XTE0AotjWAQ64i+7m4HJViSwnGWH2dwGMMYIFTDCCBUgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQUZz1zn3w0/njUXVQOeX3Oz7dzOR0wDQYJKoZIhvcNAQEBBQAEggEA
# qFy30ov03XVp8lgPu9IT2/ZAoGk5adebeRGifEmyHFge39W4ApGcSGn/s+/88sYN
# R8bDl6KeXXXlV9+5SaeUA9LecwWLVDIJ3KT73Xpy3G4on2+GNkFPcVtq+1cNszem
# QmX/bY0rx21YRgveL+ynLpsCOaiFJ9AXNBwVbiNZvcDU1Us2mlPZpUw+Ysuw5T6D
# CVYsAbs4yfgN5DR63sHBaOwCEAdeKo9GylMfQZ8mgmjhyiDjWxzdg9jk1VZFHoBe
# KLXDdCsWRizdawah/DFMaw70idejBtP8AnGAjCReDLyKQSIFpturB9uSRgirpfzQ
# ax0NMWVWGPuedAYxiUxuQ6GCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3
# MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UE
# AxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBp
# bmcgQ0ECEAxNaXJLlPo8Kko9KQeAPVowDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMzA2MjAyMDA5MjFa
# MC8GCSqGSIb3DQEJBDEiBCB4XQB82VVy5ndg1ODynVF+9F9l0HQy7cSL76N33my8
# 7TANBgkqhkiG9w0BAQEFAASCAgA1KsO3eNX+s1FgTarVbPBp2TBh1cjFsOQPmWnz
# 0l2lvmKFOBBuh/9TqVsV5btfMy5HYh/UiPHkGUmsVvqlgTUdyTXh5lZc4Uwkp1uo
# V+xaq6JiUmnn6R1wRawW2w8cyMtGdel2OrqNwzcDq4BW+vi9lkt0Gt0mAJozYUE9
# JySLdmFuXVFnGTr1CwVYuuH2ZjQQ+ROqLPPGbDmgLVPWNYd9geihGbX/gbu875Uy
# f6svo3y7H6v8OKyfhNeTFIvqTXa9ZT1xIcOUUP+GfHwsiHrRqetotNa+g5MTp3b5
# I//Ci+B/E/TwBmJ8ZD50RU57GtolJN/UsJcxWuwPXOjjf2kIPArUb31YukThtyXt
# rQMAMXlw33zJ9W2pqBH4+QpfGcgXdKbPUkUGRtdlbeQj/1iI71JRS6MHIjik27DJ
# Z3TAquZUA+Uxj6IhHaFu8h8X3x0bCBaEGvedk9w4jeFwGcg7Sa7BlNn6LPaHEUBV
# P+sP+KqRrgHxQc/7vhzqYabuPLpfS1UIXf7NS7NeE/zmcqeBKVARyB1WWyIjRz+p
# k5yThJIPtyIl21tKYgZkJtrARSC33NexWMphsXWYD427bDbSR/BvWCSLWTtd5SH3
# 3sIF/G02FQLacX7jxG7LK7vNTE1LnojYISyML4CCpvFh4w8sSTEPQaSlbf9MaEA+
# tkBIxQ==
# SIG # End signature block
