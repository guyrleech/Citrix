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
    @guyrleech  08/08/23  Fixed issue where didn't run with PS 4.0, add date filtering when cross referencing Sessions and Connections
    @guyrleech  16/04/24  Added Cloud auth via client id & secret arguments so no requirement for Citrix Remote PS SDK unless using auth profiles
    @guyrleech  17/04/24  Added count query to first request and -maximumItems to limit items returned
    @guyrleech  22/04/24  Fix for count query when no query string passed
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

.PARAMETER maximumItems

The maximum number of items to return although as results are returned in pages, it may be slightly more than this number

.PARAMETER join

Where there are id's in the retrieved query, look up the objects for the id's and substitute them in the output

.PARAMETER last

Only retrieve items such as sessions or connections which have been created in the last specified period

.PARAMETER username

The username to use when querying the Delivery Controller. If not specified the user running the script will be used

.PARAMETER password

The password for the account used to query the Delivery Controller. If the %RandomKey% environment variable is set, its contents will be used as the password

.PARAMETER clientId

The client id to be used to authenticate to Citrix DaaS (usually a GUID)

.PARAMETER clientSecret

The client secret for the client id specified by -clientId to authenticate to Citrix DaaS

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
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$clientId ,
    [Parameter(ParameterSetName='cloud',Mandatory=$false)]
    [string]$clientSecret ,
    [string]$query ,
    [ValidateSet('Yes','No')]
    [string]$join = 'no' ,
    [ValidateSet('Yes','No')]
    [string]$noId = 'no' ,
    [datetime]$from ,
    [datetime]$to ,
    [string]$last ,
    [int]$maximumItems = 0 ,
    [string]$proxyServer = $null ,
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

Function Get-BearerToken {
    ## https://www.mycugc.org/blogs/eltjo-van-gulik/2019/01/16/blog-monitoring-citrix-cloud-with-odata-and-powers
    ## https://developer.cloud.com/citrix-cloud/citrix-cloud-api-overview/docs/get-started-with-citrix-cloud-apis#bearer_token_tab_oauth_2.0_flow
    param (
        [Parameter(Mandatory=$true)][string]
        $clientId,
        [Parameter(Mandatory=$true)][string]
        $clientSecret
    )
    [string]$bearerToken = $null
    [hashtable]$body = @{
        'grant_type' = 'client_credentials'
        'client_id' = $clientId
        'client_secret' = $clientSecret
    }
    
    $response = $null
    try
    {
        $startRequestTime = [datetime]::Now
        Write-Verbose -Message "$($startRequestTime.ToString('G')): sending auth request"
        $response = Invoke-RestMethod -Uri 'https://api-us.cloud.com/cctrustoauth2/root/tokens/clients' -Method POST -Body $body
        $endRequestTime = [datetime]::Now
    }
    catch
    {
        $endRequestTime = [datetime]::Now
        Write-Verbose -Message "$($startRequestTime.ToString('G')): get-BearerToken: exception $_"
        if( $_.Exception.Message -imatch 'Unable to connect to the remote server' -or  $_.Exception.Message -imatch 'The operation has timed out' -and -not [string]::IsNullOrEmpty( $proxyServer )  )
        {
            $response = Invoke-RestMethod -Uri 'https://api-us.cloud.com/cctrustoauth2/root/tokens/clients' -Method POST -Body $body -Proxy $proxyServer -ProxyUseDefaultCredentials
        }
        else
        {
            Throw $_
        }
    }

    if( $null -ne $response )
    {
        $bearerToken = "CwsAuth Bearer=$($response | Select-Object -expandproperty access_token)"
        Write-Verbose -Message "$($startRequestTime.ToString('G')): auth seems ok"
    }
    ## else will have output error
    $bearerToken ## return    
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
        [switch]$cloud ,
        $from , 
        $to
    )
    
    Process
    {
        $properties | Where-Object -FilterScript { ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(SessionKey)$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } | Select-Object -Property Name | . { Process `
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
                [string]$dateFilter= $null
                if( $id -ieq 'Session' -or $id -ieq 'Connection' ) ## otherwise can take ages if it gets these from beginning of time
                {
                    $dateFilter = Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion
                }
                if( $cloud )
                {
                    $params[ 'Uri' ] = "$baseCloudURL/$($id)s$dateFilter"
     
                    ##$params.uri = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}s" -f $protocol , $customerid , $version ,  $id ) ## + (Get-DateRanges -query $id -from $from -to $to -selective -oDataVersion $oDataVersion)
                }
                else
                {
                    $params.uri = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}s$dateFilter" -f $protocol , $ddc , $version , $id )
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
                                $property | Resolve-CrossReferences -cloud:$cloud -include $include -exclude $exclude -from $from -to $to
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
        $properties | Where-Object -filterscript { $_.Name -ine 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } | ForEach-Object `
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
                            if( $_.Name -ine $property.Name -and ( -Not $previousProperties -or -Not ( $previousProperties | Where-Object -FilterScript { $_.Name -eq $_.Name } ))) ## don't lookup self or a key if it was one we previously looked up
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

# see https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
#     https://stackoverflow.com/questions/2859790/the-request-was-aborted-could-not-create-ssl-tls-secure-channel
if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type )
{
    Add-Type -ErrorAction SilentlyContinue -TypeDefinition @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object -TypeName TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

if( $PsCmdlet.ParameterSetName -eq 'cloud' )
{
    if( -Not $PSBoundParameters[ 'authtoken' ] )
    {
        if( $PSBoundParameters[ 'profileName' ] )
        {
            Add-PSSnapin -Name Citrix.Sdk.Proxy.*
            if( ! ( Get-Command -Name Get-XDAuthentication -ErrorAction SilentlyContinue ) )
            {
                Throw "Unable to find the Get-XDAuthentication cmdlet - is the Virtual Apps and Desktops Remote PowerShell SDK installed ?"
            }
            Get-XDAuthentication -ProfileName $profileName
            if( [string]::IsNullOrEmpty( $customerid ) )
            {
                $customerid = (Get-XDCredentials -ProfileName $profileName).Credentials.CustomerId
                if( [string]::IsNullOrEmpty( $customerid ) )
                {
                    Throw "Failed to get customer id from profile $profileName"
                }
            }
            $authtoken = $GLOBAL:XDAuthToken
        }
        elseif( -Not [string]::IsNullOrEmpty( $clientId ) ) ## don't use Remote PS SDK or use -clientId and -clientSecret
        {
            $authtoken = Get-BearerToken -clientId $clientId -clientSecret $clientSecret
        }
        else
        {
             Throw "For Cloud you must user -profilename if Remote PS SDK installed or specify -clientId and -ClientSecret"
        }
        if( -Not $? -or [string]::IsNullOrEmpty( $authtoken ) )
        {
            Throw "Failed to get authentication token for Cloud customer id $customerid"
        }
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
    
    [string]$countQuery = $null
    if( -Not [string]::IsNullOrEmpty( $query ) )
    {
        $countQuery = '&$count=true'
    }
    if( $PsCmdlet.ParameterSetName -eq 'cloud' )
    {
        ##$params[ 'Uri' ] = ( "{0}://{1}.xendesktop.net/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $customerid , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion)
        $params[ 'Uri' ] = "$baseCloudURL/$query" + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $oDataVersion) + $countQuery
        $cloud = $true
    }
    else
    {
        $params[ 'Uri' ] = ( "{0}://{1}/Citrix/Monitor/OData/v{2}/Data/{3}" -f $protocol , $ddc , $version , $query ) + (Get-DateRanges -query $query -from $from -to $to -oDataVersion $version) + $countQuery
    }

    Write-Verbose "URL : $($params.Uri)"

    try
    {
        [int]$results = 0
        [int]$requests = 0
        [string]$lasturi = $params.uri
        [bool]$firstQuery = $true
        [int]$countOfItems = 0

        do
        {
            $requests++
            $resultsPage = $null
            $resultsPage = Invoke-RestMethod @params

            if( $null -ne $resultsPage )
            {
                if( $firstQuery )
                {
                    if( $resultsPage.psobject.Properties[ '@odata.count' ] )
                    {
                        $countOfItems = $resultsPage.'@odata.count'
                        Write-Verbose -Message "There are $countOfItems $query items"
                    }
                    $firstQuery = $false
                }
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
                    $params.uri = $resultsPage.'@odata.nextLink' -replace [regex]::Escape( $countQuery )
                    ## prevent infinite loop if something goes wrong
                    if( $params.uri -ne $lasturi )
                    {
                        Write-Verbose -Message "More data available ($($countOfItems - $results)), fetching from $($params.uri)"
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
        } while( $resultsPage -and ( $maximumItems -le 0 -or $results -lt $maximumItems ))
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
        $data[0].PSObject.Properties | Resolve-CrossReferences -cloud:$cloud -Include $includePropertyRegex -Exclude $excludePropertyRegex -from $from -to $to

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

            $datum.PSObject.Properties | Where-Object -FilterScript { $_.Name -ne 'sid' -and ( $_.Name -match '^(.*)Id$' -or $_.Name -match '^(Session)Key$' ) -and -Not [string]::IsNullOrEmpty( $Matches[1] ) } | . { Process `
            {
                $property = $_
                Resolve-NestedProperties -properties $property | . { Process `
                {
                    $_.PSObject.Properties | Where-Object -FilterScript { $_.MemberType -eq 'NoteProperty' } | . { Process `
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
