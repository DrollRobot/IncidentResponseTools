New-Alias -Name "SILog" -Value "Get-UserSignInLogs" -Force
New-Alias -Name "SILogs" -Value "Get-UserSignInLogs" -Force
New-Alias -Name "GetSILog" -Value "Get-UserSignInLogs" -Force
New-Alias -Name "GetSILogs" -Value "Get-UserSignInLogs" -Force
New-Alias -Name "Get-UserSignInLog" -Value "Get-UserSignInLogs" -Force
function Get-UserSignInLogs {
    <#
	.SYNOPSIS
	Downloads user sign in logs.	
	
	.NOTES
	Version: 1.1.0
	#>
    [CmdletBinding(DefaultParameterSetName = 'User')]
    param (
        [Parameter(Position = 0,ParameterSetName='User')]
        [Alias('UserObject')]
        [psobject[]] $UserObjects,

        # [string[]] $IpAddresses, # add option to search by ip address FIXME

        [int] $Days = 30, # FIXME defaulting to 30 days because of api bug related to filters
        # https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3146#issuecomment-2752675332
        [switch] $NonInteractive,
        [switch] $DeviceCode, # FIXME not working? might relate to api bug? 

        [boolean] $Beta = $true,
        [boolean] $Xml = $true,
        [boolean] $Script = $false,
        [boolean] $Open = $true
    )

    begin {

        # if user objects not passed directly, find global
        if ( -not $UserObjects -or $UserObjects.Count -eq 0 ) {
        
            # get from global variables
            $ScriptUserObjects = Get-GraphGlobalUserObjects
                                
            # if none found, exit
            if ( -not $ScriptUserObjects -or $ScriptUserObjects.Count -eq 0 ) {
                throw "No user objects passed or found in global variables."
            }
        }
        else {
            $ScriptUserObjects = $UserObjects
        }

        # variables
        $XmlPaths = [System.Collections.Generic.List[string]]::new()

        # if ( $NonInteractive -and $Days -eq 30) { # if script default, change to 3 days
        #     $Days = 3 # FIXME temporarily commending out until api issue is fixed
        # }

        # get datetime for query
        $QueryStart = ( Get-Date ).AddDays( $Days * -1 ).ToString( "yyyy-MM-ddTHH:mm:ssZ" )

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        # get date/time string for filename
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $DateString = Get-Date -Format $FileNameDateFormat
    }

    process {

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            $FilterStrings = [System.Collections.Generic.List[string]]::new()
            $UserEmail = $ScriptUserObject.UserPrincipalName
            $UserName = $UserEmail -split '@' | Select-Object -First 1
            $UserId = $ScriptUserObject.Id 

            # build file names
            # FIXME make a case statement, include device code?
            if ( $NonInteractive ) {
                $XmlOutputPath = "NonInteractiveLogs_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"
            }
            else {
                $XmlOutputPath = "SignInLogs_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"
            }

            # build filter string
            if ( $DeviceCode ) {
                $FilterStrings.Add( "AuthenticationProtocol eq 'devicecode'" )
            }
            else {
                $FilterStrings.Add( "UserId eq '${UserId}'" )
            }
            if ($Days -ne 30) { # don't use filter if date range is default
                $FilterStrings.Add( "createdDateTime ge ${QueryStart}" )
            }
            if ( $NonInteractive ) {
                $FilterStrings.Add( "signInEventTypes/any(t: t eq 'NonInteractiveUser')" )
            }
            $FilterString = $FilterStrings -join " and "

            ### get logs
            # user messages
            if ( $NonInteractive ) {
                Write-Host @Blue "`nRetrieving ${Days} days of noninteractive sign-in logs for ${UserName}." | Out-Host
            }
            else {
                Write-Host @Blue "`nRetrieving ${Days} days of sign-in logs for ${UserName}." | Out-Host
            }
            Write-Verbose "Filter string: ${FilterString}" | Out-Host
            # Write-Host @Blue "This can take up to 5 minutes, depending on the number of logs." | Out-Host

            # query logs
            if ($Beta) { # default is to use beta, which returns more information
                $GetProperties = @(
                    'AppDisplayName'
                    'AuthenticationProtocol'
                    'CorrelationID'
                    'CreatedDateTime'
                    'DeviceDetail'
                    'IpAddress'
                    'Location'
                    'ResourceId'
                    'Status'
                    # 'UniqueTokenIdentifier'
                    'UserAgent'
                    'UserPrincipalName'
                )
                $GetParams = @{
                    Filter = $FilterString
                    Property = $GetProperties
                    All = $true
                }
                $Logs = Get-MgBetaAuditLogSignIn @GetParams | Select-Object $GetProperties
            }
            else { # if $Beta = $false
                $GetProperties = @(
                    'AppDisplayName'
                    'CorrelationID'
                    'CreatedDateTime'
                    'DeviceDetail'
                    'IpAddress'
                    'Location'
                    'ResourceId'
                    'Status'
                    'UniqueTokenIdentifier'
                    'UserAgent'
                    'UserPrincipalName'
                )
                $GetParams = @{
                    Filter = $FilterString
                    Property = $GetProperties
                    All = $true
                }
                $Logs = Get-MgAuditLogSignIn @GetParams | Select-Object $GetProperties
            }

            # show count
            $Count = @( $Logs ).Count
            if ($Count -gt 0) {
                Write-Host @Blue "Retrieved ${Count} logs."

                # export to xml
                if ($Xml) {
                    Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
                    $Logs | Export-Clixml -Depth 10 -Path $XmlOutputPath
                }

                $XmlPaths.Add( $XmlOutputPath )
            }
            else {
                Write-Host @Red "Retrieved 0 logs."
            }
        } # end of user loop

        if ($Script) {
            return $XmlPaths
        }
        elseif ($Open) {
            foreach ($XmlPath in $XmlPaths) {
                $ShowParams = @{
                    XmlPath = $XmlPath
                    Open = $Open
                }
                Show-SignInLogs @ShowParams
            }
        }
    }
}