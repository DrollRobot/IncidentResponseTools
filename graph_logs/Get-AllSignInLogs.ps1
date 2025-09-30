function Get-AllSignInLogs {
    <#
	.SYNOPSIS
	Downloads all sign in logs for the tenant.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding(DefaultParameterSetName = 'User')]
    param (
        # [int] $Days = 30, # commenting out while api bug exists FIXME
        # https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3146#issuecomment-2752675332
        [switch] $NonInteractive,
        [switch] $NoBeta,
        # [switch] $DeviceCode # not working? might relate to api bug? FIXME
        [switch] $Script,
        [boolean] $Open = $true
    )

    begin {

        # setting days to 30 while API bug exists FIXME
        $Days = 30

        # variables
        $UserName = 'AllUsers'
        $FilterStrings = [System.Collections.Generic.List[string]]::new()
        $XmlPaths = [System.Collections.Generic.List[string]]::new()
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
        $BetaGetProperties = @(
            'AppDisplayName'
            'AuthenticationProtocol'
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
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        # if ( $NonInteractive ) { # FIXME temporarily commending out until graph issue is fixed
        #     $Days = 3
        # }
        # else {
        #     $Days = 30
        # }

        # get datetime for query         # FIXME temporarily removing to address graph bug
        # $QueryStart = ( Get-Date ).AddDays( $Days * -1 ).ToString( "yyyy-MM-ddTHH:mm:ssZ" )

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        # get date/time string for filename
        $DateString = Get-Date -Format $FileNameDateFormat
    }

    process {

        # variables 

        # build file names
        if ( $NonInteractive ) {
            $XmlOutputPath = "NonInteractiveLogs_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"
        }
        else {
            $XmlOutputPath = "SignInLogs_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"
        }

        # build filter string
        # $FilterStrings.Add( "createdDateTime ge ${QueryStart}" ) # temporarily removed because of API bug FIXM
        if ( $DeviceCode ) {
            $FilterStrings.Add( "AuthenticationProtocol eq 'devicecode'" )
        }
        if ( $NonInteractive ) {
            $FilterStrings.Add( "signInEventTypes/any(t: t eq 'NonInteractiveUser')" )
        }
        $FilterString = $FilterStrings -join " and "

        ### get logs
        # user messages
        if ( $NonInteractive ) {
            Write-Host @Blue "`nRetrieving ${Days} days of noninteractive sign-in logs for ${UserName}."
        }
        else {
            Write-Host @Blue "`nRetrieving ${Days} days of sign-in logs for ${UserName}."
        }
        Write-Verbose "Filter string: ${FilterString}"
        # Write-Host @Blue "This can take up to 5 minutes, depending on the number of logs."

        # query logs
        if ( $NoBeta ) {
            $GetParams = @{
                Filter = $FilterString
                Property = $GetProperties
                All = $true
            }
            $Logs = Get-MgAuditLogSignIn @GetParams | Select-Object $GetProperties
        }
        else {
            $GetParams = @{
                Filter = $FilterString
                Property = $BetaGetProperties
                All = $true
            }
            $Logs = Get-MgBetaAuditLogSignIn @GetParams | Select-Object $GetProperties
        }

        # show count
        $Count = @($Logs).Count
        if ($Count -gt 0) {
            Write-Host @Blue "Retrieved ${Count} logs."
        }
        else {
            Write-Host @Red "Retrieved 0 logs."
            return
        }

        # export to xml
        Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
        $Logs | Export-Clixml -Depth 10 -Path $XmlOutputPath

        if ($Script -or $Open) {
            $XmlPaths.Add( $XmlOutputPath )
        }

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