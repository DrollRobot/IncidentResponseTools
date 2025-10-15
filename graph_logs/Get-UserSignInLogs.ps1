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
    [CmdletBinding(DefaultParameterSetName = 'UserObjects')]
    param (
        [Parameter(Position = 0,ParameterSetName='UserObjects')]
        [Alias('UserObject')]
        [psobject[]] $UserObjects,

        [Parameter(ParameterSetName='AllUsers')]
        [switch] $AllUsers,

        # [string[]] $IpAddresses, # add option to search by ip address FIXME

        # relative date range
        [int] $Days, # default value set at #DEFAULTDAYS
        # absolute date range
        [string] $Start,
        [string] $End,

        [switch] $NonInteractive,
        [switch] $DeviceCode, # FIXME not working? might relate to api bug? 

        [boolean] $Beta = $true,
        [boolean] $Excel = $true,
        [boolean] $IpInfo = $true,
        [switch] $Test,
        [boolean] $Xml = $true
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $DateString = Get-Date -Format $FileNameDateFormat

        # if ( $NonInteractive -and $Days -eq 30) { # if script default, change to 3 days
        #     $Days = 3 # FIXME temporarily commending out until api issue is fixed
        # }

        # colors
        $Blue = @{ForegroundColor = 'Blue'}
        # $Green = @{ForegroundColor = 'Green'}
        # $Red = @{ForegroundColor = 'Red'}
        # $Magenta = @{ForegroundColor = 'Magenta'}
        $Yellow = @{ForegroundColor = 'Yellow'}

        if ($Test) {$Script:Test = $true}

        # create user objects depending on parameters used
        switch ( $ParameterSet ) {
            'UserObjects' {
                # if users passed via script argument:
                if (($UserObjects | Measure-Object).Count -gt 0) {
                    $ScriptUserObjects = $UserObjects
                }
                # if not, look for global objects
                else {
                    
                    # get from global variables
                    $ScriptUserObjects = Get-GraphGlobalUserObjects
                    
                    # if none found, exit
                    if ( -not $ScriptUserObjects -or $ScriptUserObjects.Count -eq 0 ) {
                        Write-Host @Red "${Function}: No user objects passed or found in global variables."
                        return
                    }
                    if (($ScriptUserObjects | Measure-Object).Count -eq 0) {
                        $ErrorParams = @{
                            Category    = 'InvalidArgument'
                            Message     = "No -UserObjects argument used, no `$Global:UserObjects present."
                            ErrorAction = 'Stop'
                        }
                        Write-Error @ErrorParams
                    }
                }
            }
            'AllUsers' {
                # build user object with null principal name
                $ScriptUserObjects = @(
                    [pscustomobject]@{
                        UserPrincipalName = $null
                    }
                )
            }
        }

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        #region DATE RANGE

        # validate only days or (start and end)
        if ($Days -and ($Start -or $End)) {
            $ErrorParams = @{
                Category    = 'InvalidArgument'
                Message     = "Choose either relative range with -Days or absolute range with -Start and -End."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams  
        }

        # validate if start or end used, both were used.
        if (($Start -and -not $End) -or
            ($End -and -not $Start)
        ) {
            $ErrorParams = @{
                Category    = 'InvalidArgument'
                Message     = "Specify both -Start and -End"
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams  
        }

        # attempt to parse user input dates into datetime objects
        if ($Start -and $End) {
            $DateRangeType = 'Absolute'
            # start - convert user string into object
            try {
                $StartDate = Get-Date -Date $Start -ErrorAction 'Stop'
                $StartDateUtc = [DateTime]::SpecifyKind($StartDate, [DateTimeKind]::Local).ToUniversalTime()
            }
            catch {
                $ErrorParams = @{
                    Category    = 'InvalidArgument'
                    Message     = "-Start invalid. Use format 'MM/dd/yy hh:mm(tt)"
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
            }
            # end - convert user string into object
            try {
                $EndDate = Get-Date -Date $End -ErrorAction 'Stop'
                $EndDateUtc = [DateTime]::SpecifyKind($EndDate, [DateTimeKind]::Local).ToUniversalTime()
            }
            catch {
                $ErrorParams = @{
                    Category    = 'InvalidArgument'
                    Message     = "-End invalid. Use format 'MM/dd/yy hh:mm(tt)"
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
            }
            # make sure earliest date is the start date
            if ($StartDateUtc -gt $EndDateUtc) {
                $Temp = $StartDateUtc
                $StartDateUtc = $EndDateUtc
                $EndDateUtc = $Temp
            }
            # set days to match range
            $Days = [Int]([Math]::Floor( ($EndDate - $StartDate).TotalDays ))
        }
        # create objects based on days
        else {
            $DateRangeType = 'Relative'
            # set default value for days ### must be done after checking for relative/absolute arguments
            if (-not $Days) {
                $Days = 30 #DEFAULTDAYS
                # FIXME defaulting to 30 days because of api bug related to filters
                # https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3146#issuecomment-2752675332
            }

            $StartDateUtc = (Get-Date).AddDays($Days * -1).ToUniversalTime()
            $EndDateUtc = (Get-Date).ToUniversalTime()
        }
    }

    process {

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            $FilterStrings = [System.Collections.Generic.List[string]]::new()

            #region FILTERS

            # users
            switch ( $ParameterSet ) {
                'UserObjects' {
                    $UserEmail = $ScriptUserObject.UserPrincipalName
                    $UserName = $UserEmail -split '@' | Select-Object -First 1
                    $FilterStrings.Add( "UserId eq '$($ScriptUserObject.Id)'" )
                }
                'AllUsers' {
                    $UserName = 'AllUsers'
                    # don't add a user filter
                }
            }

            # build file names # must be after username is set
            if ( $NonInteractive ) {
                $FileNamePrefix = 'NonInteractiveLogs'
            }
            else {
                $FileNamePrefix = 'SignInLogs'
            }
            $XmlOutputPath = "${FileNamePrefix}_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"

            # device code
            if ( $DeviceCode ) {
                $FilterStrings.Add( "AuthenticationProtocol eq 'devicecode'" )
            }

            # time range
            if ($DateRangeType -eq 'Relative') {
                if ($Days -ne 30) { # don't use filter if date range is maximum
                    $DateString = $StartDateUtc.ToString( "yyyy-MM-ddTHH:mm:ssZ" )
                    $FilterStrings.Add( "createdDateTime ge ${DateString}" )
                }
            }
            elseif ($DateRangeType -eq 'Absolute') {
                $StartDateString = $StartDateUtc.ToString( "yyyy-MM-ddTHH:mm:ssZ" )
                $FilterStrings.Add( "createdDateTime ge ${StartDateString}" )
                $EndDateString = $EndtDateUtc.ToString( "yyyy-MM-ddTHH:mm:ssZ" )
                $FilterStrings.Add( "createdDateTime le ${EndDateString}" )
            }

            # non interactive
            if ( $NonInteractive ) {
                $FilterStrings.Add( "signInEventTypes/any(t: t eq 'NonInteractiveUser')" )
            }
            $FilterString = $FilterStrings -join " and "

            #region QUERY LOGS
            # user messages
            if ( $NonInteractive ) {
                Write-Host @Blue "`nRetrieving ${Days} days of noninteractive sign-in logs for ${UserName}." | Out-Host
            }
            else {
                Write-Host @Blue "`nRetrieving ${Days} days of sign-in logs for ${UserName}." | Out-Host
            }
            if ($Script:Test) {
                Write-Host @Yellow "${Function}: Filter string: '${FilterString}'" | Out-Host
            }
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
                [System.Collections.Generic.List[PSObject]]$Logs = Get-MgBetaAuditLogSignIn @GetParams |
                    Select-Object $GetProperties
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
                [System.Collections.Generic.List[PSObject]]$Logs = Get-MgAuditLogSignIn @GetParams |
                    Select-Object $GetProperties
            }

            # add metadata to results
            $Logs.Insert(0,
                [pscustomobject]@{
                    Metadata = $true
                    UserObject = $ScriptUserObject
                    UserEmail = $UserEmail
                    UserName = $UserName
                    StartDate = $StartDateUtc.ToLocalTime()
                    EndDate = $EndDateUtc.ToLocalTime()
                    Days = $Days
                    DomainName = $DomainName
                    FileNamePrefix = $FileNamePrefix
                }
            )

            #region OUTPUT

            # show count, export
            $LogCount = ($Logs | Measure-Object).Count
            if ($LogCount -gt 0) {
                Write-Host @Blue "Retrieved ${LogCount} logs."

                # export to xml
                if ($Xml) {
                    Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
                    $Logs | Export-Clixml -Depth 10 -Path $XmlOutputPath
                }

                # export excel spreadsheet
                if ($Excel) {
                    $Params = @{
                        Logs = $Logs
                        IpInfo = $IpInfo
                    }
                    Show-SignInLogs @Params
                }
            }
            else {
                Write-Host @Red "Retrieved 0 logs."
            }
        }
    }
}