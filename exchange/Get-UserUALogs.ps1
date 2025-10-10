New-Alias -Name 'UALog' -Value 'Get-UserUALogs' -Force
New-Alias -Name 'UALogs' -Value 'Get-UserUALogs' -Force
New-Alias -Name 'GetUALog' -Value 'Get-UserUALogs' -Force
New-Alias -Name 'GetUALogs' -Value 'Get-UserUALogs' -Force
New-Alias -Name 'GetUserUALog' -Value 'Get-UserUALogs' -Force
New-Alias -Name 'GetUserUALogs' -Value 'Get-UserUALogs' -Force
function Get-UserUALogs {
    <#
	.SYNOPSIS
    Runs multiple queries to pull all unified audit logs records related to a specific user.
    
	.NOTES
	Version: 1.4.0
    1.4.0 - Updating to add metadata object, use shorter file names.
    1.3.0 - Updated to output objects.
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,

        # relative date range
        [int] $Days, # default value set at #DEFAULTDAYS
        # absolute date range
        [string] $Start,
        [string] $End,

        [string[]] $Operations,
        [switch] $RiskyOperations,

        [boolean] $WaitOnMessageTrace = $false,

        [boolean] $Xml = $true,
        [boolean] $Excel = $true,
        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $AllLogs = [System.Collections.Generic.List[psobject]]::new()
        $FileNamePrefix = 'UnifiedAuditLogs'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        # $Yellow = @{ ForegroundColor = 'Yellow' }

        if ($Test) {$Global:IRTTestMode = $true}

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

        # verify connected to exchange
        try {
            [void](Get-AcceptedDomain)
        }
        catch {
            $ErrorParams = @{
                Category    = 'ConnectionError'
                Message     = "Not connected to Exchange. Run Connect-ExchangeOnline."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams
        } 

        # get client domain name for file output
        $DefaultDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
        $DomainName = $DefaultDomain.DomainName -split '\.' | Select-Object -First 1

        #region DATE RANGE

        # validate only relative or absolute
        if ($Days -and ($Start -or $End)) {
            $ErrorParams = @{
                Category    = 'InvalidArgument'
                Message     = "Choose either relative range with -Days or absolute range with -Start and -End."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams  
        }

        # validate both start and end used
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
            # make sure dates are in correct order
            if ($StartDateUtc -gt $EndDateUtc) {
                $Temp = $StartDateUtc
                $StartDateUtc = $EndDateUtc
                $EndDateUtc = $Temp
            }
        }
        # create objects based on days
        else {
            # set default value for days ### must be done after checking for relative/absolute arguments
            if (-not $Days) {
                $Days = 1 #DEFAULTDAYS
            }

            $StartDateUtc = (Get-Date).AddDays($Days * -1).ToUniversalTime() 
            $EndDateUtc = (Get-Date).ToUniversalTime()
        }

        # set file name date to query end date
        $FileNameDateFormat = 'yy-MM-dd_HH-mm'
        $FileNameDateString = $EndDateUtc.ToLocalTime().ToString($FileNameDateFormat)

        #region OPERATIONS

        $OperationsSet = [System.Collections.Generic.Hashset[string]]::new()
        # add user specified operations
        foreach ($o in $Operations) {[void]$OperationsSet.Add($o)}
        # add risk operations
        if ($RiskyOperations) {
            # import alloperations csv
            $ModulePath = $PSScriptRoot
            $AllOperationsFileName = 'unified_audit_log-all_operations.csv'
            $OperationsCsvPath = Join-Path -Path $ModulePath -ChildPath "\unified_audit_log-data\${AllOperationsFileName}"
            $OperationsCsvData = Import-Csv -Path $OperationsCsvPath

            # get high risk operations
            $HighRiskOperations = ($OperationsCsvData | Where-Object {$_.Risk -eq 'High'}).Operation

            # add to set
            foreach ($o in $HighRiskOperations) {[void]$OperationsSet.Add($o)}

            # FIXME get these properly tagged in spreadsheet.
            # app consent
            # [void]$OperationsSet.Add("Add delegated permission grant.") #FIXME
            # mfa changes
            # [void]$OperationsSet.Add("User registered security info") #FIXME
            # [void]$OperationsSet.Add("User registered all required security info") #FIXME
        }
    }

    process {

        #region USER LOOP

        foreach ($ScriptUserObject in $ScriptUserObjects) {

            $AllLogs.Clear()

            $UserEmail = $ScriptUserObject.UserPrincipalName
            $UserId = $ScriptUserObject.Id
            $UserIdNumbers = $UserId -replace '-', ''
            $UserName = $UserEmail -split '@' | Select-Object -First 1
            if ($Days) {
                $XmlOutputPath = "${FileNamePrefix}_${Days}Days_${UserName}_${FileNameDateString}.xml"
            }
            else {
                $XmlOutputPath = "${FileNamePrefix}_${UserName}_${FileNameDateString}.xml"
            }
            # build query params
            $BaseParams = @{
                ResultSize     = 5000
                SessionCommand = 'ReturnLargeSet'
                Formatted      = $true
                StartDate      = $StartDateUtc
                EndDate        = $EndDateUtc
            }

            # add operations, if specified
            if (($OperationsSet | Measure-Object).Count -gt 0) {
                $BaseParams['Operations'] = $OperationsSet
            }

            $QueryTable = [ordered]@{
                '1' = @{
                    Params = @{
                        UserIds = $UserEmail, $UserId, $UserIdNumbers
                    }
                    Text   = "Running -UserIds query for ${UserEmail}, ${UserId}, ${UserIdNumbers}"
                }
                '2' = @{
                    Params = @{
                        FreeText = $UserEmail
                    }
                    Text   = "Running -Freetext query for ${UserEmail}"
                }
                '3' = @{
                    Params = @{
                        FreeText = $UserId
                    }
                    Text   = "Running -Freetext query for ${UserId}"
                }
                '4' = @{
                    Params = @{
                        FreeText = $UserIdNumbers
                    }
                    Text   = "Running -Freetext query for ${UserIdNumbers}"
                }
            }

            # run queries
            foreach ( $QueryDict in $QueryTable.GetEnumerator() ) {

                # build final params
                $FirstPageParams = @{}
                $BaseParams.GetEnumerator() | ForEach-Object { $FirstPageParams[$_.Key] = $_.Value }
                $QueryDict.Value.Params.GetEnumerator() | ForEach-Object { $FirstPageParams[$_.Key] = $_.Value }

                $Text = $QueryDict.Value.Text

                # run query
                Write-Host @Blue $Text
                $Page = Search-UnifiedAuditLog @FirstPageParams
                $LogCount = ($Page | Measure-Object).Count

                if ($LogCount -gt 0) {

                    Write-Host @Blue "Retrieved ${LogCount} logs."

                    # add to list
                    foreach ($i in $Page) {$AllLogs.Add($i)}

                    # extract sessionid for paging
                    $SessionId = $Page[0].SessionId
                    $PageCount = 2
                    $NextPageParams = $FirstPageParams
                    $NextPageParams['SessionId'] = $SessionId
                }
                else {
                    Write-Host @Red "Retrieved 0 logs."
                }

                # retrieve pages
                while ($LogCount -eq 5000) {

                    Write-Host @Blue "Requesting page ${PageCount}."
                    $Page = Search-UnifiedAuditLog @NextPageParams
                    $LogCount = @($Page).Count

                    if ( $LogCount -gt 0 ) {

                        Write-Host @Blue "Retrieved ${LogCount} logs."

                        # add to list
                        foreach ($i in $Page) {$AllLogs.Add($i)}

                        # extract sessionid for paging
                        $SessionId = $Page[0].SessionId
                    }
                    else {
                        Write-Host @Red "Retrieved 0 logs."
                    }

                    $PageCount++
                }
            }

            # exit if no logs returned
            if (($AllLogs | Measure-Object).Count -eq 0) {
                Write-Host @Red "0 total logs retrieved. Exiting."
                return
            }

            # remove duplicates
            $UniqueLogIds = [System.Collections.Generic.HashSet[string]]::new()
            $UniqueLogs = [System.Collections.Generic.List[psobject]]::new()
            foreach ($Log in $AllLogs) {
                if ($UniqueLogIds.Add([string]$Log.Identity)) { 
                    $UniqueLogs.Add($Log) | Out-Null
                }
            }

            ### sort list
            # build comparison script
            $PropertyName = 'CreationDate'
            $Descending = $true
            $Comparison = [System.Comparison[PSObject]] {
                param($X, $Y)
                $Result = $X.$PropertyName.CompareTo($Y.$PropertyName)
                if ( $Descending ) {
                    return -1 * $Result
                }
                return $Result
            }
            $UniqueLogs.Sort($Comparison)

            # add metadata to results

            $UniqueLogs.Insert(0,
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

            # export to xml
            if ($Xml) {
                Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
                $UniqueLogs | Export-Clixml -Depth 10 -Path $XmlOutputPath
            }

            # export excel spreadsheet
            if ($Excel) {
                $Params = @{
                    Logs = $UniqueLogs
                    WaitOnMessageTrace = $WaitOnMessageTrace
                }
                Show-UALogs @Params
            }
        }
    }
}