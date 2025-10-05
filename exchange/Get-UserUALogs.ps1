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
        [int] $Days = 1,

        [boolean] $WaitOnMessageTrace = $false,

        [boolean] $Xml = $true,
        [boolean] $Excel = $true
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $AllLogs = [System.Collections.Generic.List[psobject]]::new()
        $FileNameDateFormat = 'yy-MM-dd_HH-mm'
        $FileNameDateString = (Get-Date).ToString($FileNameDateFormat)
        $FileNamePrefix = 'UnifiedAuditLogs'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        # $Yellow = @{ ForegroundColor = 'Yellow' }


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
                    Message     = "${Function}: No -UserObjects, No `$Global:UserObjects."
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
            }
        }

        # verify connected to exchange
        try {
            $Domain = Get-AcceptedDomain
        }
        catch {}
        if ( -not $Domain ) {
            $ErrorParams = @{
                Category    = 'ConnectionError'
                Message     = "${Function}: Not connected to Exchange. Run Connect-ExchangeOnline."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams
        }

        # get client domain name for file output
        $DefaultDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
        $DomainName = $DefaultDomain.DomainName -split '\.' | Select-Object -First 1
    }

    process {

        foreach ($ScriptUserObject in $ScriptUserObjects) {

            $AllLogs.Clear()

            $UserEmail = $ScriptUserObject.UserPrincipalName
            $UserId = $ScriptUserObject.Id
            $UserIdNumbers = $UserId -replace '-', ''
            $UserName = $UserEmail -split '@' | Select-Object -First 1
            $XmlOutputPath = "${FileNamePrefix}_${Days}Days_${UserName}_${FileNameDateString}.xml"

            # build query params
            $EndDateUtc = (Get-Date).ToUniversalTime()
            $StartDateUtc = (Get-Date).AddDays($Days * -1).ToUniversalTime() 
            $BaseParams = @{
                ResultSize     = 5000
                SessionCommand = 'ReturnLargeSet'
                Formatted      = $true
                StartDate      = $StartDateUtc
                EndDate        = $EndDateUtc
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
                $LogCount = @($Page).Count

                if ( $LogCount -gt 0 ) {

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
                $ErrorParams = @{
                    Category    = 'InvalidResult'
                    Message     = "${Function}: 0 total logs retrieved. Exiting."
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
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