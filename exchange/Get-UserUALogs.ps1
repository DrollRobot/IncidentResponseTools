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
	Version: 1.3.0
    1.3.0 - Updated to output objects.
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,

        [int] $Days = 1,
        [boolean] $Xml = $true,
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

        # verify connected to exchange
        try {
            $Domain = Get-AcceptedDomain
        }
        catch {}
        if ( -not $Domain ) {
            throw "Not connected to ExchangeOnlineManagement. Run Connect-ExchangeOnline. Exiting."
        }
     
        # variables
        $AllLogs = [System.Collections.Generic.List[psobject]]::new()
        $FileNameDateFormat = "yy-MM-dd_HH-mm"

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # get client domain name for file output
        $DefaultDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
        $DomainName = $DefaultDomain.DomainName -split '\.' | Select-Object -First 1

        # get date/time string for filename
        $DateString = (Get-Date).ToString($FileNameDateFormat)
    }

    process {

        foreach ($ScriptUserObject in $ScriptUserObjects) {

            $AllLogs.Clear()

            $UserPrincipalName = $ScriptUserObject.UserPrincipalName
            $UserId = $ScriptUserObject.Id
            $UserIdNumbers = $UserId -replace '-', ''
            $UserName = $UserPrincipalName -split '@' | Select-Object -First 1
            $XmlOutputPath = "UnifiedAuditLogs_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"

            # build query params
            $EndDate = (Get-Date).ToUniversalTime()
            $BaseParams = @{
                ResultSize     = 5000
                SessionCommand = 'ReturnLargeSet'
                Formatted      = $true
                StartDate      = (Get-Date).AddDays($Days * -1).ToUniversalTime() 
                EndDate        = $EndDate
            }
            $QueryTable = [ordered]@{
                '1' = @{
                    Params = @{
                        UserIds = $UserPrincipalName, $UserId, $UserIdNumbers
                    }
                    Text   = "Running -UserIds query for ${UserPrincipalName}, ${UserId}, ${UserIdNumbers}"
                }
                '2' = @{
                    Params = @{
                        FreeText = $UserPrincipalName
                    }
                    Text   = "Running -Freetext query for ${UserPrincipalName}"
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
                    $Page | ForEach-Object { $AllLogs.Add( $_ ) }

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
                        $Page | ForEach-Object { $AllLogs.Add( $_ ) }

                        # extract sessionid for paging
                        $SessionId = $Page[0].SessionId
                    }
                    else {
                        Write-Host @Red "Retrieved 0 logs."
                    }

                    $PageCount++
                }
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

            # export to xml
            if ($Xml) {
                Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
                $UniqueLogs | Export-Clixml -Depth 10 -Path $XmlOutputPath
            }

            # create spreadsheet
            $ShowParams = @{
                Logs = $UniqueLogs
                DomainName = $DomainName
                UserName = $UserName
                Days = $Days
                EndDate = $EndDate
                Open = $Open
            }
            Show-UALogs @ShowParams
        }
    }
}