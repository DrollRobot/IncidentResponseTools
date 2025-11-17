New-Alias -Name 'GetUAAppLogs' -Value 'Get-ExchangeUAAppLogs' -Force
function Get-ExchangeUAAppLogs {
    <#
	.SYNOPSIS
    Runs multiple queries to pull all unified audit logs records related to a specific application/service principal.
    
	.NOTES
	Version: 1.2.0
	#>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory)]
        [Alias( 'Guids', 'Strings' )]
        [psobject[]] $AppGuids,

        [Parameter(Mandatory)]
        [string] $SearchName,

        [int] $Days = 3,
        [boolean] $Xml = $true,
        [switch] $NoOpen,
        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        if ($Test -or $Script:Test) {
            $Script:Test = $true
            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $DateString = Get-Date -Format $FileNameDateFormat

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        $AllLogs = [System.Collections.Generic.List[psobject]]::new()
        $UniqueLogIds = [System.Collections.Generic.HashSet[psobject]]::new()
        $UniqueLogs = [System.Collections.Generic.List[psobject]]::new()


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
    }

    process {

        foreach ( $Guid in $AppGuids ) {

            $GuidJustNumbers = $Guid -replace '-', ''
            $XmlOutputPath = "UnifiedAuditLogs_Raw_${Days}Days_${DomainName}_${SearchName}_${DateString}.xml"

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
                        FreeText = $Guid
                    }
                    Text   = "Running -Freetext query for ${Guid}"
                }
                '2' = @{
                    Params = @{
                        FreeText = $GuidJustNumbers
                    }
                    Text   = "Running -Freetext query for ${GuidJustNumbers}"
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
                    $NextPageParams['SessionCommand'] = 'Continue'
                    $NextPageParams['SessionId'] = $SessionId
                }
                else {
                    Write-Host @Red "Retrieved 0 logs."
                }

                # retrieve pages
                while ( $LogCount -eq 5000 ) {

                    Write-Host "Retrieving page ${PageCount}."
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
        }

        # remove duplicates
        foreach ( $Log in $AllLogs ) {
            
            # if log identity has already been processed, skip
            if ( $UniqueLogIds -contains $Log.Identity ) {
                continue
            }
            # else, record identity and add log
            else {
                $UniqueLogIds.Add( $Log.Identity ) | Out-Null
                $UniqueLogs.Add( $Log ) | Out-Null
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
            UserName = $SearchName
            Days = $Days
            EndDate = $EndDate
        }
        if ( $NoOpen ) {
            $ShowParams['NoOpen'] = $true
        }
        Show-UALogs @ShowParams
    }
}