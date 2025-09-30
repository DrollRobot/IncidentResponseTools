function Group-ByCorrelationId {
    [CmdletBinding()]
    param (
        [Parameter( Mandatory, Position = 0 )]
        [psobject[]] $Logs
    )

    begin {

        # create table of human readable session ids
        $UniqueSessionIDs = $Logs.CorrelationID | Sort-Object -Unique
        $SessionTable = @{}
        $Counter = 1
        foreach ($ID in $UniqueSessionIDs) {
            $SessionTable[$ID] = $Counter
            $Counter++
        }
    }

    process {

        # create custom table to display logs.
        $CorrelationLogs = $Logs | Select-Object `
        @{Name = 'CreatedDateTime'; Expression = { $_.CreatedDateTime.ToLocalTime() } },
        @{Name = 'ErrorCode'; Expression = { $_.Status.ErrorCode } },
        @{Name = 'IpAddress'; Expression = { $_.IpAddress } },
        @{Name = 'City'; Expression = { $_.Location.City } },
        @{Name = 'State'; Expression = { $_.Location.State } },
        @{Name = 'Country'; Expression = { $_.Location.CountryOrRegion } },
        @{Name = 'AppDisplayName'; Expression = { $_.AppDisplayName } },
        @{Name = 'Browser'; Expression = { $_.DeviceDetail.Browser } },
        @{Name = 'OperatingSystem'; Expression = { $_.DeviceDetail.OperatingSystem } },
        @{Name = 'TrustType'; Expression = { $_.DeviceDetail.TrustType } },
        @{Name = 'CorrelationId'; Expression = { $SessionTable[$_.CorrelationId] } } | Sort-Object CorrelationId

        $CorrelationIds = $CorrelationLogs.CorrelationId | Sort-Object -Unique

        # show groups of logins
        foreach ( $CorrelationID in $CorrelationIds ) { 
            $IsolatedLogs = $CorrelationLogs | Where-Object { $_.CorrelationId -eq $CorrelationId }
            if ( $IsolatedLogs.Count -ge 2) {
                $IsolatedLogs | Format-Table * -AutoSize

                $SessionWithMultipleSignins = $true
            }
        }

        # write message if no sessions found with multiple logins
        if ( -not $SessionWithMultipleSignins ) {

            Write-Host -ForeGroundColor Red "`nNo sessions found with multiple logins."
        }
    }
}