function Find-UsersWithoutUsageLocation {
    param(
        [switch] $Excel,
        [switch] $Variable
    )

    begin {

        # variables
        $GetProperties = @(
            'AccountEnabled'
            'DisplayName'
            'Id'
            'UserPrincipalName'
            'UsageLocation'
            'UserType'
        )
        $DisplayProperties = @(
            'AccountEnabled'
            'UserType'
            'DisplayName'
            'UserPrincipalName'
            'Id'
        )
        $DateString = Get-Date -Format "yy-MM-dd"

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Red = @{ ForegroundColor = 'Red' }

        # get all users
        $Users = Get-MgUser -All -Property $GetProperties
    }

    process {
        
        # filter down to user without usagelocations
        $MatchingUsers = $Users | Where-Object {
            # $_.AccountEnabled -eq $true -and
            # $_.UserType -ne 'Guest' -and
            $null -eq $_.UsageLocation
        }
        $SortedUsers = $MatchingUsers | Sort-Object AccountEnabled,UserType,DisplayName
        $OutputTable = $SortedUsers | Select-Object $DisplayProperties
        
        # show users
        $OutputTable | Format-Table
        
        # export to excel
        if ( $Excel ) {
            $ExcelParams = @{
                Path          = "UsersWithoutUsageLocation_${DateString}.xlsx"
                WorkSheetname = "UsersWithoutUsageLocation"
                Title         = "Enabled User accounts without usage location set."
                TableStyle    = "Medium19"
                AutoSize      = $true
                FreezeTopRow  = $true
            }
            $OutputTable | Export-Excel @ExcelParams
        }
    
        if ( $Variable ) {
            Write-Host @Blue 'Setting $UserObjects global variable.'
            $Global:UserObjects = $MatchingUsers
            $null = $Global:UserObjects
        }
    }
}