New-Alias -Name 'GeoBlock' -Value 'Show-GraphGeoBlockPolicy' -Force
function Show-GraphGeoBlockPolicy {
    param(
        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        # $Function = $MyInvocation.MyCommand.Name
        # $ParameterSet = $PSCmdlet.ParameterSetName
        if ($Test -or $Script:Test) {
            $Script:Test = $true
            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }
        $GroupDisplayProperties = @(
            'DisplayName'
            'Id'
        )
        $UserProperties = @(
            'AccountEnabled'
            'DisplayName'
            'UserPrincipalName'
            'OnPremisesSamAccountName'
            'Id'
        )
        $GeoblockPatterns = @(
            'geo.?block'
            'Block Access from Non-US'
            'Block Untrusted Locations'
            'Block Connections Outside U[.]?S[.]?'
        )
        $GeoblockPattern = $GeoBlockPatterns -join '|'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        $CAPolicies = Get-MgIdentityConditionalAccessPolicy -All
        $Users = Request-GraphUsers
        $Groups = Request-GraphGroups
    }

    process {

        # display all CA policies
        $SortedPolicies = $CAPolicies | Sort-Object -Property `
        @{Expression = "State"; Descending = $true },
        @{Expression = "DisplayName"; Descending = $false }
        Write-Host @Blue "`nShowing all Conditional Access policies." | Out-Host
        $SortedPolicies | Format-Table State, DisplayName | Out-Host

        # display possible geoblocking policies
        $GeoBlockPolicies = $CAPolicies | Where-Object { $_.DisplayName -match $GeoblockPattern }
        Write-Host @Blue "`nShowing possible GeoBlocking policies." | Out-Host
        if ( @( $GeoBlockPolicies ).Count -gt 0 ) {
            $GeoBlockPolicies | Format-Table State, DisplayName | Out-Host
        }
        else {
            Write-Host "None found." | Out-Host
        }

        foreach ( $Policy in $GeoBlockPolicies ) {

            $PolicyName = $Policy.DisplayName

            Write-Host @Cyan "`nShowing policy '${PolicyName}'" | Out-Host

            # show included users
            $IncludeUserIds = $Policy.Conditions.Users.IncludeUsers
            Write-Host @Blue "`nINCLUDED users:" | Out-Host
            if ( $IncludeUserIds -and @( $IncludeUserIds ).Count -gt 0 ) {
                if ( $IncludeUserIds -eq 'All' ) {
                    Write-Host 'All' | Out-Host
                }
                else {
                    if ( @( $IncludeUserIds ).Count -le 10 ) {
                        $Users | 
                            Where-Object { $_.Id -in $IncludeUserIds } |
                            Format-Table $UserProperties |
                            Out-Host
                    }
                    else {
                        $AllUserCount = @( $Users ).Count
                        $IncludeUsersCount = @( $IncludeUserIds ).Count
                        Write-Host "${IncludeUsersCount} of ${AllUserCount} users included" | Out-Host
                    }
                }
            }
            else {
                Write-Host "None" | Out-Host
            }
            
            # show included Groups
            $IncludeGroupIds = $Policy.Conditions.Users.IncludeGroups
            $IncludeGroups = $Groups | Where-Object { $_.Id -in $IncludeGroupIds }
            Write-Host @Blue "`nINCLUDED groups:" | Out-Host
            if ( $IncludeGroupIds -and @( $IncludeGroupIds ).Count -gt 0 ) {
                $IncludeGroups |
                    Format-Table $GroupDisplayProperties |
                    Out-Host
            }
            else {
                Write-Host "None" | Out-Host
            }

            # show excluded users
            $ExcludeUserIds = $Policy.Conditions.Users.ExcludeUsers
            Write-Host @Blue "`nEXCLUDED users:" | Out-Host
            if ( @( $ExcludeUserIds -and $ExcludeUserIds ).Count -gt 0 ) {
                $Users |
                    Where-Object { $_.Id -in $ExcludeUserIds } |
                    Format-Table $UserProperties |
                    Out-Host
            }
            else {
                Write-Host "None" | Out-Host
            }


            # show excluded groups
            $ExcludeGroupIds = $Policy.Conditions.Users.ExcludeGroups
            $ExcludeGroups = $Groups | Where-Object { $_.Id -in $ExcludeGroupIds }
            Write-Host @Blue "`nEXCLUDED groups:" | Out-Host
            if ( $ExcludeGroupIds -and @( $ExcludeGroupIds ).Count -gt 0 ) {
                $ExcludeGroups |
                    Format-Table $GroupDisplayProperties |
                    Out-Host
            }
            else {
                Write-Host "None" | Out-Host
            }

            # show members of exclude groups
            foreach ( $ExcludeGroup in $ExcludeGroups ) {

                $GroupName = $ExcludeGroup.DisplayName
                $GroupMemberIds = ( Get-MgGroupMember -GroupId $ExcludeGroup.Id ).Id

                Write-Host @Blue "`nUsers in group '${GroupName}':" | Out-Host
                if ( $GroupMemberIds -and @( $GroupMemberIds ).Count -gt 0 ) {
                    $Users |
                        Where-Object { $_.Id -in $GroupMemberIds } |
                        Format-Table $UserProperties |
                        Out-Host
                }
                else {
                    Write-Host "None"
                }
            }
        }
    }
}