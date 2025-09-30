New-Alias -Name 'FindInteresting' -Value 'Find-InterestingUsers' -Force

function Find-InterestingUsers {
	<#
	.SYNOPSIS
	Finds admins plus users with interesting job titles or descriptions.
	
	.NOTES
	Version: 1.0.1
	#>
    [CmdletBinding()]
    param (
    )

    begin {

        # variables
        $MatchingUsers = [System.Collections.Generic.List[pscustomobject]]::new()

        $GetProperties = @(
            'AccountEnabled'
            'Activities'
            'EmployeeType'
            'DisplayName'
            'GivenName'
            'Id'
            'JobTitle'
            'JoinedTeams'
            'Mail'
            'ManagedAppRegistrations'
            'OfficeLocation'
            'ProxyAddresses'
            'Surname'
            'UserPrincipalName'
        )
        $Users = Get-MgUser -All -Property $GetProperties

        $Keywords = @(
            'admin'
            'ceo'
            'cfo'
            'cio'
            'ciso'
            'president'
        )

        $Cyan = @{
            ForegroundColor = 'Cyan'
        }
	}

    process {

        Write-Host @Cyan "`nSearching for keywords: ${Keywords}"

        foreach ( $User in $Users ) {

            $MatchingProperties = [System.Collections.Generic.List[pscustomobject]]::new()

            foreach ( $Property in $GetProperties ) {

                foreach ( $Keyword in $Keywords ) {

                    if ( $User.$Property -match $Keyword ) {
                        $MatchingProperties.Add( $Property )
                    }
                }
            }

            if ( $MatchingProperties ) {
                $Custom =  [pscustomobject]@{
                    AccountEnabled = $User.AccountEnabled
                    DisplayName = $User.DisplayName
                    UserPrincipalName = $User.UserPrincipalName
                }
                $Param = @{
                    MemberType = 'NoteProperty'
                    ErrorAction = 'SilentlyContinue'
                }
                foreach ( $Property in $MatchingProperties ) {
                    $Custom | Add-Member @Param -Name $Property -Value $User.$Property
                }

                $MatchingUsers.Add( $Custom )
            }
        }

        # display any matching users
        if ( $MatchingUsers ) {
            $MatchingUsers | Format-Table
        }
        else {
            Write-Host "No interesting users found."
        }
    }
}