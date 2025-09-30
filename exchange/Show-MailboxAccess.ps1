New-Alias -Name 'ShowAccess' -Value 'Show-MailboxAccess' -Force
New-Alias -Name 'MailboxAccess' -Value 'Show-MailboxAccess' -Force
function Show-MailboxAccess {
    <#
	.SYNOPSIS
	Grants the currently logged in user full access to the target user's mailbox.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects
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
            $Exchange = Get-ConnectionInformation
        }
        catch {}
        if ( -not $Exchange ) {
            throw "Not connected to ExchangeOnlineManagement. Run Connect-ExchangeOnline."
        }
     
        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
    }

    process {

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            $UserEmail = $ScriptUserObject.UserPrincipalName

            # show users who have access to target mailbox
            Write-Host @Blue "Showing users who have access to ${UserEmail}" | Out-Host
            $Properties = @(
                'User'
                'AccessRights'
                'IsInherited'
                'InheritanceType'
            )
            $MailboxPermissions = Get-MailboxPermission -Identity $UserEmail
            $MailboxPermissions | Format-Table $Properties -AutoSize
        }
    }
}