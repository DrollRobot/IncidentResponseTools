New-Alias -Name 'OpenMailbox' -Value 'Open-MailboxInOWA' -Force
function Open-MailboxInOWA {
    <#
	.SYNOPSIS
	Opens user mailbox in OWA in a browser.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,

        [ValidateSet( 'msedge','chrome','firefox','brave','default' )]
        [string] $Browser = 'default',

        [switch] $Private
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
            throw "Not connected to ExchangeOnlineManagement. Run Connect-ExchangeOnline. Exiting."
        }
     
        # variables

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
            $MailboxUrl = "https://outlook.office.com/mail/${UserEmail}/?offline=disabled"

            Write-Host @Blue "Opening ${UserEmail}'s mailbox in web browser." | Out-Host
            $Params = @{
                Browser = $Browser
                Url = $MailboxUrl
            }
            if ( $Private ) {
                $Params['Private'] = $true
            }
            Open-Browser @Params
        }
    }
}



