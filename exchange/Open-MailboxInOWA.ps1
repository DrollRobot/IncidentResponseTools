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

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

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



