New-Alias -Name 'FullAccess' -Value 'Grant-MailboxFullAccess' -Force
function Grant-MailboxFullAccess {
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
        [psobject[]] $UserObjects,

        [string] $GrantAccessTo,

        [switch] $Remove
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
    }

    process {

        #region CURRENT USER

        if ( -not $GrantAccessTo ) {
            $AccountsList = [System.Collections.Generic.List[string]]::new()
            try {
                $Accounts = @((Get-ConnectionInformation ).UserPrincipalName)
            }
            catch {
                $_
                throw "Unable to retrieve connected Exchange account."
            }

            # remove empty entries
            $Accounts = $Accounts | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}

            # add to list
            foreach ($Object in $Accounts) { 
                $AccountsList.Add($Object)
            }
        }

        if ( $AccountsList.Count -lt 1 ) {
            throw "Must specify -GrantAccessTo"
        }
        elseif ( $AccountsList.Count -gt 1 ) {

            # remove duplicates
            $HashSet = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($Object in $AccountsList) { $HashSet.Add($Object) | Out-Null }
            $AccountsList = @($HashSet)

            # if more than one option, have user choose
            if ( $AccountsList.Count -gt 1 ) {
                $MenuParams = @{
                    Title = "Choose account to receive full access to mailbox."
                    Options = $AccountsList
                    List = $true
                }
                $GrantAccessTo = Build-Menu @MenuParams
            }
            else {
               $GrantAccessTo = $AccountsList | Select-Object -First 1
            }
        }
        else {
            $GrantAccessTo = $AccountsList
        }

        #region USER LOOP

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            $UserEmail = $ScriptUserObject.UserPrincipalName

            if ( $Remove ) {

                # remove access
                Write-Host @Blue "Removing access to ${UserEmail} from ${GrantAccessTo}" | Out-Host
                $Params = @{
                    Identity = $UserEmail
                    User = $GrantAccessTo
                    AccessRights = 'FullAccess'
                    Confirm = $false
                }
                Remove-MailboxPermission @Params | Out-Null
            }
            else {

                # add access
                Write-Host @Blue "Adding access to ${UserEmail} to ${GrantAccessTo}" | Out-Host
                $Params = @{
                    Identity = $UserEmail
                    User = $GrantAccessTo
                    AccessRights = 'FullAccess'
                    InheritanceType = 'All'
                }
                Add-MailboxPermission @Params | Out-Null
            }

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