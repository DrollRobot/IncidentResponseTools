New-Alias -Name 'ResetPassword' -Value 'Reset-GraphUserPasswords' -Force
New-Alias -Name 'ResetPasswords' -Value 'Reset-GraphUserPasswords' -Force
New-Alias -Name 'Reset-GraphUserPassword' -Value 'Reset-GraphUserPasswords' -Force
function Reset-GraphUserPasswords {
    <#
	.SYNOPSIS
	Resets Graph user password.	
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding( DefaultParameterSetName = 'RandomCharacters' )]
    param(
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,

        [Parameter( ParameterSetName = 'RandomCharacters' )]
        [Alias( 'Random' )]
        [switch] $RandomCharacters,

        # [Parameter( ParameterSetName = 'PassPhrase' )]
        # [Alias( 'Phrase' )]
        # [switch] $PassPhrase,

        [Parameter( ParameterSetName = 'Custom' )]
        [switch] $Custom,

        [string] $TenantId
    )

    begin {

        # if not passed directly, find global
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

        # variables
        if ( $PSCmdlet.ParameterSetName -eq 'RandomCharacters' ) {
            $RandomCharacters = $true
        }
        $GetProperties = @(
            'AccountEnabled'
            'DisplayName'
            'Id'
            'LastPasswordChangeDateTime'
            'OnPremisesSamAccountName'
            'OnPremisesSyncEnabled'
            'UserPrincipalName'
        )
        $DisplayProperties = @(
            'LastPasswordChangeDateTime'
            'AccountEnabled'
            'DisplayName'
            'OnPremisesSamAccountName'
            'UserPrincipalName'
        )

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
    }

    process {

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            switch ( $PSCmdlet.ParameterSetName ) {
                'Custom' { 
                    $Password = Read-Host -Prompt "Enter new password"
                }
                'RandomCharacters' {
                    $UserEmail = $ScriptUserObject.UserPrincipalName
                    $Password = Get-RandomPassword 30
                    Write-Host "${UserEmail} new password:`n${Password}"
                }
            }

            # create password profile and reset password
            $PasswordProfile = @{
                Password = $Password
                ForceChangePasswordNextSignIn = $false
                ForceChangePasswordNextSignInWithMfa = $false
            }
            Update-MgUser -UserId $UserObject.Id -PasswordProfile $PasswordProfile

            # get new user object
            Write-Host @Blue "`nGetting updated user information."
            $FullUserObject = Get-MgUser -UserId $ScriptUserObject.Id -Property $GetProperties
            try {
                $FullUserObject.LastPasswordChangeDateTime = $FullUserObject.LastPasswordChangeDateTime.ToLocalTime()
            }
            catch {}

            # display new object
            $FullUserObject | Format-Table $DisplayProperties

            # warn user if onpremsynced
            if ( $FullUserObject.OnPremisesSyncEnabled ) {
                Write-Host @Red "`nUser is synced from on-premises. Reset password in local AD too!"
            }
        }
    }
}
