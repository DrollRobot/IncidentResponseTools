New-Alias -Name 'FindUser' -Value 'Find-Users' -Force
New-Alias -Name 'FindUsers' -Value 'Find-Users' -Force
New-Alias -Name 'Find-User' -Value 'Find-Users' -Force
function Find-Users {
    <#
    .SYNOPSIS
    Finds graph user by displayname, email address, or user id guid. Creates $UserObject or $UserObjects variables.

    .EXAMPLE
    Find-Users flast
    Find-Users -Search flast,flast,flast
    Find-Users flast@domain.com
    Find-Users -Search bf7573a5844f (partial user id number)

    .NOTES
    Version: 1.1.4
    1.1.4 - Fixed bug with $UserObjects not being a collection. Moved getting full object to Show-User function.
    1.1.3 - Removed checks for modules and permissions. Checking at module level instead.
    1.1.2 - Added enabled as a displayed field.
    1.1.1 - Bug fix. Script was passing collections rather than user objects.
    1.1.0 - Major rewrite. Renamed to Find-Users.
    #>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0, Mandatory )]
        [string[]] $Search,
        [string] $VarPrefix,
        [switch] $Script,
        [string] $TenantId
    )

    begin {

        # variables
        $ScriptUserObjects = [System.Collections.Generic.List[PsObject]]::new()
        $GetProperties = @(
            'AccountEnabled'
            'DisplayName'
            'Id'
            'OnPremisesSamAccountName'
            'ProxyAddresses'
            'UserPrincipalName'
        )
        $DisplayProperties = @(
            'AccountEnabled'
            'DisplayName'
            'UserPrincipalName'
            'OnPremisesSamAccountName'
            'Id'
        )

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # get all users
        $GraphUsers = Get-MgUser -All -Property $GetProperties
    }

    process {

        Write-Host ''

        foreach ( $SearchString in $Search ) {

            # find matching users
            $MatchingUsers = $GraphUsers | Where-Object {
                $_.DisplayName -match $SearchString -or
                $_.UserPrincipalName -match $SearchString -or
                $_.Id -match $SearchString -or
                $_.ProxyAddresses -match $SearchString -or
                $_.OnPremisesSamAccountName -match $SearchString
            }

            if (($MatchingUsers | Measure-Object).Count -eq 1) {
                
                if ( -not $Script ) {

                    # show user info
                    Write-Host @Blue "Showing results for search: ${SearchString}"
                    $MatchingUsers | Format-Table $DisplayProperties
                }

                # add user to array
                $ScriptUserObjects.Add( ( $MatchingUsers | Select-Object -First 1 ) )
            }
            elseif (($MatchingUsers | Measure-Object).Count -gt 1) {

                if ( -not $Script ) {

                    # show user info
                    Write-Host @Blue "Showing results for search: ${SearchString}"
                    $MatchingUsers | Format-Table $DisplayProperties
                    Write-Host @Red 'Multiple users found. Refine search.'
                }
            }
            else {
                if ( -not $Script ) {
                    Write-Host @Red "$SearchString not found. Try a different search."
                }
            }
        }

        # if script, just return objects
        if ($Script) {
            return @($ScriptUserObjects)
        }

        # if one user
        if ( ($ScriptUserObjects | Measure-Object).Count -eq 1) {

            # set objects
            $VariableParams = @{
                Name  = "${VarPrefix}UserObject"
                Value = $ScriptUserObjects[0]
                Scope = 'Global'
                Force = $true
            }
            New-Variable @VariableParams
            $VariableParams = @{
                Name  = "${VarPrefix}UserObjects"
                Value = @($ScriptUserObject)
                Scope = 'Global'
                Force = $true
            }
            New-Variable @VariableParams
            $VariableParams = @{
                Name  = "${VarPrefix}UserEmail"
                Value = $ScriptUserObject.UserPrincipalName
                Scope = 'Global'
                Force = $true
            }
            New-Variable @VariableParams
            Write-Host @Blue "`nCreated `$${VarPrefix}UserObject, `$${VarPrefix}UserObjects, and `$${VarPrefix}UserEmail"
        }
        elseif (($ScriptUserObjects | Measure-Object).Count -gt 1) {
            
            # set objects
            $VariableParams = @{
                Name  = "${VarPrefix}UserObjects"
                Value = @($ScriptUserObjects)
                Scope = 'Global'
                Force = $true
            }
            New-Variable @VariableParams
            Write-Host @Blue "`nCreated `$${VarPrefix}UserObjects"
            $ScriptUserObjects | Format-Table $DisplayProperties
        }        
    }
}


