function Show-CustomObjects {
    <#
	.SYNOPSIS
	
	
	.NOTES
	Version: 1.0.1
    1.0.1 - Updated output formatting.
	#>
    [CmdletBinding()]
    param (
        [pscustomobject[]] $CustomObjects
    )

    begin {

        # variables

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
    }

    process {

        # users
        Write-Host @Blue "`nUsers with admin roles:"
        $Users = $CustomObjects |
            Where-Object { $_.ObjectType -eq 'User' } |
            Sort-Object AccountEnabled -Descending
        if ( $Users ) {
            $Users | Format-Table -AutoSize $UserDisplayProperties | Out-Host
        }
        else {
            Write-Host "None" | Out-Host
        }
    
        # service principals
        Write-Host @Blue "`nService Principals with admin roles:"
        $ServicePrincipals = $CustomObjects |
            Where-Object { $_.ObjectType -eq 'ServicePrincipal' } |
            Sort-Object AccountEnabled -Descending
        if ( $ServicePrincipals ) {
            $ServicePrincipals | Format-Table -AutoSize $ServicePrincipalDisplayProperties | Out-Host
        }
        else {
            Write-Host "None" | Out-Host
        }
    
        # groups
        Write-Host @Blue "`nGroups with admin roles:"
        $Groups = $CustomObjects |
            Where-Object { $_.ObjectType -eq 'Group' }
        if ( $Groups ) {
            $Groups | Format-Table -AutoSize $GroupDisplayProperties | Out-Host
        }
        else {
            Write-Host "None`n" | Out-Host
        }
    }
}