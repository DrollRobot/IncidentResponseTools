function Show-SignInLog {
    [CmdletBinding()]
    param (
        [Parameter( Mandatory, Position = 0 )]
        [PSCustomObject] $Log
    )

    begin {

        # variables
        $Properties = @(
            'AppliedConditionalAccessPolicies'
            'AuthenticationAppDeviceDetails'
            'AuthenticationDetails'
            'AuthenticationProcessingDetails'
            'AuthenticationRequirementPolicies'
            'DeviceDetail'
            'Location'
            'ManagedServiceIdentity'
            'MfaDetail'
            'PrivateLinkDetails'
            'SessionLifetimePolicies'
            'Status'
            'TokenProtectionStatusDetails'
            'AdditionalProperties'
        )
        $Blue = @{ ForegroundColor = 'Blue' }
    }

    process {
    
        # check if custom object
        if ( $Log -is [PSCustomObject] ) {

            Write-Host @Blue "`nShowing: `$Log"
            $Log | Format-List *

            foreach ( $Property in $Properties ) {

                if ( $null -ne $Property ) {
                    Write-Host @Blue "`nShowing: `$Log.$Property"
                    $Log.$Property | Format-List *
                }
            }
        } 
        # if not custom object, display error
        else {
            Write-Error "`nWarning: Input object is not a PSCustomObject."
        }
    }
}