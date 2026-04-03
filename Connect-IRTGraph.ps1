function Connect-IRTGraph {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph with default incident response scopes.

    .PARAMETER TenantId
    The TenantId GUID for the environment you want to connect to.

    .PARAMETER GCCHigh
    Connect to a GCC High tenant environment.

    .PARAMETER DeviceCode
    Use device code authentication flow instead of interactive browser auth.

    .PARAMETER AdditionalScopes
    Additional Graph scopes to request beyond the default set.

    .PARAMETER Browser
    Browser to use for device code login. Valid values: msedge, chrome, firefox, brave, default.

    .PARAMETER Private
    Open the browser in private/incognito mode.

    .NOTES
    Version: 1.0.0
    #>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory )]
        [string] $TenantId,

        [switch] $GCCHigh,

        [switch] $DeviceCode,

        [string[]] $AdditionalScopes,

        [ValidateSet('msedge','chrome','firefox','brave','default')]
        [string] $Browser = 'default',

        [switch] $Private
    )

    process {

        $DefaultScopes = @(
            'Application.ReadWrite.All'
            'AuditLog.Read.All'
            'AuditLogsQuery.Read.All'
            'BitLockerKey.Read.All'
            'DelegatedPermissionGrant.ReadWrite.All'
            'Device.ReadWrite.All'
            'DeviceLocalCredential.Read.All'
            'DeviceManagementApps.ReadWrite.All'
            'DeviceManagementConfiguration.ReadWrite.All'
            'DeviceManagementManagedDevices.ReadWrite.All'
            'DeviceManagementServiceConfig.ReadWrite.All'
            'Directory.AccessAsUser.All'
            'Directory.ReadWrite.All'
            'Domain.Read.All'
            'Group.ReadWrite.All'
            'GroupMember.ReadWrite.All'
            'IdentityRiskEvent.ReadWrite.All'
            'IdentityRiskyServicePrincipal.ReadWrite.All'
            'IdentityRiskyUser.ReadWrite.All'
            'Mail.ReadBasic.Shared'
            'Organization.Read.All'
            'Policy.Read.All'
            'Policy.Read.ConditionalAccess'
            'Policy.ReadWrite.Authorization'
            'RoleManagement.ReadWrite.Directory'
            'SecurityEvents.ReadWrite.All'
            'SecurityIncident.ReadWrite.All'
            'User-Mail.ReadWrite.All'
            'User-PasswordProfile.ReadWrite.All'
            'User-Phone.ReadWrite.All'
            'User.EnableDisableAccount.All'
            'User.ManageIdentities.All'
            'User.ReadWrite.All'
            'User.RevokeSessions.All'
            'UserAuthenticationMethod.ReadWrite'
            'UserAuthenticationMethod.ReadWrite.All'
            'UserAuthMethod-Passkey.ReadWrite.All'
        )

        $Scopes = $DefaultScopes
        if ( $AdditionalScopes ) {
            $Scopes = $DefaultScopes + $AdditionalScopes | Select-Object -Unique
        }

        $ConnectParams = @{
            TenantId = $TenantId
            Scopes   = $Scopes
            NoWelcome = $true
            ContextScope = 'Process'
        }

        if ( $GCCHigh ) {
            $ConnectParams['Environment'] = 'USGov'
        }

        if ( $DeviceCode ) {
            $ConnectParams['UseDeviceCode'] = $true
        }

        # check if already connected to this tenant
        $ExistingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ( $ExistingContext -and $ExistingContext.TenantId -eq $TenantId ) {
            Write-Host "Already connected to Microsoft Graph for tenant $TenantId." -ForegroundColor Yellow
            return
        }

        if ($DeviceCode) {
            # pipe information stream through ForEach-Object so each record is
            # processed as it arrives, before Connect-MgGraph finishes blocking
            Connect-MgGraph @ConnectParams 6>&1 | ForEach-Object {
                if ($_ -match 'enter the code\s+(\S+)') {
                    $Matches[1] | Set-Clipboard
                    Write-Host "Device code '$( $Matches[1] )' copied to clipboard." -ForegroundColor Green
                    Open-Browser -Browser $Browser -Url 'https://microsoft.com/devicelogin' -Private:$Private
                }
            }
        }
        else {
            Connect-MgGraph @ConnectParams
        }
    }
}
