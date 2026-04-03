# New-Alias -Name 'Connect' -Value 'Connect-IncidentResponseTools' -Force
New-Alias -Name 'ConnectIRT' -Value 'Connect-IncidentResponseTools' -Force
New-Alias -Name 'Connect-IRT' -Value 'Connect-IncidentResponseTools' -Force

function Connect-IncidentResponseTools {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph and Exchange Online for incident response.

    .DESCRIPTION
    Orchestrates connections to Graph and Exchange Online.
    When no service switches are specified, both services are connected. Use -Graph
    or -Exchange to connect to specific services only.

    .PARAMETER TenantId
    The TenantId GUID for the environment you want to connect to.

    .PARAMETER UserPrincipalName
    The UserPrincipalName (Email) for the user account to connect with.
    Required when connecting to Exchange.

    .PARAMETER GCCHigh
    Connect to a GCC High tenant environment.

    .PARAMETER DeviceCode
    Use device code authentication flow instead of interactive browser auth.

    .PARAMETER AdditionalScopes
    Additional Graph scopes to request beyond the default set.

    .PARAMETER Graph
    Connect to Microsoft Graph only.

    .PARAMETER Exchange
    Connect to Exchange Online only.

    .PARAMETER Browser
    Browser to use for device code login and URL opening. Valid values: msedge, chrome, firefox, brave, default.

    .PARAMETER Private
    Open the browser in private/incognito mode.

    .EXAMPLE
    Connect-IncidentResponseTools -TenantId $tid -UserPrincipalName admin@contoso.com
    Connects to Graph and Exchange Online.

    .EXAMPLE
    Connect-IncidentResponseTools -TenantId $tid -Graph -DeviceCode
    Connects to Graph only using device code auth.

    .EXAMPLE
    Connect-IncidentResponseTools -TenantId $tid -UserPrincipalName admin@contoso.com -Exchange -GCCHigh
    Connects to Exchange in a GCC High environment.

    .NOTES
    Version: 1.0.0
    #>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory )]
        [string] $TenantId,

        [string] $UserPrincipalName,

        [switch] $GCCHigh,

        [switch] $DeviceCode,

        [string[]] $AdditionalScopes,

        [switch] $Graph,

        [switch] $Exchange,

        [ValidateSet('msedge','chrome','firefox','brave','default')]
        [string] $Browser = 'default',

        [switch] $Private
    )

    process {

        # if no service switches specified, connect to both
        $ConnectAll = -not ( $Graph -or $Exchange )

        $ConnectGraph    = $ConnectAll -or $Graph
        $ConnectExchange = $ConnectAll -or $Exchange

        # validate UPN is provided when Exchange is requested
        if ( $ConnectExchange -and -not $UserPrincipalName ) {
            throw 'UserPrincipalName is required when connecting to Exchange.'
        }

        # --- Graph ---
        if ( $ConnectGraph ) {

            $GraphParams = @{
                TenantId = $TenantId
            }

            if ( $GCCHigh )    { $GraphParams['GCCHigh']    = $true }
            if ( $DeviceCode ) { $GraphParams['DeviceCode'] = $true }

            $GraphParams['Browser'] = $Browser
            if ( $Private ) { $GraphParams['Private'] = $true }

            if ( $AdditionalScopes ) {
                $GraphParams['AdditionalScopes'] = $AdditionalScopes
            }

            Connect-IRTGraph @GraphParams
        }

        # --- Exchange Online ---
        if ( $ConnectExchange ) {

            $ExchangeParams = @{
                TenantId          = $TenantId
                UserPrincipalName = $UserPrincipalName
            }

            if ( $GCCHigh )    { $ExchangeParams['GCCHigh']    = $true }
            if ( $DeviceCode ) { $ExchangeParams['DeviceCode'] = $true }

            $ExchangeParams['Browser'] = $Browser
            if ( $Private ) { $ExchangeParams['Private'] = $true }

            Connect-IRTExchange @ExchangeParams
        }

        # --- Update prompt to show connected services ---
        if ( -not $Global:IRT_OriginalPrompt ) {
            $Global:IRT_OriginalPrompt = $function:prompt
        }

        function Global:prompt {
            $parts = @()

            $GraphCtx = Get-MgContext -ErrorAction SilentlyContinue
            if ( $GraphCtx -and $GraphCtx.Account ) {
                $graphDomain = ($GraphCtx.Account -split '@')[-1]
                $parts += "Graph:$graphDomain"
            }

            $ExoConn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
            if ( $ExoConn -and $ExoConn.UserPrincipalName ) {
                $exoDomain = ($ExoConn.UserPrincipalName -split '@')[-1]
                $parts += "Exchange:$exoDomain"
            }

            if ( $parts.Count -gt 0 ) {
                Write-Host "[IRT] $( $parts -join '; ' )" -NoNewline -ForegroundColor Cyan
                Write-Host ' ' -NoNewline
            }

            & $Global:IRT_OriginalPrompt
        }
    }
}
