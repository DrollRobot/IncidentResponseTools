function Get-ExchangeAccessToken {
    <#
	.SYNOPSIS
	Function for retrieving Exchange access token.
	
	.PARAMETER TenantId
	The TenantId GUID for the environment you want to connect to. 
		
	.PARAMETER UserPrincipalName
	The UserPrincipalName (Email) for the user account to connect with. 
	
	.NOTES
	Version: 1.2.0
	#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory )]
        [string] $TenantId,

        [string] $UserPrincipalName
    )

    begin {

        # show module version
        $CallStack = @( Get-PSCallStack )
        if ( $CallStack.Count -eq 2 ) {
            $ModuleVersion = $ExecutionContext.SessionState.Module.Version
            Write-Host "Module version: ${ModuleVersion}"
        }
    }

    process {

        # retrieve token
        $MsalParams = @{
            ClientId   = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
            TenantId   = $TenantId
            Scopes     = 'https://outlook.office365.com/.default'
            DeviceCode = $true
        }
        Start-Process "https://microsoft.com/devicelogin"
        $Token = ( Get-MsalToken @MsalParams ).AccessToken

        # connect to exchange in current session
        Connect-ExchangeOnline -AccessToken $Token -UserPrincipalName $UserPrincipalName -ShowBanner:$false

        # save token as global variable
        $Global:Exchange = [pscustomobject]@{
            Token = $Token
            UserPrincipalName = $UserPrincipalName
        }
        # prevent vscode error about unused variable
        $null = $Global:Exchange
    }
}