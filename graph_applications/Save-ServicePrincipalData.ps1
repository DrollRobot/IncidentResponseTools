function Save-ServicePrincipalData {
    <#
	.SYNOPSIS
	Collects information about applications registered in a tenant.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
    )

    begin {

        # variables
        $NewServicePrincipals = [System.Collections.Generic.List[pscustomobject]]::new()
        $ModulePath = $PSScriptRoot
        $DateString = ( Get-Date ).ToString("o")

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        # get application information
        $ServicePrincipals = Get-MgServicePrincipal -All

        # build file name
        $FileName = "${DomainName}_serviceprincipals.csv"
        $FilePath = Join-Path -Path $ModulePath -ChildPath "\known_servicePrincipals\${FileName}"

        # import file, if it exists
        $ResolveParams = @{
            Path        = $FilePath
            ErrorAction = 'SilentlyContinue'
        }
        $ExistingFile = ( Resolve-Path @ResolveParams ).Path
        if ( $ExistingFile ) {
            $OldServicePrincipals = Import-Csv -Path $ExistingFile
        }
        else {
            $OldServicePrincipals = @()
        }
    }

    process {

        # make new application object
        foreach ( $ServicePrincipal in $ServicePrincipals ) {

            $NewServicePrincipals.Add(
                [pscustomobject]@{
                    Type                   = 'ServicePrincipal'
                    LastFound              = $DateString
                    DisplayName            = $ServicePrincipal.DisplayName
                    AppId                  = $ServicePrincipal.AppId
                    AppOwnerOrganizationId = $ServicePrincipal.AppOwnerOrganizationId
                    CreatedDateTime        = $ServicePrincipal.AdditionalProperties.createdDateTime
                    Id                     = $ServicePrincipal.Id
                    SignInAudience         = $ServicePrincipal.SignInAudience
                }
            )
        }

        ### add any old applications that aren't already in new applications
        # loop through old applications
        foreach ( $OldServicePrincipal in $OldServicePrincipals ) {

            # variables
            $AddToNewServicePrincipals = $true

            # if application matches, set variable to false
            foreach ( $NewServicePrincipal in $NewServicePrincipals ) {
                if ( $OldServicePrincipal.DisplayName -eq $NewServicePrincipal.DisplayName -and
                    $OldServicePrincipal.AppId -eq $NewServicePrincipal.AppId -and
                    $OldServicePrincipal.AppOwnerOrganizationId -eq $NewServicePrincipal.AppOwnerOrganizationId
                ) {
                    $AddToNewServicePrincipals = $false
                }
            }

            # if variable is true, add to new applications
            if ( $AddToNewServicePrincipals ) {
                $NewServicePrincipals.Add( $OldServicePrincipal ) 
            }
        }

        # save updated data to file
        $NewServicePrincipals | Export-Csv -Path $FilePath -Force
    }
}