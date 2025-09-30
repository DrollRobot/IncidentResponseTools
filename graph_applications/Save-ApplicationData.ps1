function Save-ApplicationData {
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
        $NewApplications = [System.Collections.Generic.List[pscustomobject]]::new()
        $ModulePath = $PSScriptRoot
        $DateString = ( Get-Date ).ToString("o")

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

    }

    process {

        Write-Host @Blue "`nSaving application and service principal information..."

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        # get application information
        $Applications = Get-MgApplication -All

        # build file name
        $FileName = "${DomainName}_applications.csv"
        $FilePath = Join-Path -Path $ModulePath -ChildPath "\client_applications\${FileName}"

        # import file, if it exists
        $ResolveParams = @{
            Path        = $FilePath
            ErrorAction = 'SilentlyContinue'
        }
        $ExistingFile = ( Resolve-Path @ResolveParams ).Path
        if ( $ExistingFile ) {
            $OldApplications = Import-Csv -Path $ExistingFile
        }
        else {
            $OldApplications = @()
        }

        # make new application object
        foreach ( $Application in $Applications ) {

            $NewApplications.Add(
                [pscustomobject]@{
                    Type                   = 'Application'
                    LastFound              = $DateString
                    DisplayName            = $Application.DisplayName
                    AppId                  = $Application.AppId
                    PublisherDomain        = $Application.PublisherDomain
                    CreatedDateTime        = $Application.CreatedDateTime
                    Id                     = $Application.Id
                    SignInAudience         = $Application.SignInAudience
                    RequiredResourceAccess = $Application.RequiredResourceAccess.ResourceAppId -join ";"
                }
            )
        }

        ### add any old applications that aren't already in new applications
        # loop through old applications
        foreach ( $OldApplication in $OldApplications ) {

            # variables
            $AddToNewApplications = $true

            # if application matches, set variable to false
            foreach ( $NewApplication in $NewApplications ) {
                if ( $OldApplication.DisplayName -eq $NewApplication.DisplayName -and
                    $OldApplication.AppId -eq $NewApplication.AppId -and
                    $OldApplication.PublisherDomain -eq $NewApplication.PublisherDomain
                ) {
                    $AddToNewApplications = $false
                }
            }
            

            # if variable is true, add to new applications
            if ( $AddToNewApplications ) {
                $NewApplications.Add( $OldApplication ) 
            }
        }

        # save updated data to file
        $NewApplications | Export-Csv -Path $FilePath -NoTypeInformation -Force
    }
}