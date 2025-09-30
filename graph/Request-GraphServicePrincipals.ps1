function Request-GraphServicePrincipals {
    <#
	.SYNOPSIS
    Gets serviceprincipal information from local file, or from graph if local file doesn't exist
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
    )

    begin {

        # variables
        $CurrentPath = Get-Location
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileNameDate = ( Get-Date ).ToString( $FileNameDateFormat )
        $GetProperties = @(
            'AdditionalProperties'
            'AppDescription'
            'AppId'
            'AppOwnerOrganizationId'
            'Description'
            'DisplayName'
            'Id'
        )

        # get client domain name
        $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1

        $String = "ServicePrincipals_Raw_${DomainName}"
        $FilterString = "${String}_*.xml"
        $FileName = "${String}_${FileNameDate}.xml"
    }

    process {

        # get files in current directory that match pattern
        $Files = Get-ChildItem -Filter $FilterString
        if ( $Files ) {
            $File = $Files | Sort-Object 'LastWriteTime' -Descending | Select-Object -First 1
            $Objects = Import-CliXml -Path $File.FullName
        }
        else {
            $Objects = Get-MgServicePrincipal -All | Select-Object $GetProperties
            $XmlOutputPath = Join-Path -Path $CurrentPath -ChildPath $FileName
            $Objects | Export-Clixml -Depth 5 -Path $XmlOutputPath
        }

        return $Objects
    }
}