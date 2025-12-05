function Compare-ServicePrincipals {
    <#
	.SYNOPSIS
	
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
    )

    begin {

        # variables
        $AllContent = [System.Collections.Generic.List[pscustomobject]]::new()
        $OutputTable = [System.Collections.Generic.List[pscustomobject]]::new()

        $ModulePath = $PSScriptRoot
        $Folder = Join-Path -Path $ModulePath -ChildPath "\client_serviceprincipals\"
        $Files = Get-ChildItem -Path $Folder -File
        $GroupProperties = @(
            'DisplayName'
            'AppId'
            'AppOwnerOrganizationId'
        )

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
    }

    process {

        Write-Host @Blue "Importing $( @($Files).Count) files."

        # import content from all files
        foreach ( $File in $Files ) {
            $FileContent = Import-Csv -Path $File.FullName
            foreach ( $Line in $FileContent ) {
                $AllContent.Add( $Line )
            }
        }

        # group
        $Grouped = $AllContent | Group-Object -Property $GroupProperties | Sort-Object Count -Descending

        # loop through groups, create output
        foreach ( $Group in $Grouped ) {

            $MostRecent = $Group.Group | Sort-Object -Property LastFound -Descending | Select-Object -First 1

            $OutputTable.Add( [pscustomobject]@{
                    Count                  = $Group.Count
                    LastFound              = $MostRecent.LastFound
                    DisplayName            = $MostRecent.DisplayName
                    AppId                  = $MostRecent.AppId
                    AppOwnerOrganizationId = $MostRecent.AppOwnerOrganizationId
                } )
        }
        
        # export to new file
        $OutputTable | Export-Csv -Path 'TEST.csv' -Force
    }
}