function Save-AllAppData {
    <#
	.SYNOPSIS
	Runs both Save-ApplicationData and Save-ServicePrincipalData
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
    )

    Save-ApplicationData
    Save-ServicePrincipalData
}
