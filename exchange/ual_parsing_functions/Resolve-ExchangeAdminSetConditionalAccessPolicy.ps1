function Resolve-ExchangeAdminSetConditionalAccessPolicy {
    <#
	.SYNOPSIS
    Parses ExchangeAdmin Set-ConditionalAccessPolicy events from UAL.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory )]
        [psobject] $Log,

        [Parameter( Mandatory )]
        [psobject] $AuditData
    )

    begin {

        # variables
        $SummaryStrings = [System.Collections.Generic.List[string]]::new()
    }

    process {

        # DisplayName
        $DisplayName = ( $AuditData.Parameters | Where-Object { $_.Name -eq 'DisplayName' } ).Value
        $SummaryStrings.Add( $DisplayName )

        # join strings, create return object
        $SummaryString = $SummaryStrings -join ', '
        $EventObject = [pscustomobject]@{
            Summary = $SummaryString
        }

        return $EventObject
    }
}