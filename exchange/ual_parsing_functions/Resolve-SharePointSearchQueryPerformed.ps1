function Resolve-SharePointSearchQueryPerformed {
    <#
	.SYNOPSIS
    Parses SearchQueryPerformed events from UAL.
	
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

        # SearchQueryText
        $SearchQueryText = $AuditData.SearchQueryText
        $SummaryStrings.Add( "SearchQueryText: ${SearchQueryText}" )

        # join strings, create return object
        $SummaryString = $SummaryStrings -join ', '
        $EventObject = [pscustomobject]@{
            Summary = $SummaryString
        }

        return $EventObject
    }
}