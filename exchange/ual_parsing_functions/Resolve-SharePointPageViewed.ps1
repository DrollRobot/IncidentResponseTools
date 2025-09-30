function Resolve-SharePointPageViewed {
    <#
	.SYNOPSIS
    Parses PageViewed events from UAL.
	
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

        # ObjectId
        $ObjectId = $AuditData.ObjectId
        $SummaryStrings.Add( "ObjectId: ${ObjectId}" )


        # join strings, create return object
        $SummaryString = $SummaryStrings -join ', '
        $EventObject = [pscustomobject]@{
            Summary = $SummaryString
        }

        return $EventObject
    }
}