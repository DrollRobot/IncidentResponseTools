function Resolve-ExchangeItemSubject {
    <#
	.SYNOPSIS
    Parses ExchangeItem events from UAL.
	
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

        # Items
        foreach ( $Item in $AuditData.Item ) {

            $Subject = $Item.Subject
            $SummaryStrings.Add( "Subject: ${Subject}" )
        }

        # join strings, create return object
        $SummaryString = $SummaryStrings -join ', '
        $EventObject = [pscustomobject]@{
            Summary = $SummaryString
        }

        return $EventObject
    }
}