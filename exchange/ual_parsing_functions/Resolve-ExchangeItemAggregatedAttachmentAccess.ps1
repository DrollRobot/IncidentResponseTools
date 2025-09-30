function Resolve-ExchangeItemAggregatedAttachmentAccess {
    <#
	.SYNOPSIS
    Parses ExchangeItemAggregated AttachmentAccess events from UAL.
	
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
        # $SummaryStrings = [System.Collections.Generic.List[string]]::new()
    }

    process {

        # need to lookup email by ID.

        return $EventObject
    }
}