function Resolve-ExchangeItemUpdate {
    <#
	.SYNOPSIS
    Parses ExchangeItem Update events from UAL.
	
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

        # ModifiedProperties
        foreach ( $Item in $AuditData.ModifiedProperties ) {
            $SummaryStrings.Add( "Modified: ${Item}" )
        }

        # Items
        foreach ( $Item in $AuditData.Item ) {
            $Subject = $Item.Subject
            $SummaryStrings.Add( "Item: ${Subject}" )
        }

        # join strings, create return object
        $SummaryString = $SummaryStrings -join ', '
        $EventObject = [pscustomobject]@{
            Summary = $SummaryString
        }

        return $EventObject
    }
}