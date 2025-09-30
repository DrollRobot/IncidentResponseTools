function Resolve-ExchangeItemGroupDelete {
    <#
	.SYNOPSIS
    Parses ExchangeItemGroup HardDelete events from UAL.
	
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

        # Folders
        foreach ( $Folder in $AuditData.Folder ) {

            $FolderPath = $Folder.Path
            $SummaryStrings.Add( "DeletedFrom: ${FolderPath}" )
        }

        # AffectedItems
        foreach ( $AffectedItem in $AuditData.AffectedItems ) {

            $Subject = $AffectedItem.Subject
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