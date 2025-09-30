function Resolve-ExchangeItemAggregatedMailItemsAccessed {
    <#
	.SYNOPSIS
    Parses ExchangeItemAggregated MailItemsAccessed events from UAL.
	
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
        $Summary = [System.Collections.Generic.List[string]]::new()
    }

    process {

        # Folders
        foreach ($Folder in $AuditData.Folders) {

            $Summary.Add( "Folder: $($Folder.Path)" )

            $FolderItems = $Folder.FolderItems
            foreach ($Item in $FolderItems) {
                $Summary.Add( "    Item: $($Item.InternetMessageId)" )
            }
        }

        # join strings, create return object
        $AllSummary = $Summary -join "`n"
        $EventObject = [pscustomobject]@{
            Summary = $AllSummary
        }

        return $EventObject
    }
}