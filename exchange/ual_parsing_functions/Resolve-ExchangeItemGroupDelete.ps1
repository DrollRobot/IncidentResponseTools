function Resolve-ExchangeItemGroupDelete {
    <#
	.SYNOPSIS
    Parses ExchangeItemGroup HardDelete events from UAL.
	
	.NOTES
	Version: 1.1.0
    1.1.0 - Removed Auditdata param, added parsing for email subjects.
	#>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory )]
        [psobject] $Log,

        [Parameter( Mandatory )]
        [boolean] $WaitOnMessageTrace,

        [Parameter( Mandatory )]
        [string] $UserName
    )

    begin {
        $Function = $MyInvocation.MyCommand.Name
        $VariableName = "IRT_MessageTraceTable_${UserName}"

        # colors
        # $Blue = @{ ForegroundColor = 'Blue' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        $Yellow = @{ ForegroundColor = 'Yellow' }

        $Summary = [System.Collections.Generic.List[string]]::new()

        # check for message trace table

        if ($WaitOnMessageTrace) {
            while (-not (Test-Path "variable:global:${VariableName}")) {
                Write-Host @Yellow "${Function}: Waiting on `$Global:${VariableName}..."
                Start-Sleep -Seconds 15
            }
        }

        if (Test-Path "variable:global:${VariableName}") {
            $Params = @{
                Name = $VariableName
                Scope = 'Global'
            }
            $MessageTraceTable = Get-Variable @Params
        }
    }

    process {

        # AffectedItems

        # build table by folder
        $FolderTable = @{}

        foreach ( $AffectedItem in $Log.AuditData.AffectedItems ) {

            $FolderPath = $AffectedItem.ParentFolder.Path

            # if table key doesn't exist, create it.
            if (-not $FolderTable.ContainsKey($FolderPath)) {
                $FolderTable[$FolderPath] = [System.Collections.Generic.List[psobject]]::new()
            }

            # add object to table
            $FolderTable[$FolderPath].Add($AffectedItem)
        }

        # loop through folders
        foreach ($Folder in $FolderTable.GetEnumerator()) {

            $Summary.Add( "Folder: $($Folder.Name)" )

            # loop through items
            foreach ($Item in $Folder.Value) {

                # if item has subject property, use it
                if ($Item.Subject) {
                    $Subject = $Item.Subject
                }
                elseif ($Item.InternetMessageId) {
                    # if not, try to retrieve from message trace table.
                    $InternetMessageId = $Item.InternetMessageId
                    if ($MessageTraceTable.Value) {
                        if ($MessageTraceTable.Value.ContainsKey($InternetMessageId)) {
                            $Subject = $MessageTraceTable.Value[$InternetMessageId].Subject
                        }
                    }
                }

                # add best option to summary
                if ($Subject) {
                    $Summary.Add( "    Subject: ${Subject}" )            
                }
                elseif ($InternetMessageId) {
                    $Summary.Add( "    Item: ${InternetMessageId}" )
                }
                else {
                    $Summary.Add( "    Item: $($Item.Id)" )
                }
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