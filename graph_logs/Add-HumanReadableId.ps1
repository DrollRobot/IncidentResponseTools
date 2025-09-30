function Add-HumanReadableId {
    <#
	.SYNOPSIS
	Translates GUIDs into human readable session numbers.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
        [psobject[]] $Logs,

        [string] $Property,

        [string] $ColumnHeader
    )

    begin {

        # variables
        $SessionTable = @{}
    }

    process {

        # create table of session ids
        $UniqueIds = $Logs.$Property | Sort-Object -Unique
        for ($i = 0; $i -lt @( $UniqueIds ).Count; $i++) {

            # get specific id
            $Id = $UniqueIds[$i]

            # add sub table to main table
            $SessionTable[$Id] = @{}

            # add human readable id to table
            $SessionTable[$Id].HumanId = $i + 1

            # add count to table
            $Count = @( $Logs.$Property | Where-Object { $_ -eq $Id } ).Count
            $SessionTable[$Id].Count = $Count
        }

        # add human readable session id to logs variable
        foreach ( $Log in $Logs ) {

            $HumanId = $SessionTable[$Log.$Property].HumanId
            $Count = $SessionTable[$Log.$Property].Count

            $Params = @{
                MemberType = 'NoteProperty'
                Name       = $ColumnHeader
                Value      = "${HumanId} (${Count})"
            }
            $Log | Add-Member @Params
        }

        return $Logs
    }
}