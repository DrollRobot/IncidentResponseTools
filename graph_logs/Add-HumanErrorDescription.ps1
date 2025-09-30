function Add-HumanErrorDescription {
	<#
	.SYNOPSIS
	Helper function for Entra sign in logs. Accepts sign in log object, adds "Error" property with text description of error number.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param (
        [psobject[]] $Logs
    )

    begin {

        # variables
        $ModulePath = $PSScriptRoot
        $CsvPath = Join-Path $ModulePath -ChildPath "\entra_error_codes.csv"

        # import csv table
        $ErrorTable = Import-Csv -Path $CsvPath
	}

    process {

        foreach ( $Log in $Logs ) {

            $ErrorCode = $Log.Status.ErrorCode
            $ErrorInfo = $ErrorTable | Where-Object { $_.Error -eq $ErrorCode }
            $CustomDescription = $ErrorInfo.CustomDescription
            $ShortDescription = $ErrorInfo.ShortDescription

            if ( $ErrorCode -eq 0 ) {
                $ErrorString = "0:SUCCESS"
            }
            elseif ( $CustomDescription ) {
                $ErrorString = "${ErrorCode}:${CustomDescription}"
            }
            elseif ( $ShortDescription ) {
                $ErrorString = "${ErrorCode}:${ShortDescription}"
            }
            else {
                $ErrorString = $ErrorCode
            }

            $Params = @{
                MemberType = 'NoteProperty'
                Name = 'Error'
                Value = $ErrorString
            }
            $Log | Add-Member @Params
        }

        return $Logs
    }
}