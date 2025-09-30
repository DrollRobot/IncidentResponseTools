New-Alias -Name 'Get-GraphGlobalUserObject' -Value 'Get-GraphGlobalUserObjects' -Force
function Get-GraphGlobalUserObjects {
	<#
	.SYNOPSIS
	Gets user objects from global variables. Designed to be used by other scripts.
	
	.NOTES
	Version: 1.0.3
	#>
    [CmdletBinding()]
    param (
    )

    begin {

        # variables
		$ScriptUserObjects = [System.Collections.Generic.List[PsObject]]::new()
	}

    process {

		# add userobject
		if ( $Global:UserObject ) {
			$ScriptUserObjects.Add( $Global:UserObject )
		}

		# add userobjects
		if ( $Global:UserObjects ) {
            $IterationList = @( $Global:UserObjects )  
			foreach ( $i in $IterationList ) {
				$ScriptUserObjects.Add( $i )
			}
		}

		# return user objects
		return $ScriptUserObjects | Sort-Object Id -Unique | Sort-Object DisplayName
    }
}