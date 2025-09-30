function Get-UnknownObject {
    <#
	.SYNOPSIS
	Uses Get-MgDirectoryObject to find object type, then uses dedicated command for that type to return object.	
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding()]
    param(
        [string] $Id
    )

    begin {

        # variables
        $DirectoryObject = Get-MgDirectoryObject -DirectoryObjectId $Id
        $ObjectType = $DirectoryObject.AdditionalProperties.'@odata.type' -replace '#microsoft\.graph\.', ''

        $UserGetProperties = @(
            'AccountEnabled'
            'DisplayName'
            'Id'
            'UserPrincipalName'
        )

        $ServicePrincipalGetProperties = @(
            'AccountEnabled'
            'Description'
            'DisplayName'
            'Id'
            'ServicePrincipalType'
        )

        $GroupGetProperties = @(
            'Description'
            'DisplayName'
            'Id'
        )
    }

    process {

        switch ( $ObjectType ) {
            'group' {

                $Object = Get-MgGroup -GroupId $Id -Property $GroupGetProperties
                return $Object
            }
            'servicePrincipal' {

                $Object = Get-MgServicePrincipal -ServicePrincipalId $Id -Property $ServicePrincipalGetProperties
                return $Object
            }
            'user' {

                $Object = Get-MgUser -UserId $Id -Property $UserGetProperties
                return $Object
            }
            default {

                Write-Error "Unknown object type: ${ObjectType}"
            }
        }
    }
}