New-Alias -Name 'GetAdmins' -Value 'Get-AdminRoles' -Force
function Get-AdminRoles {
    [CmdletBinding( DefaultParameterSetName = 'All' )]
    param(
        [Parameter( ParameterSetName = 'All' )]
        [switch] $All,

        [Parameter( ParameterSetName = 'Group' )]
        [string] $GroupId,

        [string] $ObjectId,
        [switch] $Script,
        [switch] $Csv,
        [string] $TenantId
    )

    begin {

        if ( -not $GroupId -and -not $ObjectId ) {

            $All = $true
        }
 
        # variables
        if ( -not $Script:CustomObjects ) {
            $Script:CustomObjects = [system.collections.generic.list[pscustomobject]]::new()
        }
        if ( -not $Script:GroupsChecked ) {
            $Script:GroupsChecked = [system.collections.generic.list[string]]::new()
        }
        $ExportDateFormat = "yy-MM-dd"

        $UserDisplayProperties = @(
            'AccountEnabled'
            'DisplayName'
            'UserPrincipalName'
            'RoleSource'
            'Roles'
        )
        $null = $UserDisplayProperties

        $ServicePrincipalDisplayProperties = @(
            'AccountEnabled'
            'ServicePrincipalType'
            'DisplayName'
            'RoleSource'
            'Roles'
        )
        $null = $ServicePrincipalDisplayProperties

        $GroupDisplayProperties = @(
            'DisplayName'
            'RoleSource'
            'Roles'
        )
        $null = $GroupDisplayProperties

        $CsvSortOrder = @(
            'ObjectType'
            'AccountEnabled'
            'ServicePrincipalType'
            'Id'
            'DisplayName'
            'UserPrincipalName'
            'Description'
            'RoleSource'
            'Roles'
        )

        $AllProperties = @()

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        # $Red = @{ ForegroundColor = 'Red' }
    }

    process {

        if ( $All ) {

            # get all direct role assignments
            $RoleObjects = Get-MgDirectoryRole -ExpandProperty Members
            $MemberIds = $RoleObjects.Members.Id | Sort-Object -Unique
        }
        elseif ( $GroupId ) {

            # get all group members
            $MemberIds = ( Get-MgGroupMember -GroupId $GroupId ).Id
        }

        foreach ( $MemberId in $MemberIds ) {

            if ( $All ) {

                # get member roles
                $MemberRoles = ( $RoleObjects | Where-Object { $MemberId -in $_.Members.Id } ).DisplayName
                $RolesString = $MemberRoles -join ', '

                # create custom object
                $CustomObject = [pscustomobject]@{
                    Id         = $MemberId
                    Roles      = $RolesString
                    RoleSource = 'Direct Assignment'
                }
            }
            elseif ( $GroupId ) {

                # get custom object for group
                $GroupCustomObject = $Script:CustomObjects | Where-Object { $_.Id -eq $GroupId }

                # create custom object
                $CustomObject = [pscustomobject]@{
                    Id         = $MemberId
                    Roles      = $GroupCustomObject.Roles
                    RoleSource = "Group: $($GroupCustomObject.DisplayName)"
                }
            }

            # find object type
            $Object = Get-UnknownObject -Id $MemberId
            $ObjectType = $Object.GetType().Name

            switch ( $ObjectType ) {
                'MicrosoftGraphUser' {
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'ObjectType' -Value 'User'
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'AccountEnabled' -Value $Object.AccountEnabled
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'DisplayName' -Value $Object.DisplayName
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'UserPrincipalName' -Value $Object.UserPrincipalName
                }
                'MicrosoftGraphServicePrincipal' {
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'ObjectType' -Value 'ServicePrincipal'
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'AccountEnabled' -Value $Object.AccountEnabled
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'ServicePrincipalType' -Value $Object.ServicePrincipalType
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'DisplayName' -Value $Object.DisplayName
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'Description' -Value $Object.Description
                }
                'MicrosoftGraphGroup' {
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'ObjectType' -Value 'Group'
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'DisplayName' -Value $Object.DisplayName
                    $CustomObject | Add-Member -MemberType 'NoteProperty' -Name 'Description' -Value $Object.Description
                }
                default {
                    Write-Error "Unknown object type: ${ObjectType}"
                }
            }

            $Script:CustomObjects.Add( $CustomObject )

            # if group, add members of group
            if ( $ObjectType -eq 'MicrosoftGraphGroup' ) {
                if ( $MemberId -notin $Script:GroupsChecked ) {
                    $Script:GroupsChecked.Add( $MemberId )
                    Get-AdminRoles -GroupId $MemberId
                }
            }
        }
    }

    end {

        if ( $All ) {

            # sort objects
            $Script:CustomObjects = $Script:CustomObjects | Sort-Object @{Expression = "ObjectType"; Descending = $true }, @{Expression = "AccountEnabled"; Descending = $true }

            # display table
            if ( $Script ) {

                return $Script:CustomObjects
            }
            elseif ( $Csv ) {

                # show in terminal
                Show-CustomObjects -CustomObjects $Script:CustomObjects

                # build file name
                $DefaultDomain = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
                $DomainName = $DefaultDomain.Id -split '\.' | Select-Object -First 1
                $Date = Get-Date -Format $ExportDateFormat
                $FileName = "AdminRoles_${DomainName}_${Date}.csv"

                ### build sorted list of properties to export
                # get all properties
                foreach ( $Object in $Script:CustomObjects ) {
                    $AllProperties += ( $Object | Get-Member -MemberType 'NoteProperty' ).Name
                }
                $AllProperties = $AllProperties | Sort-Object -Unique
                # sort based on custom sort order
                $SortedProperties = $AllProperties | Sort-Object -Property @{
                    Expression = {
                        $Index = $CsvSortOrder.IndexOf( $_ )
                        # if not in the list, make last
                        if ( $Index -eq -1 ) {
                            [int]::MaxValue
                        }
                        else {
                            $Index
                        }
                    }
                    Ascending  = $true
                }

                # export
                Write-Host @Blue "Exporting CSV: ${FileName}"
                $Script:CustomObjects | Select-Object -Property $SortedProperties | Export-Csv -Path $FileName -Force
            }
            else {

                Show-CustomObjects -CustomObjects $Script:CustomObjects
            }

            $Script:CustomObjects = $null
        }
    }
}


