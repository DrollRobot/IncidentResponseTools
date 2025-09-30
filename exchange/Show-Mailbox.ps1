New-Alias -Name 'ShowMailbox' -Value 'Show-Mailbox' -Force
function Show-Mailbox {
    <#
	.SYNOPSIS
	Displays mailbox properties.	
	
	.NOTES
	Version: 1.1.0
	#>
    [CmdletBinding()]
    param(
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects
    )

    begin {

        # if not passed directly, find global
        if ( -not $UserObjects -or $UserObjects.Count -eq 0 ) {

            # get from global variables
            $ScriptUserObjects = Get-GraphGlobalUserObjects

            # if none found, exit
            if ( -not $ScriptUserObjects -or $ScriptUserObjects.Count -eq 0 ) {
                throw "No user objects passed or found in global variables."
            }
        }
        else {
            $ScriptUserObjects = $UserObjects
        }

        # variables
        # $PermissionsList = [System.Collections.Generic.List[pscustomobject]]::new()
        $GuidPattern = '\b[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\b'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
    }

    process {

        # # get mailbox permissions
        # # get all mailboxes 
        # $AllMailboxes = Get-EXOMailbox -ResultSize Unlimited
        # if ( $AllMailboxes.Count -lt 100 ) {
        #     foreach ( $Mailbox in $AllMailboxes ) {

        #         $Permissions = Get-EXOMailboxPermission -Identity $Mailbox.UserPrincipalName

        #         $AddParams = @{
        #             MemberType  = 'NoteProperty'
        #             Name        = 'Permissions'
        #             Value       = $Permissions
        #         }
        #         $Mailbox | Add-Member @AddParams
        #     }
        # }

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            # get user mailbox info
            $UserPrincipalName = $ScriptUserObject.UserPrincipalName
            try {
                $Mailbox = Get-EXOMailbox -UserPrincipalName $UserPrincipalName -PropertySets All
            }
            catch {}
            if ( -not $Mailbox ) {
                Write-Host @Red "`nNo mailbox for: ${UserPrincipalName}"
                continue
            }
            $Permissions = Get-EXOMailboxPermission -Identity $UserPrincipalName

            # # find other mailboxes user has permissions for
            # foreach ( $Mailbox in $AllMailboxes ) {
            #     $Mailbox.Permissions | 
            #         Where-Object { $_.User -eq $UserPrincipalName } |
            #         ForEach-Object { 
            #             $PermissionsList.Add(
            #                 [pscustomobject]@{
            #                     Mailbox = $UserPrincipalName
            #                     User = $_.User
            #                     AccessRights = $_.AccessRights
            #                 }
            #             )
            #         }
            # }
            # Write-Host @Blue "`nShowing mailboxes ${UserEmail} has access to: "
            # if ( $PermissionsList ) {
            #     $PermissionsList | Format-Table -AutoSize
            # }
            # else {
            #     Write-Host "None"
            # }

            # if forwarding address is GUID, look up user
            if ( $Mailbox.ForwardingAddress -match $GuidPattern ) {

                $UserGuid = $Mailbox.ForwardingAddress

                # get user object
                $Users = Request-GraphUsers
                $MatchingUser = $Users | Where-Object { $_.Id -eq $UserGuid }

                $ForwardingAddress = $MatchingUser.Mail
            }
            else {
                $ForwardingAddress = $Mailbox.ForwardingAddress
            }

            # convert dates to local
            try {
                $WhenCreatedLocal = $Mailbox.WhenCreatedUTC.ToLocalTime()
                $WhenChangedLocal = $Mailbox.WhenChangedUTC.ToLocalTime()
            }
            catch {}

            Write-Host @Blue "`nShowing Mailbox information for: ${UserPrincipalName}"
            $OutputTable = [PSCustomObject]@{
                IsMailboxEnabled      = $Mailbox.IsMailboxEnabled
                AuditEnabled          = $Mailbox.AuditEnabled
                AuditLogAgeLimit      = $Mailbox.AuditLogAgeLimit
                DisplayName           = $Mailbox.DisplayName
                PrimarySmtpAddress    = $Mailbox.PrimarySmtpAddress
                EmailAddresses        = $Mailbox.EmailAddresses
                WhenCreated           = $WhenCreatedLocal
                WhenChanged           = $WhenChangedLocal
                ForwardingAddress     = $ForwardingAddress
                ForwardingSmtpAddress = $Mailbox.ForwardingSmtpAddress
                DeliverToMailboxAndForward = $Mailbox.DeliverToMailboxAndForward
                LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                RetentionHoldEnabled = $Mailbox.RetentionHoldEnabled
                UsageLocation = $Mailbox.UsageLocation
            }
            $OutputTable | Format-List | Out-Host

            Write-Host @Blue "`nShowing users who have delegated access to: ${UserPrincipalName}"
            $DisplayProperties = @(
                "User"
                "AccessRights"
                "IsInherited"
                "Deny"
                "InheritanceType"
            )
            $Permissions | Format-Table $DisplayProperties -AutoSize | Out-Host
        }
    }
}
