function Get-MailboxesUserHasFullAccessTo {
    param(
        [string] $UserPrincipalName
    )
    $Blue = @{ ForegroundColor = 'Blue' }
    if ( -not $UserPrincipalName ) {
        $UserPrincipalName = (Get-ConnectionInformation).UserPrincipalName | Sort-Object -Unique
    }
    if ( -not $UserPrincipalName ) {
        Write-Error "Specify -UserPrincipalName."
        return
    }
    # retrieve all mailboxes
    $Mailboxes = Get-Mailbox -ResultSize Unlimited
    $Total = $Mailboxes.Count
    $Index = 0
    # create output containers
    $FullAccessList = [System.Collections.Generic.List[psobject]]::new()
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($Mailbox in $Mailboxes) {
        $Index++
        if ($Index -gt 0) {
            $Elapsed = $Stopwatch.Elapsed
            $EstimatedTotal = [timespan]::FromSeconds($Elapsed.TotalSeconds / $Index * $Total)
            $Remaining = $EstimatedTotal - $Elapsed
            $TimeRemainingText = "{0:hh\:mm\:ss}" -f $Remaining
        } else {
            $TimeRemainingText = "Estimating..."
        }
        # display progress
        $ProgressParams = @{
            Activity         = "Checking mailbox permissions"
            Status           = "Processing $($Mailbox.Name) [$Index of $Total] — Est. remaining: $TimeRemainingText"
            PercentComplete  = ($Index / $Total * 100)
        }
        Write-Progress @ProgressParams
        # check FullAccess permissions
        try {
            $Permissions = Get-MailboxPermission -Identity $Mailbox.Identity -ErrorAction Stop
            foreach ($Permission in $Permissions) {
                if ($Permission.User.ToString() -eq $User -and $Permission.AccessRights -contains 'FullAccess') {
                    $FullAccessList.Add($Mailbox)
                    break
                }
            }
        } catch {
            # silently continue on permission errors
        }
    }
    # output results
    Write-Host @Blue "`nMailboxes with FullAccess rights granted to $User"
    if ( -not $FullAccessList ) { Write-Host "None" }
    return $FullAccessList | Select-Object Name, PrimarySmtpAddress
}
