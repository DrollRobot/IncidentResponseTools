New-Alias -Name 'Playbook' -Value 'Start-IncidentResponsePlaybook' -Force

function Start-IncidentResponsePlaybook {
    <#
	.SYNOPSIS
    Runs multiple functions to assist in investigating a user's activity.
    
	.NOTES
	Version: 2.2.0
    2.2.0 - Added license report, added error handling to close runspaces when script exits.
    2.0.0 - Added ability to run mulitple operations in parallel using runspaces.
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,
        [string] $Ticket,
        [switch] $NoFolder,
        [int] $Threads = 15,
        [switch] $Test
    )

    begin {

        if ($Test -or $Script:Test) {
            $Script:Test = $true
            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }

        # if users passed via script argument:
        if (($UserObjects | Measure-Object).Count -gt 0) {
            $ScriptUserObjects = $UserObjects
        }
        # if not, look for global objects
        else {
            
            # get from global variables
            $ScriptUserObjects = Get-GraphGlobalUserObjects
            
            # if none found, exit
            if ( -not $ScriptUserObjects -or $ScriptUserObjects.Count -eq 0 ) {
                Write-Host @Red "${Function}: No user objects passed or found in global variables."
                return
            }
            if (($ScriptUserObjects | Measure-Object).Count -eq 0) {
                $ErrorParams = @{
                    Category    = 'InvalidArgument'
                    Message     = "No -UserObjects argument used, no `$Global:IRT_UserObjects present."
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
            }
        }

        # verify connected to graph
        if (-not (Get-MgContext)) {
            $ErrorParams = @{
                Category    = 'ConnectionError'
                Message     = "Not connected to Graph. Run Connect-MgGraph."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams
        }

        # verify connected to exchange
        try {
            [void](Get-AcceptedDomain)
        }
        catch {
            $ErrorParams = @{
                Category    = 'ConnectionError'
                Message     = "Not connected to Exchange. Run Connect-ExchangeOnline."
                ErrorAction = 'Stop'
            }
            Write-Error @ErrorParams
        } 
    }

    process {

        if ( -not $NoFolder ) {
            
            # make directory
            $DirParams = @{
                UserObjects = $ScriptUserObjects
            }
            if ( $Ticket ) {
                $DirParams['Ticket'] = $Ticket
            }
            else {
                $DirParams['Confirm'] = $true
            }
            New-InvestigationDirectory @DirParams
        }

        $WorkingPath = Get-Location

        $Steps = @(
            # Get-LicenseReport
            @{  Name   = 'Get-LicenseReport'
                Script = {
                    param( 
                        $WorkingPath
                    )
                    Set-Location -Path $WorkingPath
                    Get-LicenseReport
                }
                Args  = @(
                    $WorkingPath
                )
            }
            # Show-UserInfo
            @{  Name   = 'Show-UserInfo'
                Script = {
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Show-UserInfo -UserObjects $RunspaceUserObjects 
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
            # Get-UserApplications
            @{  Name   = 'Get-UserApplications'
                Script = {
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Get-UserApplications -UserObjects $RunspaceUserObjects 
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
            # Show-GraphGeoBlockPolicy
            @{  Name   = 'Show-GraphGeoBlockPolicy'
                Script = { 
                    param( 
                        $WorkingPath
                    )
                    Set-Location -Path $WorkingPath
                    Show-GraphGeoBlockPolicy
                }
                Args  = @(
                    $WorkingPath
                )
            }
            # Get-AdminRoles
            @{  Name   = 'Get-AdminRoles'
                Script = { 
                    param( 
                        $WorkingPath
                    )
                    Set-Location -Path $WorkingPath
                    Get-AdminRoles
                }
                Args  = @(
                    $WorkingPath
                )
            }
            # Find-RogueApps
            @{  Name   = 'Find-RogueApps'
                Script = { 
                    param( 
                        $WorkingPath
                    )
                    Set-Location -Path $WorkingPath
                    Find-RogueApps
                }
                Args  = @(
                    $WorkingPath
                )
            }
            # Show-UserMFA
            @{  Name   = 'Show-UserMFA'
                Script = { 
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Show-UserMFA -UserObjects $RunspaceUserObjects
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
            # Get-EntraAuditLogs
            @{  Name   = 'Get-EntraAuditLogs'
                Script = { 
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Get-EntraAuditLogs -UserObjects $RunspaceUserObjects
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
            # Get-SignInLogs
            @{  Name   = 'Get-SignInLogs'
                Script = { 
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Get-SignInLogs -UserObjects $RunspaceUserObjects
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
            # Get-NonInteractiveLogs
            @{  Name   = 'Get-NonInteractiveLogs'
                Script = {
                    param( 
                        $WorkingPath,
                        $RunspaceUserObjects
                    )
                    Set-Location -Path $WorkingPath
                    Get-NonInteractiveLogs -UserObjects $RunspaceUserObjects
                }
                Args  = @(
                    $WorkingPath,
                    $ScriptUserObjects
                )
            }
        )

        try {

            $Global:IRT_Playbook_JobList = @()
            $Global:IRT_Playbook_RunspacePool = $null

            ### build a runspace pool
            $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $InitialSessionState.ImportPSModule(
                # 'ExchangeOnlineManagement', # not needed unless running exhchange commands in runspaces.
                'IncidentResponseTools',
                'Microsoft.Graph.Authentication',
                'Microsoft.Graph.Applications',
                'Microsoft.Graph.Beta.Reports',
                'Microsoft.Graph.Users'
            )
            $Global:IRT_Playbook_RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads, $InitialSessionState, $Host)
            $Global:IRT_Playbook_RunspacePool.Open()

            ### queue tasks
            $Global:IRT_Playbook_JobList = foreach ($Step in $Steps) {

                $PowerShell = [PowerShell]::Create()
                $PowerShell.RunspacePool = $Global:IRT_Playbook_RunspacePool

                $null = $PowerShell.AddScript($Step.Script)
                foreach ($Arg in $Step.Args) {
                    $null = $PowerShell.AddArgument($Arg) 
                }

                # loop output
                [pscustomobject]@{
                    Name       = $Step.Name
                    PowerShell = $PowerShell
                    Handle     = $PowerShell.BeginInvoke()
                    Completed  = $false
                }
            }
            ### exchange functions
            # show mailbox properties
            try {
                Show-Mailbox -UserObjects $ScriptUserObjects
            } 
            catch { Write-Warning "Show-Mailbox error: $_" }

            # download email rules
            try {
                Get-IRTInboxRules -UserObjects $ScriptUserObjects
            }
            catch { Write-Warning "Get-IRTInboxRules error: $_" }

            # download maximum available user message trace
            try {
                Get-IRTMessageTrace -Days 90 -UserObjects $ScriptUserObjects
            }
            catch { Write-Warning "Get-IRTMessageTrace error: $_" }

            # download user UAL
            try {
                Get-UALogs -UserObjects $ScriptUserObjects -WaitOnMessageTrace:$true -Days 1
            }
            catch { Write-Warning "Get-UALogs error: $_" }

            # download high risk UAL
            try {
                Get-UALogs -UserObjects $ScriptUserObjects -WaitOnMessageTrace:$true -RiskyOperations
            }
            catch { Write-Warning "Get-UALogs (risky) error: $_" }

            # download 180 days of sign ins
            try {
                Get-UALogs -UserObjects $ScriptUserObjects -WaitOnMessageTrace:$true -SignInLogs
            }
            catch { Write-Warning "Get-UALogs (sign ins) error: $_" }

            # # # download 2 day message trace for all users
            try {
                Get-IRTMessageTrace -AllUsers -Days 2
            }
            catch { Write-Warning "Get-IRTMessageTrace (all users) error: $_" }

            ### wait for completion, collect errors
            while ($Global:IRT_Playbook_JobList.Completed -contains $false) {
                foreach ($Job in $Global:IRT_Playbook_JobList) {
                    if ( -not $Job.Completed -and $Job.Handle.IsCompleted ) {
                        try {
                            $Job.PowerShell.EndInvoke( $Job.Handle )

                            # output errors, if any
                            if ( $Job.PowerShell.Streams.Error.Count -gt 0 ) {
                                Write-Warning "Errors:"
                                $Job.PowerShell.Streams.Error | ForEach-Object {
                                    Write-Warning $_.ToString()
                                }
                            }
                        }
                        catch {
                            Write-Warning "$($Job.Name) error: $_"
                        }
                        finally {
                            $Job.PowerShell.Dispose()
                            $Job.Completed = $true
                        }
                    }
                }

                $TotalJobs     = $Global:IRT_Playbook_JobList.Count
                $CompletedCount = ($Global:IRT_Playbook_JobList | Where-Object { $_.Completed }).Count
                $RemainingNames = $Global:IRT_Playbook_JobList | Where-Object { -not $_.Completed } | Select-Object -ExpandProperty Name
                $PercentComplete = [int](($CompletedCount / $TotalJobs) * 100)
                Write-Progress -Activity 'Playbook Running' -Status "Waiting on: $($RemainingNames -join ', ')" -PercentComplete $PercentComplete
                Start-Sleep -Seconds 10
            }
            Write-Progress -Activity 'Playbook Running' -Completed
        }
        finally {

            ### cleanup
            # stop all runspaces
            foreach ($Job in $Global:IRT_Playbook_JobList) {
                try   { $Job.PowerShell.Stop() } catch {}
                try   { $Job.PowerShell.Dispose() } catch {}
            }
            $Global:IRT_Playbook_JobList = @()

            # close pool
            if ($Global:IRT_Playbook_RunspacePool) {
                try { $Global:IRT_Playbook_RunspacePool.Close() }  catch {}
                try { $Global:IRT_Playbook_RunspacePool.Dispose() } catch {}
            }
            $Global:IRT_Playbook_RunspacePool = $null
        }

        ### cleanup
        if ($Stopwatch) {
            $Stopwatch.Stop()
            Write-Host "Playbook complete. Elapsed time: $($Stopwatch.Elapsed.ToString())"
        }
    }
}