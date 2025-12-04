New-Alias -Name 'Playbook' -Value 'Start-IncidentResponsePlaybook'

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

        # show module version
        $CallStack = @( Get-PSCallStack )
        if ( $CallStack.Count -eq 2 ) {
            $ModuleVersion = $ExecutionContext.SessionState.Module.Version
            Write-Host "Module version: ${ModuleVersion}"
        }

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
                    Message     = "No -UserObjects argument used, no `$Global:UserObjects present."
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
            @{  Script = {
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
            @{  Script = {
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
            @{  Script = {
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
            @{  Script = { 
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
            @{  Script = { 
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
            @{  Script = { 
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
            @{  Script = { 
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
            @{  Script = { 
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
            @{  Script = { 
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
            @{  Script = {
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

            $Global:Playbook_JobList = @()
            $Global:Playbook_RunspacePool = $null

            ### build a runspace pool
            $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $InitialSessionState.ImportPSModule(
                # 'ExchangeOnlineManagement',
                'IncidentResponseTools',
                'Microsoft.Graph.Authentication',
                'Microsoft.Graph.Applications',
                'Microsoft.Graph.Beta.Reports',
                'Microsoft.Graph.Users'
            )
            $Global:Playbook_RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads, $InitialSessionState, $Host)
            $Global:Playbook_RunspacePool.Open()

            ### queue tasks
            $Global:Playbook_JobList = foreach ($Step in $Steps) {

                $PowerShell = [PowerShell]::Create()
                $PowerShell.RunspacePool = $Global:Playbook_RunspacePool

                $null = $PowerShell.AddScript($Step.Script)
                foreach ($Arg in $Step.Args) {
                    $null = $PowerShell.AddArgument($Arg) 
                }

                # loop output
                [pscustomobject]@{
                    PowerShell = $PowerShell
                    Handle     = $PowerShell.BeginInvoke()
                    Completed  = $false
                }
            }

            ### exchange functions
            # show mailbox properties
            Show-Mailbox -UserObjects $ScriptUserObjects

            # download email rules
            Get-IRTInboxRules -UserObjects $ScriptUserObjects

            # download maximum available user message trace
            Get-IRTMessageTrace -Days 90 -UserObjects $ScriptUserObjects

            # download user UAL
            Get-UALogs -Days 1 -WaitOnMessageTrace:$true -UserObjects $ScriptUserObjects

            # # download high risk UAL
            # Get-UALogs -AllUsers -RiskyOperations -Days 180

            # # # download 2 day message trace for all users
            Get-IRTMessageTrace -AllUsers -Days 2

            ### wait for completion, collect errors
            while ($Global:Playbook_JobList.Completed -contains $false) {
                foreach ($Job in $Global:Playbook_JobList) {
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
                            Write-Error $_
                        }
                        finally {
                            $Job.PowerShell.Dispose()
                            $Job.Completed = $true
                        }
                    }
                }

                Start-Sleep -Seconds 10
                if ( $Script:Test -and $Stopwatch.Elapsed.Minutes -ge 5 ) {
                    Write-Host @Yellow "Waiting on "
                }
            }
        }
        finally {

            ### cleanup
            # stop all runspaces
            foreach ($Job in $Global:Playbook_JobList) {
                try   { $Job.PowerShell.Stop() } catch {}
                try   { $Job.PowerShell.Dispose() } catch {}
            }
            $Global:Playbook_JobList = @()

            # close pool
            if ($Global:Playbook_RunspacePool) {
                try { $Global:Playbook_RunspacePool.Close() }  catch {}
                try { $Global:Playbook_RunspacePool.Dispose() } catch {}
            }
            $Global:Playbook_RunspacePool = $null
        }

        ### cleanup
        $Stopwatch.Stop()
        Write-Host "Playbook complete. Elapsed time: $($Stopwatch.Elapsed.ToString())"
    }
}