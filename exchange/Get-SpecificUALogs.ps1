New-Alias -Name 'SpecUALogs' -Value 'Get-SpecificUALogs' -Force
New-Alias -Name 'SpecificUALogs' -Value 'Get-SpecificUALogs' -Force
function Get-SpecificUALogs {
    <#
	.SYNOPSIS
    Queries unified audit log for app consent events.
    
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding(DefaultParameterSetName='UserObjects')]
    param (
        [Parameter( Position = 0 )]
        [Alias( 'UserObject' )]
	    [Parameter(ParameterSetName='UserObjects')]
        [psobject[]] $UserObjects,
	    [Parameter(Mandatory,ParameterSetName='AllUsers')]
        [switch] $AllUsers,

        [switch] $AppConsent,
        [switch] $InboxRules,
        [switch] $MFA,
        [switch] $AllOperations,

        [int] $Days = 365,
        [switch] $Passthru,
        [boolean] $Xml = $true,
        [boolean] $Open = $true
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $EndDate = Get-Date

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Cyan = @{ ForegroundColor = 'Cyan' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        if ($PSCmdlet.ParameterSetName -eq 'UserObjects') {
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
     
        # get client domain name for file output
        $DefaultDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
        $DomainName = $DefaultDomain.DomainName -split '\.' | Select-Object -First 1
    }

    process {
        # build query base params
        $QueryParams = @{
            ResultSize     = 5000
            SessionCommand = 'ReturnLargeSet'
            Formatted      = $true
            StartDate      = (Get-Date).AddDays( $Days * -1 ).ToUniversalTime() 
            EndDate        = $EndDate.ToUniversalTime()
        }

        # build list of all userid variations for query
        if ($PSCmdlet.ParameterSetName -eq 'UserObjects') {
            $UserIds = [System.Collections.Generic.List[string]]::new()
            foreach ( $ScriptUserObject in $ScriptUserObjects ) {
                $UserIds.Add($ScriptUserObject.UserPrincipalName)
                $UserIds.Add($ScriptUserObject.Id)
                $UserIds.Add(($ScriptUserObject.Id -replace '-', ''))
            }
            if (($UserIds | Measure-Object).Count -gt 0) {
                $QueryParams['UserIds'] = $UserIds
            }
        }

        # build list of operations for query
        $Operations = [System.Collections.Generic.List[string]]::new()
        $FileNameOperations = [System.Collections.Generic.List[string]]::new()
        if ($AppConsent -or $AllOperations) {
                $FileNameOperations.Add("AppConsent")
                $Operations.Add("Consent to application.")
                $Operations.Add("Add app role assignment grant to user.")
                $Operations.Add("Add service principal.")
                $Operations.Add("Add delegated permission grant.") 
        }
        if ($InboxRules -or $AllOperations) {
            $FileNameOperations.Add("InboxRules")
            $Operations.Add("New-InboxRule") # pulled from https://learn.microsoft.com/en-us/purview/audit-log-activities, may be wrong
            $Operations.Add("Set-InboxRule") # pulled from https://learn.microsoft.com/en-us/purview/audit-log-activities, may be wrong
        }
        if ($MFA -or $AllOperations) {
            $FileNameOperations.Add("MFA")
            $Operations.Add("User registered security info")
            $Operations.Add("Admin registered security info")
            $Operations.Add("User registered all required security info")
        }

        if (($Operations | Measure-Object).Count -gt 0) {
            $QueryParams['Operations'] = $Operations
        }
        else {
            throw "An operation must be selected: -AppConsent, -InboxRules, -MFA"
        }

        # build file name
        if ($PSCmdlet.ParameterSetName -eq 'AllUsers') {
            $UserNameString = $PSCmdlet.ParameterSetName
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'UserObjects') {
            $UserCount = ($ScriptUserObjects | Measure-Object).Count
            $UserNameString = switch ($UserCount) {
                0 {'SCRIPTERROR'}
                1 {$ScriptUserObjects.UserPrincipalName -split '@' | Select-Object -First 1}
                default {"${UserCount}Users"}
            }
        }
        $OperationCount = ($FileNameOperations | Measure-Object).Count
        $OperationsString = switch ($OperationCount) {
            0 {'SCRIPTERROR'}
            1 {$FileNameOperations[0]}
            default {"${OperationCount}Operations"}
        }
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileNameDate = Get-Date -Format $FileNameDateFormat
        $XmlOutputPath = "${OperationsString}_UAL_Raw_${Days}Days_${DomainName}_${UserNameString}_${FileNameDate}.xml"

        # run first query
        Write-Host @Blue "Running Search-UnifiedAuditLog:"
        $UsersString = if ($PSCmdlet.ParameterSetName -eq 'AllUsers') {
            'AllUsers'
        } 
        elseif ($PSCmdlet.ParameterSetName -eq 'UserObjects') {
            ($ScriptUserObjects.UserPrincipalName) -join ', ' 
        }
        Write-Host @Blue "Users: ${UsersString}"
        Write-Host @Blue ("Operations: " + ($Operations -join ', '))

        $AllLogs = [System.Collections.Generic.List[psobject]]::new()

        $Page = Search-UnifiedAuditLog @QueryParams
        $LogCount = @($Page).Count

        if ( $LogCount -gt 0 ) {

            Write-Host @Blue "Retrieved ${LogCount} logs."

            # add to list
            foreach ($i in $Page) {$AllLogs.Add($i)}

            # extract sessionid for paging
            $SessionId = $Page[0].SessionId
            $PageCount = 2
            $NextPageParams = $QueryParams
            $NextPageParams['SessionId'] = $SessionId
        }
        else {
            Write-Host @Red "Retrieved 0 logs."
        }

        # retrieve additional pages, if first page was 5000
        while ( $LogCount -eq 5000 ) {

            Write-Host @Blue "Requesting page ${PageCount}."
            $Page = Search-UnifiedAuditLog @NextPageParams
            $LogCount = @($Page).Count

            if ( $LogCount -gt 0 ) {

                Write-Host @Blue "Retrieved ${LogCount} logs."

                # add to list
                foreach ($i in $Page) {$AllLogs.Add($i)}

                # extract sessionid for paging
                $SessionId = $Page[0].SessionId
            }
            else {
                Write-Host @Red "Retrieved 0 logs."
            }

            $PageCount++
        }


        # remove duplicates
        $UniqueLogIds = [System.Collections.Generic.HashSet[string]]::new()
        $UniqueLogs = [System.Collections.Generic.List[psobject]]::new()
        foreach ($Log in $AllLogs) {
            if ($UniqueLogIds.Add([string]$Log.Identity)) { 
                $UniqueLogs.Add($Log) | Out-Null
            }
        }

        ### sort list
        # build comparison script
        $PropertyName = 'CreationDate'
        $Descending = $true
        $Comparison = [System.Comparison[PSObject]] {
            param($X, $Y)
            $Result = $X.$PropertyName.CompareTo($Y.$PropertyName)
            if ($Descending) {
                return -1 * $Result
            }
            return $Result
        }
        $UniqueLogs.Sort($Comparison)

        # export raw data to xml
        if ($Xml) {
            Write-Host @Blue "`nSaving logs to: ${XmlOutputPath}"
            $UniqueLogs | Export-Clixml -Depth 10 -Path $XmlOutputPath
        }


        $Showparams = @{
            Logs = $UniqueLogs
            DomainName = $DomainName
            UserName = $UserNameString
            Days = $Days
            EndDate = $EndDate
            OperationString = $OperationsString
            Open = $Open
        }
        Show-UALogs @ShowParams


        if ($Passthru) {
            Write-Output $UniqueLogs
        }
    }
}