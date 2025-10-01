New-Alias -Name 'MessageTrace' -Value 'Get-IRTMessageTrace' -Force
function Get-IRTMessageTrace {
    <#
	.SYNOPSIS
	Downloads incoming and outgoing message trace for specified user, or all users.
	
	.NOTES
	Version: 1.4.0
    1.4.0 - Switched to separate get/show functions. Updated to passing objects, not files. Added global variables.
	#>
    [CmdletBinding( DefaultParameterSetName = 'UserObjects' )]
    param (
        [Parameter( ParameterSetName = 'UserObjects', Position = 0 )]
        [Alias( 'UserObject' )]
        [psobject[]] $UserObjects,

        [Parameter( ParameterSetName = 'UserEmails' )]
        [Alias( 'UserEmail' )]
        [string[]] $UserEmails,

        [Parameter( ParameterSetName = 'AllUsers' )]
        [switch] $AllUsers,

        [int] $Days = 10,
        [int] $ResultLimit = 50000,
        [string] $TableStyle = 'Dark8',
        [boolean] $Variable = $true,
        [boolean] $Xml = $true,
        [boolean] $Script = $false,
        [boolean] $Excel = $true
    )

    begin {

        #region BEGIN

        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        switch ( $ParameterSet ) {
            'UserObjects' {
                
                # if passed via script argument:
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
                }
            }
            'UserEmails' {

                # variables
                $ScriptUserObjects = [System.Collections.Generic.list[psobject]]::new()

                foreach ( $Email in $UserEmails ) {

                    # create object with userprincipalname property
                    $ScriptUserObjects.Add( 
                        [pscustomobject]@{
                            UserPrincipalName = $Email
                        }
                    )
                }
            }
            'AllUsers' {

                # build user object with null principal name
                $ScriptUserObjects = @(
                    [pscustomobject]@{
                        UserPrincipalName = $null
                    }
                )
            }
        }

        # verify connected to exchange
        try {
            $Domain = Get-AcceptedDomain
        }
        catch {}
        if ( -not $Domain ) {
            Write-Host @Red "${Function}: Not connected to Exchange. Run Connect-ExchangeOnline."
            return
        }
        
        #region CONSTANTS

        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $RawDateProperty = 'Received'

        # get client domain name for file output
        $DefaultDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
        $DomainName = $DefaultDomain.DomainName -split '\.' | Select-Object -First 1

        # get date/time string for filename
        $DateString = Get-Date -Format $FileNameDateFormat
    }

    process {

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            # variables
            $UserEmail = $ScriptUserObject.UserPrincipalName
            if ( $null -eq $UserEmail ) {
                $AllUsers = $true
            }
            
            # build file name
            if ( $AllUsers ) {
                $UserName = 'AllUsers'
            }
            else {
                $UserName = $UserEmail -split '@' | Select-Object -First 1
            }
            $XmlOutputPath = "MessageTrace_${Days}Days_${UserName}_${DateString}.xml"

            ### request message trace records
            # if single user
            if ( $UserEmail ) {

                # get sender records
                Write-Host @Blue "Requesting message trace records with sender: ${UserEmail}"
                $Params = @{
                    SenderAddress = $UserEmail
                    Days = $Days
                    ResultLimit = $ResultLimit
                }
                $SenderTrace = Request-IRTMessageTrace @Params

                # get recipient records
                Write-Host @Blue "Requesting message trace records with recipient: ${UserEmail}"
                $Params = @{
                    RecipientAddress = $UserEmail
                    Days = $Days
                    ResultLimit = $ResultLimit
                }
                $RecipientTrace = Request-IRTMessageTrace @Params

                # if both traces have results
                if (($SenderTrace | Measure-Object).Count -gt 0 -and 
                    ($RecipientTrace | Measure-Object).Count -gt 0
                ) {
                    # merge the two together
                    $MergeParams = @{
                        PropertyName = $RawDateProperty
                        ListOne          = $SenderTrace
                        ListTwo          = $RecipientTrace
                        Descending   = $true
                    }
                    [System.Collections.Generic.List[psobject]]$MessageTrace = Merge-SortedListsOnDate @MergeParams
                }
                # if no results, exit
                elseif (($SenderTrace | Measure-Object).Count -eq 0 -and 
                        ($RecipientTrace | Measure-Object).Count -eq 0
                ) {
                    Write-Host @Red "${Function}: No message trace results found. Exiting."
                    return
                }
                # if only results in one, no need to merge.
                else {
                    $MessageTrace = $SenderTrace + $RecipientTrace
                }
            }
            # if -allusers
            else {
                Write-Host @Blue "Getting message trace records for all users."
                $Params = @{
                    Days = $Days
                    ResultLimit = $ResultLimit
                }
                [System.Collections.Generic.List[psobject]]$MessageTrace = Request-IRTMessageTrace @Params
            }

            # add metadata to results
            $StartDate = (Get-Date).AddDays($Days * -1)
            $EndDate = Get-Date
            $MessageTrace.Insert(0,
                [pscustomobject]@{
                    Metadata = $true
                    UserObject = $ScriptUserObject
                    UserEmail = $UserEmail
                    UserName = $UserName
                    StartDate = $StartDate
                    EndDate = $EndDate
                    Days = $Days
                    DomainName = $DomainName
                }
            )

            #region OUTPUT

            # export to variables
            if ($Variable) {
                if ($AllUsers) {
                    # do nothing
                }
                else {
                    # export raw message trace
                    $VariableName = "IRT_MessageTrace_${UserName}"
                    Write-Host @Blue "Exporting message trace to `$Global:${VariableName}"
                    $VariableParams = @{
                        Name = $VariableName
                        Value = $MessageTrace
                        Scope = 'Global'
                        Force = $True
                    }
                    New-Variable @VariableParams

                    # export table by internetmessageid
                    $Table = @{}
                    foreach ($Trace in $MessageTrace) {
                        if (-not $Trace.Metadata) {
                            $InternetMessageId = $Trace.MessageId
                            $Table[$InternetMessageId] = $Trace
                        }
                    }
                    $VariableName = "IRT_MessageTraceTable_${UserName}"
                    Write-Host @Blue "Exporting message trace to `$Global:${VariableName}"
                    $VariableParams = @{
                        Name = $VariableName
                        Value = $Table
                        Scope = 'Global'
                        Force = $True
                    }
                    New-Variable @VariableParams
                }
            }

            # export raw data
            if ($Xml) {
                Write-Host @Blue "Exporting raw data to: ${XmlOutputPath}"
                $MessageTrace | Export-CliXml -Depth 10 -Path $XmlOutputPath
            }

            if ($Script) {
                Write-Output $MessageTrace
                return
            }

            # create excel sheet
            if ($Excel) {
                Show-IRTMessageTrace -TraceObjects $MessageTrace
            }
        }
    }
}