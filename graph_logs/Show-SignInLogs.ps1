function Show-SignInLogs {
	<#
	.SYNOPSIS
	Processes Sign in log .XML file into Excel spreadsheet.
	
	.NOTES
	Version: 1.1.3
    1.1.3 - Added timers/progress for testing.
	#>
    [CmdletBinding(DefaultParameterSetName='Objects')]
    param (
	    [Parameter(Position=0, Mandatory, ParameterSetName='Objects')]
        [System.Collections.Generic.List[PSObject]] $Logs,

        [Parameter(Mandatory, ParameterSetName='Xml')]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8',

        [boolean] $IpInfo = $true,
        [boolean] $Open = $true,
        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        $RawDateProperty = 'CreatedDateTime'
        $DateColumnHeader = 'DateTime'
        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($Script:Test) {
            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        # $Red = @{ ForegroundColor = 'Red' }
        $Yellow = @{ ForegroundColor = 'Yellow' }

        if ($Test) {$Script:Test = $true}

        # import from xml
        if ($ParameterSet -eq 'Xml') {
            if ($Script:Test) {
                $TestText = "Importing from Xml"
                $TimerStart = $Stopwatch.Elapsed
                Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
            }

            try {
                $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
                [System.Collections.Generic.List[PSObject]]$Logs = Import-CliXml -Path $ResolvedXmlPath
            }
            catch {
                $_
                $ErrorParams = @{
                    Category    = 'ReadError'
                    Message     = "Error importing from ${XmlPath}."
                    ErrorAction = 'Stop'
                }
                Write-Error @ErrorParams
            }

            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
            }
        }

        # import metadata
        if ($Logs[0].Metadata) {

            # remove metadata from beginning of list
            $Metadata = $Logs[0]
            $Logs.RemoveAt(0)

            # $UserEmail = $Metadata.UserEmail
            $UserName = $Metadata.UserName
            $StartDate = $Metadata.StartDate
            $EndDate = $Metadata.EndDate
            $Days = $Metadata.Days
            $DomainName = $Metadata.DomainName
            $FileNamePrefix = $Metadata.FileNamePrefix
        }
        else {
            Write-Host @Red "${Function}: No Metadata found."
        }

        # build file name
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileDateString = $EndDate.ToLocalTime().ToString($FileNameDateFormat)
        $ExcelOutputPath =  "${FileNamePrefix}_${Days}Days_${UserName}_${FileDateString}.xlsx"

        # build worksheet title
        $TitleDateFormat = "M/d/yy h:mmtt"
        $TitleEndDate = $EndDate.ToLocalTime().ToString($TitleDateFormat)
        $TitleStartDate = $StartDate.ToLocalTime().ToString($TitleDateFormat)
        # if allusers, use domain as username
        if ( $UserName -eq 'AllUsers' ) {
            $UserName = $DomainName
        }
        # build title
        if ( $FileNamePrefix -eq 'SignInLogs' ) {
            $WorksheetTitle = "Interactive sign-in logs for ${UserName}. Covers ${Days} days, ${TitleStartDate} to ${TitleEndDate}."
        }
        elseif ( $FileNamePrefix -eq 'NonInteractiveLogs' ) {
            $WorksheetTitle = "Non-Interactive sign-in logs for ${UserName}. Covers ${Days} days, ${TitleStartDate} to ${TitleEndDate}."
        }

        # ipinfo
        if ($IpInfo) {

            $IpInfoAddresses = [System.Collections.Generic.HashSet[string]]::new()

            # check for presence of ip_info package
            $IpInfoPackage = Test-PythonPackage -Name 'ip_info'
        }
    }

    process {
        
        #region FIRST LOOP
        
        foreach ($Log in $Logs) {
            # collect ip addresses
            if ($IpInfo) {
                if ($Log.IpAddress) {
                    try {
                        $IpObject = [System.Net.IPAddress]$Log.IpAddress
                    }
                    catch {}
                    if ($IpObject) {
                        [void]$IpInfoAddresses.Add($IpObject.ToString())
                    }
                }
            }
        }

        #region QUERY IPS
        if ($IpInfo -and 
            $IpInfoPackage.Present -and
            ($IpInfoAddresses | Measure-Object).Count -gt 0
        ) {

            # query information for all IP addresses
            if ($Script:Test) {
                $TestText = "Querying ip info"
                $TimerStart = $Stopwatch.Elapsed
                Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
            }

            $Code = 'import sys; from ip_info.main import cli; sys.exit(cli())'
            & $IpInfoPackage.Python '-c' $Code --apis bulk --output_format none --ip_addresses $IpInfoAddresses
            if ($LASTEXITCODE -ne 0) {
                Write-Host @Red "${Function}: ip_info query failed." | Out-Host
            }

            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
            }

            # add ip info to global colection
            if ($Script:Test) {
                $TestText = "Creating ip info collection in global scope"
                $TimerStart = $Stopwatch.Elapsed
                Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
            }

            if (($Global:IRT_IpInfo | Measure-Object).Count -eq 0) {
                $Global:IRT_IpInfo = @{}
            }
            foreach ($Ip in $IpInfoAddresses) {
                # if ip doesn't exist in table, add it.
                if (-not $Global:IRT_IpInfo.ContainsKey($Ip)) {
                    $Code = 'import sys; from ip_info.main import cli; sys.exit(cli())'
                    $Params = @(
                        '-c', $Code,
                        '--apis','none',
                        '--output_format','table',
                        '--ip_addresses', $Ip.ToString()
                    )
                    $NewLine = [Environment]::NewLine
                    $Output = ((& $IpInfoPackage.Python @Params) -join $NewLine).Trim()
                    $Global:IRT_IpInfo[$Ip] = $Output
                }
            }

            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
            }
        }

        #region Row Loop

        if ($Script:Test) {
            $TestText = "Row loop"
            $TimerStart = $Stopwatch.Elapsed
            Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
        }
    
        # proccess each log
        $RowCount = ($Logs | Measure-Object).Count
        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new($RowCount)
        for ($i = 0; $i -lt $RowCount; $i++) {  
        
            $Log = $Logs[$i]

            # Raw
            $Raw = $Log | ConvertTo-Json -Depth 10

            # IpAddress
            $IpText = if ($Global:IRT_IpInfo.ContainsKey($Log.IpAddress)) {
                $Global:IRT_IpInfo[$Log.IpAddress]
            }
            else {
                $Log.IpAddress
            }
            
            # application display name / resource id
            if ( $Log.AppDisplayName ) {
                $AppDisplayName = $Log.AppDisplayName
            }
            else {
                $AppDisplayName = $Log.ResourceId
            }

            # compress trust
            $Trust = switch ( $Log.DeviceDetail.TrustType ) {
                'Hybrid Azure AD joined' {
                    'Hybrid'
                }
                'Azure AD joined' {
                    'Az Joined'
                }
                'Azure AD registered' {
                    'Az Registered'
                }
                default {
                    $Log.DeviceDetail.TrustType
                }
            }

            # add to list
            $Rows.Add([PSCustomObject]@{
                Raw = $Raw
                $DateColumnHeader = Format-EventDateString $Log.$RawDateProperty
                UserPrincipalName = $Log.UserPrincipalName
                Error = ConvertTo-HumanErrorDescription -ErrorCode $Log.Status.ErrorCode
                IpAddress = $IpText
                City = $Log.Location.City
                State = $Log.Location.State
                Co = $Log.Location.CountryOrRegion
                Application = $AppDisplayName
                Browser = $Log.DeviceDetail.Browser
                OS = $Log.DeviceDetail.OperatingSystem
                Trust = $Trust
                UserAgent = $Log.UserAgent
                Session = $Log.CorrelationId
                Token = $Log.UniqueTokenIdentifier
            })

            if ($Script:Test -and ($i % 1000 -eq 0)) {
                $Percent = [int]( ($i / $RowCount ) * 100 )
                $ProgressParams = @{
                    Id              = 1
                    Activity        = 'Row loop'
                    Status          = "Completed ${i} of ${RowCount}"
                    PercentComplete = $Percent
                }
                Write-Progress @ProgressParams
            }
        }

        if ($Script:Test) {
            Write-Progress -Id 1 -Activity 'Row loop' -Completed

            $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
            Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
        }

        # export spreadsheet
        if ($Script:Test) {
            $TestText = "Exporting to excel"
            $TimerStart = $Stopwatch.Elapsed
            Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
        }

        $ExcelParams = @{
            Path          = $ExcelOutputPath
            WorkSheetname = $FileNamePrefix
            Title         = $WorksheetTitle
            TableStyle    = $TableStyle
            # AutoSize      = $true # apparently very slow?
            FreezeTopRow  = $true
            Passthru      = $true
        }
        try {
            $Workbook = $Rows | Export-Excel @ExcelParams
        }
        catch {
            Write-Error "Unable to open new Excel document."
            if ( Get-YesNo "Try closing open files." ) {
                try {
                    $Workbook = $Rows | Export-Excel @ExcelParams
                }
                catch {
                    throw "Unable to open new Excel document. Exiting."
                }
            }
        }
        $Worksheet = $Workbook.Workbook.Worksheets[$ExcelParams.WorksheetName]

        if ($Script:Test) {
            $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
            Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
        }

        # get table ranges
        $SheetStartColumn = $WorkSheet.Dimension.Start.Column | Convert-DecimalToExcelColumn
        $SheetStartRow = $WorkSheet.Dimension.Start.Row
        $TableStartColumn = ( $workSheet.Tables.Address | Select-Object -First 1 ).Start.Column | Convert-DecimalToExcelColumn
        $TableStartRow = ( $workSheet.Tables.Address | Select-Object -First 1 ).Start.Row
        $EndColumn = $WorkSheet.Dimension.End.Column | Convert-DecimalToExcelColumn
        $EndRow = $WorkSheet.Dimension.End.Row

        $IpAddressColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'IpAddress'}).Id | Convert-DecimalToExcelColumn
        $ApplicationColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'Application'}).Id | Convert-DecimalToExcelColumn
        $UserAgentColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'UserAgent'}).Id | Convert-DecimalToExcelColumn

        #region CELL COLORING

        # ip addresses
        # vpn
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue  = 'vpn'
            BackgroundColor = 'LightPink'
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # tor
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue = 'tor'
            BackgroundColor = 'LightPink'
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # proxy
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue = 'proxy'
            BackgroundColor = 'LightPink'
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # microsoft
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue  = 'microsoft'
            BackgroundColor = 'LightBlue'
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # hosting
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue  = 'hosting'
            BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#FACD90') 
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # cloud
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue  = 'cloud'
            BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#FACD90') 
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams
        # mobile
        $CFParams = @{
            Worksheet       = $WorkSheet
            Address         = "${IpAddressColumn}:${IpAddressColumn}"
            RuleType        = 'ContainsText'
            ConditionValue  = 'mobile'
            BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#F2CEEF') 
            StopIfTrue = $true
        }
        Add-ConditionalFormatting @CFParams

        # applications
        $Strings = @(
            'Azure Active Directory PowerShell'
            'Microsoft Azure CLI'
            'Microsoft Exchange REST API Based Powershell'
            'Microsoft Graph Command Line Tools'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${ApplicationColumn}:${ApplicationColumn}"
                RuleType        = 'Equal'
                ConditionValue  = $String
                BackgroundColor = 'LightPink'
            }
            Add-ConditionalFormatting @CFParams
        }

        # user agents
        $Strings = @(
            'axios'
            'BAV2ROPC'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${UserAgentColumn}:${UserAgentColumn}"
                RuleType        = 'ContainsText'
                ConditionValue  = $String
                BackgroundColor = 'LightPink'
            }
            Add-ConditionalFormatting @CFParams
        }

        #region COLUMN WIDTH

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id
        $Worksheet.Column($Column).Width = 26

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'UserPrincipalName' } ).Id 
        $Worksheet.Column($Column).Width = 30

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Error' } ).Id 
        $Worksheet.Column($Column).Width = 25
        
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'IpAddress' } ).Id
        $Worksheet.Column($Column).Width = 20

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'City' } ).Id
        $Worksheet.Column($Column).Width = 10

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'State' } ).Id
        $Worksheet.Column($Column).Width = 10

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Co' } ).Id
        $Worksheet.Column($Column).Width = 6

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Application' } ).Id
        $Worksheet.Column($Column).Width = 25

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Browser' } ).Id
        $Worksheet.Column($Column).Width = 20

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'OS' } ).Id
        $Worksheet.Column($Column).Width = 12

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Trust' } ).Id
        $Worksheet.Column($Column).Width = 12

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'UserAgent' } ).Id
        $Worksheet.Column($Column).Width = 150

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Session' } ).Id
        $Worksheet.Column($Column).Width = 10

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Token' } ).Id
        $Worksheet.Column($Column).Width = 10

        #region FORMATTING

        # set text wrapping on ip address column
        $WrapParams = @{
            Worksheet = $Worksheet
            Range = "${IpAddressColumn}:${IpAddressColumn}"
            WrapText = $true
        }
        Set-Format @WrapParams

        # set font and size
        $SetParams = @{
            Worksheet = $Worksheet
            Range     = "${SheetStartColumn}${SheetStartRow}:${EndColumn}${EndRow}"
            FontName  = 'Roboto'
        }
        try {
            Set-ExcelRange @SetParams
        } catch {}

        # add left side border
        $BorderParams = @{
            Worksheet = $Worksheet
            Range = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
            BorderLeft = 'Thin'
            BorderColor = 'Black'
        }
        Set-Format @BorderParams

        # set row height
        # $HeightParams = @{
        #     Worksheet = $Worksheet
        #     Row = ($TableStartRow..$EndRow)
        #     Height = 15
        # }
        # Set-ExcelRow @HeightParams
        for ( $i = $TableStartRow; $i -le $EndRow; $i++ ) {
            $Row = $Worksheet.Row($i)
            $Row.Height = 15
            $Row.CustomHeight = $true
        }

        #region OUTPUT
                    
        # save and close
        Write-Host @Blue "Exporting to: ${ExcelOutputPath}"
        if ($Open) {
            Write-Host @Blue "Opening Excel."
            $Workbook | Close-ExcelPackage -Show
        }
        else {
            $Workbook | Close-ExcelPackage
        }
    }
}