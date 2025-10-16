function Show-IRTMessageTrace {
    <#
	.SYNOPSIS
	Processes message trace data and creates spreadsheet.
	
	.NOTES
	Version: 1.0.0
	#>
    [CmdletBinding( DefaultParameterSetName = 'Objects' )]
    param (
        [Parameter(Position = 0, Mandatory, ValueFromPipeline, ParameterSetName = 'Objects')]
        [Alias( 'Message' )]
        [System.Collections.Generic.List[PSObject]] $Messages,

        [Parameter(Position = 0, Mandatory, ParameterSetName = 'Xml')]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8',

        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        if ($Test) {
            $Script:Test = $true
            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }
        $TitleDateFormat = "M/d/yy h:mmtt"
        $RawDateProperty = 'Received'
        $DateColumnHeader = 'DateTime'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        $Yellow = @{ ForegroundColor = 'Yellow' }

        if ($Test) {}

        # import from xml
        if ($ParameterSet -eq 'Xml') {
            if ($Script:Test) {
                $TestText = "Importing from Xml"
                $TimerStart = $Stopwatch.Elapsed
                Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
            }

            try {
                $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
                [System.Collections.Generic.List[PSObject]]$Messages = Import-CliXml -Path $ResolvedXmlPath
            }
            catch {
                $_
                Write-Host @Red "${Function}: Error importing from ${XmlPath}."
                return
            }

            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $TimerStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: ${TestText} took ${ElapsedString}" | Out-Host
            }
        }

        # import metadata
        if ($Messages[0].Metadata) {

            # remove metadata from beginning of list
            $Metadata = $Messages[0]
            $Messages.RemoveAt(0)

            $UserEmail = $Metadata.UserEmail
            $UserName = $Metadata.UserName
            $StartDate = $Metadata.StartDate
            $EndDate = $Metadata.EndDate
            $Days = $Metadata.Days
            $DomainName = $Metadata.DomainName
            $FileNamePrefix = $Metadata.FileNamePrefix
        }
        else {
            Write-Error "${Function}: No Metadata found."
        }

        # build file name
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileDateString = $EndDate.ToLocalTime().ToString($FileNameDateFormat)
        $ExcelOutputPath = "${FileNamePrefix}_${Days}Days_${UserName}_${FileDateString}.xlsx"

        # build worksheet title
        $StartString = $StartDate.ToString($TitleDateFormat).ToLower()
        $EndString = $EndDate.ToString($TitleDateFormat).ToLower()
        if ($null -eq $UserEmail) {
            $WorksheetTitle = "Message Trace for ${DomainName}. Covers ${Days} days, from ${StartString} to ${EndString}."
        }
        else {
            $WorksheetTitle = "Message Trace for ${UserEmail}. Covers ${Days} days, from ${StartString} to ${EndString}."
        }
    }

    process {

        #region ROW LOOP

        if ($Script:Test) {
            $TestText = "Row loop"
            $TimerStart = $Stopwatch.Elapsed
            Write-Host @Yellow "${Function}: ${TestText} started at $(Get-Date -Format 'hh:mm:sstt')" | Out-Host
        }

        $RowCount = $Messages.Count
        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new($RowCount)
        for ($i = 0; $i -lt $RowCount; $i++) {

            $Message = $Messages[$i]

            # Raw
            $Raw = $Message | ConvertTo-Json -Depth 10

            $Rows.Add([pscustomobject]@{
                Raw               = $Raw
                $DateColumnHeader = (Format-EventDateString $Message.$RawDateProperty)
                Status            = $Message.Status
                SenderAddress     = $Message.SenderAddress
                RecipientAddress  = $Message.RecipientAddress
                Subject           = $Message.Subject
                FromIP            = $Message.FromIP
                ToIP              = $Message.ToIP
                MessageTraceId    = $Message.MessageTraceId
                MessageId         = $Message.MessageId
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

        #region EXPORT EXCEL
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
        $TableStartColumn = ($Worksheet.Tables.Address | Select-Object -First 1).Start.Column | Convert-DecimalToExcelColumn
        $TableStartRow = ($Worksheet.Tables | Select-Object -First 1).Address.Start.Row + 1
        $EndColumn = $WorkSheet.Dimension.End.Column | Convert-DecimalToExcelColumn
        $EndRow = $WorkSheet.Dimension.End.Row

        $SenderColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'SenderAddress'}).Id | Convert-DecimalToExcelColumn
        $RecipientColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'RecipientAddress'}).Id | Convert-DecimalToExcelColumn

        #region BOLD OTHER EMAIL

        if ($UserEmail) {
            $CfParamsSender = @{
                WorkSheet        = $Worksheet
                Address          = "${SenderColumn}${TableStartRow}:${SenderColumn}${EndRow}"
                RuleType         = 'NotEqual'
                ConditionValue   = $UserEmail
                Bold = $true
            }
            Add-ConditionalFormatting @CfParamsSender

            $CfParamsRecipient = @{
                WorkSheet        = $Worksheet
                Address          = "${RecipientColumn}${TableStartRow}:${RecipientColumn}${EndRow}"
                RuleType         = 'NotEqual'
                ConditionValue   = $UserEmail
                Bold = $true
            }
            Add-ConditionalFormatting @CfParamsRecipient
        }

        #region SAME TO/FROM

        $CfParams = @{
            WorkSheet        = $Worksheet
            Address          = "${SenderColumn}${TableStartRow}:${RecipientColumn}${EndRow}"
            RuleType         = 'Expression'
            ConditionValue   = "=`$${SenderColumn}${TableStartRow}=`$${RecipientColumn}${TableStartRow}"
            BackgroundColor  = 'LightYellow'
        }
        Add-ConditionalFormatting @CfParams

        #region COLUMN WIDTH

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id 
        $Worksheet.Column($Column).Width = 26

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Status' } ).Id 
        $Worksheet.Column($Column).Width = 15

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'SenderAddress' } ).Id 
        $Worksheet.Column($Column).Width = 30

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'RecipientAddress' } ).Id 
        $Worksheet.Column($Column).Width = 30

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Subject' } ).Id 
        $Worksheet.Column($Column).Width = 100

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'FromIp' } ).Id 
        $Worksheet.Column($Column).Width = 20

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'ToIp' } ).Id 
        $Worksheet.Column($Column).Width = 20

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'MessageTraceId' } ).Id
        $Worksheet.Column($Column).Width = 20

        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'MessageTraceId' } ).Id
        $Worksheet.Column($Column).Width = 200

        #region FORMATTING

        # set font and size
        $SetParams = @{
            Worksheet = $Worksheet
            Range     = "${SheetStartColumn}${SheetStartRow}:${EndColumn}${EndRow}"
            FontName  = 'Consolas'
        }
        Set-ExcelRange @SetParams

        # add left side border
        $BorderParams = @{
            Worksheet = $Worksheet
            Range = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
            BorderLeft = 'Thin'
            BorderColor = 'Black'
        }
        Set-Format @BorderParams

        #region OUTPUT

        # save and close
        Write-Host @Blue "${Function}: Exporting to: ${ExcelOutputPath}"
        $Workbook | Close-ExcelPackage -Show

        if ($Script:Test) {
            $ElapsedString = ($StopWatch.Elapsed).ToString('mm\:ss')
            Write-Host @Yellow "${Function} took ${ElapsedString}" | Out-Host
        }
    }
}