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
        [Alias( 'TraceObject' )]
        [System.Collections.Generic.List[PSObject]] $TraceObjects,

        [Parameter(Position = 0, Mandatory, ParameterSetName = 'Xml')]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8'
    )

    begin {

        #region BEGIN

        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # import from xml
        if ($ParameterSet -eq 'Xml') {
            try {
                $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
                $TraceObjects = Import-CliXml -Path $ResolvedXmlPath
            }
            catch {
                $_
                Write-Host @Red "${Function}: Error importing from ${XmlPath}."
                return
            }
        }

        # constants
        $WorksheetName = 'MessageTrace'
        $TitleDateFormat = "M/d/yy h:mmtt"
        $RawDateProperty = 'Received'
        $DateColumnHeader = 'DateTime'

        # import metadata
        if ($TraceObjects[0].Metadata) {

            # remove metadata from beginning of list
            $Metadata = $TraceObjects[0]
            $TraceObjects.RemoveAt(0)

            $UserEmail = $Metadata.UserEmail
            $UserName = $Metadata.UserName
            $StartDate = $Metadata.StartDate
            $EndDate = $Metadata.EndDate
            $Days = $Metadata.Days
        }
        else {
            Write-Host @Red "${Function}: No Metadata found."
        }

        # build file name
        $ExcelOutputPath = "MessageTrace_${Days}Days_${UserName}_${DateString}.xlsx"

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

        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new()
        for ($i = 0; $i -lt ($TraceObjects | Measure-Object).Count; $i++) {

            $Message = $TraceObjects[$i]
            $Row = [PSCustomObject]@{}

            # Raw
            $Raw = $Message | ConvertTo-Json -Depth 10
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Raw'
                Value      = $Raw
            }
            $Row | Add-Member @AddParams

            # Date/Time
            $AddParams = @{
                MemberType  = 'NoteProperty'
                Name        = $DateColumnHeader
                Value       = Format-EventDateString $Message.$RawDateProperty
            }
            $Row | Add-Member @AddParams

            # Status
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Status'
                Value      = $Message.Status
            }
            $Row | Add-Member @AddParams

            # SenderAddress
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'SenderAddress'
                Value      = $Message.SenderAddress
            }
            $Row | Add-Member @AddParams

            # RecipientAddress
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'RecipientAddress'
                Value      = $Message.RecipientAddress
            }
            $Row | Add-Member @AddParams

            # Subject
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Subject'
                Value      = $Message.Subject
            }
            $Row | Add-Member @AddParams

            # FromIP
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'FromIP'
                Value      = $Message.FromIP
            }
            $Row | Add-Member @AddParams

            # ToIP
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'ToIP'
                Value      = $Message.ToIP
            }
            $Row | Add-Member @AddParams

            # MessageTraceId
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'MessageTraceId'
                Value      = $Message.MessageTraceId
            }
            $Row | Add-Member @AddParams

            # MessageId
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'MessageId'
                Value      = $Message.MessageId
            }
            $Row | Add-Member @AddParams

            $Rows.Add($Row)
        }

        # export spreadsheet
        $ExcelParams = @{
            Path          = $ExcelOutputPath
            WorkSheetname = $WorksheetName
            Title         = $WorksheetTitle
            TableStyle    = $TableStyle
            AutoSize      = $true
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

        # get table ranges
        $SheetStartColumn = $WorkSheet.Dimension.Start.Column | Convert-DecimalToExcelColumn
        $SheetStartRow = $WorkSheet.Dimension.Start.Row
        $TableStartColumn = ( $workSheet.Tables.Address | Select-Object -First 1 ).Start.Column | Convert-DecimalToExcelColumn
        $TableStartRow = $Worksheet.Tables[0].Address.Start.Row + 1
        $EndColumn = $WorkSheet.Dimension.End.Column | Convert-DecimalToExcelColumn
        $EndRow = $WorkSheet.Dimension.End.Row

        #region SAME TO/FROM
        # highlight where sender and recipient are the same
        $SenderColumn = ( $Worksheet.Tables[0].Columns | 
            Where-Object { $_.Name -eq 'SenderAddress' } ).Id | 
            Convert-DecimalToExcelColumn
        $RecipientColumn = ( $Worksheet.Tables[0].Columns | 
            Where-Object { $_.Name -eq 'RecipientAddress' } ).Id | 
            Convert-DecimalToExcelColumn

        for ($Row = $TableStartRow; $Row -le $EndRow; $Row++) {

            $SenderAddress = $Worksheet.Cells[$SenderColumn + $Row].Text
            $RecipientAddress = $Worksheet.Cells[$RecipientColumn + $Row].Text

            # highlight where sender and recipient are the same
            if ($SenderAddress -eq $RecipientAddress) {
                $ColorParams = @{
                    Worksheet = $Worksheet
                    Range = $SenderColumn + $Row + ":" + $RecipientColumn + $Row
                    BackgroundColor = 'LightPink'
                }
                Set-ExcelRange @ColorParams
            }

            # bold user email, if not -AllUsers
            if ( $ScriptUserObject.UserPrincipalName ) {
                if ( $SenderAddress -ne $ScriptUserObject.UserPrincipalName ) {
                    $BoldParams = @{
                        Worksheet = $Worksheet
                        Range = $SenderColumn + $Row
                        Bold = $true
                    }
                    Set-ExcelRange @BoldParams
                }
                if ( $RecipientAddress -ne $ScriptUserObject.UserPrincipalName ) {
                    $BoldParams = @{
                        Worksheet = $Worksheet
                        Range = $RecipientColumn + $Row
                        Bold = $true
                    }
                    Set-ExcelRange @BoldParams
                }
            }
        }

        #region COLUMN WIDTH

        # resize DateTime column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id 
        $Worksheet.Column($Column).Width = 26

        # resize Raw column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8

        # resize MessageTraceId column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'MessageTraceId' } ).Id
        $Worksheet.Column($Column).Width = 20

        #region FORMATTING

        # # set row height
        # for ( $i = $TableStartRow; $i -le $EndRow; $i++ ) {  
        #     $workSheet.Row($i).CustomHeight = 15
        # }
        # FIXME maybe not needed?

        # set font and size
        $SetParams = @{
            Worksheet = $Worksheet
            Range     = "${SheetStartColumn}${SheetStartRow}:${EndColumn}${EndRow}"
            FontName  = 'Roboto'
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
    }
}