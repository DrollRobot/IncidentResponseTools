New-Alias -Name 'MessageTraceV1' -Value 'Get-IRTMessageTraceV1' -Force
function Get-IRTMessageTraceV1 {
    <#
	.SYNOPSIS
	Downloads incoming and outgoing message trace for provided users, merges into one array, saves raw xml, then saves as excel spreadsheet.
	
	.NOTES
	Version: 1.2.1
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
        [int] $PageLimit = 10, # 1000 is server-side page limit, 200 represents 1m lines, the max for excel
        [string] $TableStyle = 'Dark8',
        [boolean] $Open = $true
    )

    begin {

        #region BEGIN

        # constants
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        $StartDate = ( Get-Date ).AddDays( $Days * -1 )
        $EndDate = Get-Date
        $WorksheetName = 'MessageTrace'
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $DateString = Get-Date -Format $FileNameDateFormat
        $TitleDateFormat = "M/d/yy h:mmtt"
        $RawDateProperty = 'Received'
        $DateColumnHeader = 'DateTime'
        $DisplayProperties = @(
            $DateColumnHeader
            'Raw'
            'Status'
            'Size'
            'SenderAddress'
            'RecipientAddress'
            'Subject'
            'FromIP'
            'ToIP'
            'MessageTraceId'
            'MessageId'
        )

        $OutputTable = [System.Collections.Generic.List[psobject]]::new()

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }


        switch ( $ParameterSet ) {
            'UserObjects' {
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

        #region USER LOOP

        foreach ( $ScriptUserObject in $ScriptUserObjects ) {

            # variables
            $UserEmail = $ScriptUserObject.UserPrincipalName
            if ( $null -eq $UserEmail ) {
                $AllUsers = $true
            }
            
            # build file names
            if ( $AllUsers ) {
                $UserName = 'AllUsers'
            }
            else {
                $UserName = $UserEmail -split '@' | Select-Object -First 1
            }
            $XmlOutputPath = "MessageTrace_Raw_${Days}Days_${DomainName}_${UserName}_${DateString}.xml"
            $ExcelOutputPath = "MessageTrace_${Days}Days_${DomainName}_${UserName}_${DateString}.xlsx"

            # build worksheet title
            $StartString = $StartDate.ToString( $TitleDateFormat ).ToLower()
            $EndString = $EndDate.ToString( $TitleDateFormat ).ToLower()
            if ( $AllUsers ) {
                $WorksheetTitle = "Message Trace for ${DomainName}. Covers ${Days} days, from ${StartString} to ${EndString}."
            }
            else {
                $WorksheetTitle = "Message Trace for ${UserEmail}. Covers ${Days} days, from ${StartString} to ${EndString}."
            }

            # if user objects or emails were provided
            if ( $UserEmail ) {

                # get sender records
                Write-Host @Blue "Getting message trace records with sender: ${UserEmail}"
                $Params = @{
                    SenderAddress = $UserEmail
                    StartDate = $StartDate
                    EndDate = $EndDate
                    PageLimit = $PageLimit
                }
                $SenderTrace = Get-MessageTraceWithPaging @Params

                # get recipient records
                Write-Host @Blue "Getting message trace records with recipient: ${UserEmail}"
                $Params = @{
                    RecipientAddress = $UserEmail
                    StartDate = $StartDate
                    EndDate = $EndDate
                    PageLimit = $PageLimit
                }
                $RecipientTrace = Get-MessageTraceWithPaging @Params

                # if both traces have results
                if ( @( $SenderTrace ).Count -gt 0 -and @( $RecipientTrace ).Count -gt 0 ) {

                    # merge the two together
                    Write-Host @Blue "Merging results."
                    $MergeParams = @{
                        PropertyName = $RawDateProperty
                        ListOne          = $SenderTrace
                        ListTwo          = $RecipientTrace
                        Descending   = $true
                    }
                    $OutputTable = Merge-SortedListsOnDate @MergeParams
                }
                # if no results, exit
                elseif ( @( $SenderTrace ).Count -eq 0 -and @( $RecipientTrace ).Count -eq 0 ) {

                    Write-Host @Red "No message trace results found. Exiting."
                    return
                }
                # if only results in one, no need to merge.
                else {

                    $OutputTable = $SenderTrace + $RecipientTrace
                }
            }
            # if all users
            else {

                Write-Host @Blue "Getting message trace records for all users."
                $Params = @{
                    StartDate = $StartDate
                    EndDate = $EndDate
                    PageLimit = $PageLimit
                }
                $OutputTable = Get-MessageTraceWithPaging @Params
            }

            #region ROW LOOP

            for ($i = 0; $i -lt $MessageTrace.Count; $i++) {

                $Message = $MessageTrace[$i]

                # Date/Time
                $AddParams = @{
                    MemberType  = 'NoteProperty'
                    Name        = $DateColumnHeader
                    Value       = Format-EventDateString $Message.$RawDateProperty
                }
                $Message | Add-Member @AddParams

                # Raw
                $Raw = $Message | ConvertTo-Json -Depth 10
                $AddParams = @{
                    MemberType = 'NoteProperty'
                    Name       = 'Raw'
                    Value      = $Raw
                }
                $CustomObject | Add-Member @AddParams
            }

            # export raw data
            Write-Host @Blue "Exporting raw data to: ${XmlOutputPath}"
            $OutputTable | Export-CliXml -Depth 10 -Path $XmlOutputPath

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
                $Workbook = $OutputTable | Select-Object $DisplayProperties | Export-Excel @ExcelParams
            }
            catch {
                Write-Error "Unable to open new Excel document."
                if ( Get-YesNo "Try closing open files." ) {
                    try {
                        $Workbook = $OutputTable | Export-Excel @ExcelParams
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
            $TableStartRow = ( $workSheet.Tables.Address | Select-Object -First 1 ).Start.Row
            $EndColumn = $WorkSheet.Dimension.End.Column | Convert-DecimalToExcelColumn
            $EndRow = $WorkSheet.Dimension.End.Row
            $SenderColumn = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'SenderAddress' } ).Id | Convert-DecimalToExcelColumn
            $RecipientColumn = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'RecipientAddress' } ).Id | Convert-DecimalToExcelColumn

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
            $Worksheet.Column($Column).Width = 25

            # resize Raw column
            $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
            $Worksheet.Column($Column).Width = 8

            # resize MessageTraceId column
            $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'MessageTraceId' } ).Id
            $Worksheet.Column($Column).Width = 20

            #region FORMATTING

            # # set text wrapping in description column
            # $WrappingParams = @{
            #     Worksheet = $Worksheet
            #     Range     = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
            #     WrapText  = $true
            # }
            # Set-ExcelRange @WrappingParams

            # set row height
            for ( $i = $TableStartRow; $i -le $EndRow; $i++ ) {  
                $workSheet.Row($i).CustomHeight = 15
            }

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
            Write-Host @Blue "Exporting to: ${ExcelOutputPath}"
            if ( $Open ) {
                Write-Host @Blue "Opening Excel."
                $Workbook | Close-ExcelPackage -Show
            }
            else {
                $Workbook | Close-ExcelPackage
            }
        }
    }
}



