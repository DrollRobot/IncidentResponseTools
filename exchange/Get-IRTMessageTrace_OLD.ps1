New-Alias -Name 'MessageTraceOLD' -Value 'Get-IRTMessageTrace_OLD' -Force
function Get-IRTMessageTrace_OLD {
    <#
	.SYNOPSIS
	Downloads incoming and outgoing message trace for specified user, or all users.
	
	.NOTES
	Version: 1.3.0
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
        [boolean] $Open = $true
    )

    begin {

        #region BEGIN

        $ParameterSet = $PSCmdlet.ParameterSetName

        switch ( $ParameterSet ) {
            'UserObjects' {
                
                # if passed via parameter
                if ( $UserObjects ) {
                    $ScriptUserObjects = $UserObjects
                }
                # if not, find global objects
                else {
                    
                    # get from global variables
                    $ScriptUserObjects = Get-GraphGlobalUserObjects
                    
                    # if none found, exit
                    if ( -not $ScriptUserObjects -or $ScriptUserObjects.Count -eq 0 ) {
                        throw "No user objects passed or found in global variables."
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
            throw "Not connected to ExchangeOnlineManagement. Run Connect-ExchangeOnline. Exiting."
        }
        
        #region CONSTANTS

        $WorksheetName = 'MessageTrace'
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
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

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

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
            $StartString = (Get-Date).AddDays($Days * -1).ToString($TitleDateFormat).ToLower()
            $EndString = (Get-Date).ToString($TitleDateFormat).ToLower()
            if ( $AllUsers ) {
                $WorksheetTitle = "Message Trace for ${DomainName}. Covers ${Days} days, from ${StartString} to ${EndString}."
            }
            else {
                $WorksheetTitle = "Message Trace for ${UserEmail}. Covers ${Days} days, from ${StartString} to ${EndString}."
            }

            ### request message trace records
            # if single user
            if ( $UserEmail ) {

                # get sender records
                Write-Host @Blue "Getting message trace records with sender: ${UserEmail}"
                $Params = @{
                    SenderAddress = $UserEmail
                    Days = $Days
                    ResultLimit = $ResultLimit
                }
                $SenderTrace = Request-IRTMessageTrace @Params

                # get recipient records
                Write-Host @Blue "Getting message trace records with recipient: ${UserEmail}"
                $Params = @{
                    RecipientAddress = $UserEmail
                    Days = $Days
                    ResultLimit = $ResultLimit
                }
                $RecipientTrace = Request-IRTMessageTrace @Params

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
                    $MessageTrace = Merge-SortedListsOnDate @MergeParams
                }
                # if no results, exit
                elseif ( @( $SenderTrace ).Count -eq 0 -and @( $RecipientTrace ).Count -eq 0 ) {

                    Write-Host @Red "No message trace results found. Exiting."
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
                $MessageTrace = Request-IRTMessageTrace @Params
            }

            #region ROW LOOP

            for ($i = 0; $i -lt $MessageTrace.Count; $i++) {

                $Message = $MessageTrace[$i]

                # DateTime
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
                $Message | Add-Member @AddParams
            }

            # export raw data
            Write-Host @Blue "Exporting raw data to: ${XmlOutputPath}"
            $MessageTrace | Export-CliXml -Depth 10 -Path $XmlOutputPath

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
                $Workbook = $MessageTrace | 
                    Select-Object $DisplayProperties | 
                    Export-Excel @ExcelParams
            }
            catch {
                Write-Error "Unable to open new Excel document."
                if ( Get-YesNo "Try closing open files." ) {
                    try {
                        $Workbook = $MessageTrace | 
                            Select-Object $DisplayProperties | 
                            Export-Excel @ExcelParams
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