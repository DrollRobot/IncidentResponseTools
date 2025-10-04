function Show-SignInLogs {
	<#
	.SYNOPSIS
	Processes Sign in log .XML file into Excel spreadsheet.
	
	.NOTES
	Version: 1.1.2
	#>
    [CmdletBinding()]
    param (
        [Parameter( Position = 0 )]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8',
        [boolean] $Open = $true,
        [switch] $Test
    )

    begin {

        #region BEGIN

        if ($Test) {
            $Global:IRTTestMode = $true
        }

        # get file path
        if ( $XmlPath ) {

            $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
            $Logs = Import-Clixml -Path $ResolvedXmlPath
        }
        else {
        
            # run import-logs to get file name
            $ImportParams = @{
                Pattern    = "^(SignInLogs|NonInteractiveLogs)_Raw_.*\.xml$"
                ReturnPath = $true
            }
            $ResolvedXmlPath = Import-LogFile @ImportParams
        
            # use path to import logs
            $Logs = Import-Clixml -Path $ResolvedXmlPath
        }
        
        #region CONSTANTS

        $OutputTable = [System.Collections.Generic.List[PSCustomObject]]::new()

        # file variables
        $WorksheetName = 'SignInLogs'
        $FileNameDatePattern = "\d{2}-\d{2}-\d{2}_\d{2}-\d{2}"
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $TitleDateFormat = "M/d/yy h:mmtt"

        # event date formatting
        $RawDateProperty = 'CreatedDateTime'
        $DateColumnHeader = 'DateTime'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        # build new file name out of old one
        $OldFileName = Split-Path -Path $ResolvedXmlPath -Leaf
        $SplitFileName = $OldFileName -split '_'
        $SplitFileName = $SplitFileName | Where-Object { $_ -ne 'Raw' }
        $UserString = $SplitFileName[3]
        $SplitFileName = $SplitFileName -replace '\.xml', '.xlsx'
        $ExcelOutputPath = $SplitFileName -join '_'
        
        ### build worksheet title
        # get number of days
        $ExcelOutputPath -match "(\d{1,3})Days" | Out-Null
        $Days = $Matches[1]
        # get date range
        $QueryDateString = $ExcelOutputPath | Select-String -Pattern $FileNameDatePattern -AllMatches | ForEach-Object { $_.Matches.Value }
        $ParsedDate = [DateTime]::ParseExact( $QueryDateString, $FileNameDateFormat, $null )
        $StartString = $ParsedDate.AddDays([int]$Days * -1).ToString( $TitleDateFormat ).ToLower()
        $EndString = $ParsedDate.ToString( $TitleDateFormat ).ToLower()
        # get username
        if ( $UserString -eq 'AllUsers' ) {
            # if all users, use domain as username
            $UserString = $SplitFileName[2]
        }
        # build title
        if ( $SplitFileName[0] -eq 'SignInLogs' ) {
            $WorksheetTitle = "Interactive sign-in logs for ${UserString}. Covers ${Days} days, ${StartString} to ${EndString}."
        }
        elseif ( $SplitFileName[0] -eq 'NonInteractiveLogs' ) {
            $WorksheetTitle = "Non-Interactive sign-in logs for ${UserString}. Covers ${Days} days, ${StartString} to ${EndString}."
        }
    }

    process {

        # add error description
        $Logs = Add-HumanErrorDescription -Logs $Logs            
    
        # proccess each log
        for ($i = 0; $i -lt $Logs.Count; $i++) {  
        
            $Log = $Logs[$i]
            $CustomObject = [PSCustomObject]@{}

            # Raw
            $Raw = $Log | ConvertTo-Json -Depth 10
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Raw'
                Value      = $Raw
            }
            $CustomObject | Add-Member @AddParams

            # Date/Time
            $AddParams = @{
                MemberType  = 'NoteProperty'
                Name        = $DateColumnHeader
                Value       = Format-EventDateString $Log.$RawDateProperty
            }
            $CustomObject | Add-Member @AddParams

            # user
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'UserPrincipalName'
                Value      = $Log.UserPrincipalName
            }
            $CustomObject | Add-Member @AddParams

            # error
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Error'
                Value      = $Log.Error
            }
            $CustomObject | Add-Member @AddParams

            # IpAddress
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'IpAddress'
                Value      = $Log.IpAddress
            }
            $CustomObject | Add-Member @AddParams

            # City
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'City'
                Value      = $Log.Location.City
            }
            $CustomObject | Add-Member @AddParams

            # State
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'State'
                Value      = $Log.Location.State
            }
            $CustomObject | Add-Member @AddParams

            # country
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Co'
                Value      = $Log.Location.CountryOrRegion
            }
            $CustomObject | Add-Member @AddParams
            
            # application display name / resource id
            if ( $Log.AppDisplayName ) {
                $AppDisplayName = $Log.AppDisplayName
            }
            else {
                $AppDisplayName = $Log.ResourceId
            }
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Application'
                Value      = $AppDisplayName
            }
            $CustomObject | Add-Member @AddParams

            # Browser
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Browser'
                Value      = $Log.DeviceDetail.Browser
            }
            $CustomObject | Add-Member @AddParams

            # OS
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'OS'
                Value      = $Log.DeviceDetail.OperatingSystem
            }
            $CustomObject | Add-Member @AddParams

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
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Trust'
                Value      = $Trust
            }
            $CustomObject | Add-Member @AddParams

            # # add human readable useragent
            # $AddParams = @{
            #     Logs         = $Logs
            #     Property     = 'UserAgent'
            #     ColumnHeader = 'UserAgentHuman'
            # }
            # $Logs = Add-HumanReadableId @AddParams

            # # UserAgentHuman
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'UserAgent'
            #     Value      = $Log.UserAgentHuman
            # }
            # $CustomObject | Add-Member @AddParams

            # UserAgent (raw value)
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'UserAgent'
                Value      = $Log.UserAgent
            }
            $CustomObject | Add-Member @AddParams

            # # add human readable session id
            # $AddParams = @{
            #     Logs         = $Logs
            #     Property     = 'CorrelationID'
            #     ColumnHeader = 'SessionHuman'
            # }
            # $Logs = Add-HumanReadableId @AddParams

            # # SessionHuman
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'Session'
            #     Value      = $Log.SessionHuman
            # }
            # $CustomObject | Add-Member @AddParams

            # Session (raw value)
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Session'
                Value      = $Log.CorrelationId
            }
            $CustomObject | Add-Member @AddParams

            # # add human readable token id
            # $AddParams = @{
            #     Logs         = $Logs
            #     Property     = 'UniqueTokenIdentifier'
            #     ColumnHeader = 'TokenHuman'
            # }
            # $Logs = Add-HumanReadableId @AddParams

            # # Token
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'Token'
            #     Value      = $Log.TokenHuman
            # }
            # $CustomObject | Add-Member @AddParams

            # Token (raw value)
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Token'
                Value      = $Log.UniqueTokenIdentifier
            }
            $CustomObject | Add-Member @AddParams

            # add to list
            $OutputTable.Add( $CustomObject )
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
            $Workbook = $OutputTable | Export-Excel @ExcelParams
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

        #region CELL COLORING

        # if cell matches EXACTLY, make background RED
        $Strings = @(
            'Azure Active Directory PowerShell'
            'Microsoft Azure CLI'
            'Microsoft Exchange REST API Based Powershell'
            'Microsoft Graph Command Line Tools'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
                RuleType        = 'Equal'
                ConditionValue  = $String
                BackgroundColor = 'LightPink'
            }
            Add-ConditionalFormatting @CFParams
        }

        # if cell CONTAINSTEXT, make background RED
        $Strings = @(
            'axios'
            'BAV2ROPC'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
                RuleType        = 'ContainsText'
                ConditionValue  = $String
                BackgroundColor = 'LightPink'
            }
            Add-ConditionalFormatting @CFParams
        }

        #region COLUMN WIDTH

        # resize DateTime column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id
        $Worksheet.Column($Column).Width = 26

        # resize Raw column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8
        
        # resize Co column (Country)
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Co' } ).Id
        $Worksheet.Column($Column).Width = 6

        # resize Session column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Session' } ).Id
        $Worksheet.Column($Column).Width = 10

        # resize Token column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Token' } ).Id
        $Worksheet.Column($Column).Width = 10

        #region FORMATTING

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