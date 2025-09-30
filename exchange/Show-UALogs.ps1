function Show-UALogs {
    <#
	.SYNOPSIS
	Parse and show unified audit logs.
	
	.NOTES
	Version: 1.0.1
    1.0.1 - Added option pass raw log objects, not just import from file.
	#>
    [CmdletBinding(DefaultParameterSetName='Objects')]
    param (
	    [Parameter(Position=0, Mandatory, ParameterSetName='Objects')]
        [psobject] $Logs, 
	    [Parameter(Mandatory, ParameterSetName='Objects')]
        [string] $DomainName, 
	    [Parameter(Mandatory, ParameterSetName='Objects')]
        [string] $UserName, 
        [Parameter(Mandatory, ParameterSetName='Objects')]
        [int] $Days, 
	    [Parameter(Mandatory, ParameterSetName='Objects')]
        [datetime] $EndDate,
	    [Parameter(ParameterSetName='Objects')]
        [string] $OperationString, 

        [Parameter(Position=0, Mandatory, ParameterSetName='Xml')]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8',
        [boolean] $Xml = $true,
        [boolean] $Open = $true
    )

    begin {
        
        # get file path
        if ($PSCmdlet.ParameterSetName -eq 'Xml') {

            $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
            $Logs = Import-Clixml -Path $ResolvedXmlPath
        }

        # variables
        $ModulePath = $PSScriptRoot
        $OutputTable = [System.Collections.Generic.List[PSCustomObject]]::new()
        $LogCount = ($Logs | Measure-Object).Count

        # $Groups = Request-GraphGroups
        # $Roles = Request-DirectoryRoles
        # $RoleTemplates = Request-DirectoryRoleTemplates
        # $ServicePrincipals = Request-GraphServicePrincipals
        # $Users = Request-GraphUsers

        # event date formatting
        $RawDateProperty = 'CreationDate'
        $DateColumnHeader = 'DateTime'

        $DisplayProperties = @(
            'Index'
            $DateColumnHeader
            'Raw'
            # 'Tree'
            'UserIds'
            'IpAddresses'
            'RecordType'
            'OperationFriendlyName'
            'Summary'
        )

        # import all operations csv
        $OperationsCsvPath = Join-Path -Path $ModulePath -ChildPath '\unified_audit_log-data\unified_audit_log-all_operations.csv'
        $OperationsData = Import-Csv -Path $OperationsCsvPath

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        # $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileNameDatePattern = "\d{2}-\d{2}-\d{2}_\d{2}-\d{2}"
        $TitleDateFormat = "M/d/yy h:mmtt"

        if ($PSCmdlet.ParameterSetName -eq 'Objects') {

            # build new file name
            if ([string]::IsNullOrWhiteSpace($OperationString)) {
                $OperationString = 'UnifiedAuditLogs'
            }
            else {
                $OperationString = "${OperationString}_UAL"
            }
            $FileEndDate = $EndDate.ToLocalTime().ToString($FileNameDateFormat)
            $ExcelOutputPath =  "${OperationString}_${Days}Days_${DomainName}_${UserName}_${FileEndDate}.xlsx"

            # build worksheet title
            $TitleEndDate = $EndDate.ToLocalTime().ToString($TitleDateFormat)
            $TitleStartDate = $EndDate.AddDays($Days * -1).ToLocalTime().ToString($TitleDateFormat)
            $WorksheetTitle = "Unified audit logs for ${UserName}. Covers ${Days} days, ${TitleStartDate} to ${TitleEndDate}."
        }

        if ($PSCmdlet.ParameterSetName -eq 'Xml') {

            # build new file name out of old one
            $OldFileName = Split-Path -Path $ResolvedXmlPath -Leaf
            $SplitFileName = $OldFileName -split '_'
            $SplitFileName = $SplitFileName | Where-Object { $_ -ne 'Raw' }
            $TargetString = $SplitFileName[3]
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
            if ( $TargetString -eq 'AllLogs' ) {
                # if all logs, use domain as target
                $TargetString = $SplitFileName[2]
            }
            # build worksheet title
            $WorksheetTitle = "Unified audit logs for ${TargetString}. Covers ${Days} days, ${StartString} to ${EndString}."
        }
    }

    process {

        # convert audit data to powershell objects
        foreach ($Log in $Logs) {
            $Log.AuditData = $Log.AuditData | ConvertFrom-Json -Depth 10
        }

        #region ROW LOOP

        # process each log
        for ($i = 0; $i -lt $LogCount; $i++) { 
        
            $Log = $Logs[$i]
            $CustomObject = [PSCustomObject]@{
                Index = $i
            }

            # extract auditdata
            $AuditData = $Log.AuditData # FIXME transition to using original log object converted from json

            # Date/Time
            $AddParams = @{
                MemberType  = 'NoteProperty'
                Name        = $DateColumnHeader
                Value       = Format-EventDateString $Log.$RawDateProperty
            }
            $CustomObject | Add-Member @AddParams

            # Raw
            # convert audit data to ps format
            $Raw = $Log | ConvertTo-Json -Depth 10
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Raw'
                Value      = $Raw
            }
            $CustomObject | Add-Member @AddParams

            # Tree
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'Tree'
            #     Value      = $Log | Format-Tree -Depth 10 | Out-String # need to figure out how to output to pipeline, not host
            # }
            # $CustomObject | Add-Member @AddParams

            # UserIds # needs to see if this works consistently across multipe log types
            if ( $Log.UserIds -match '^ServicePrincipal_.*$' ) {
                $SpName = $AuditData.Actor[0].ID
                $UserIds = "SP: ${SpName}"
            }
            else {
                $UserIds = $Log.UserIds
            }
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'UserIds'
                Value      = $UserIds
            }
            $CustomObject | Add-Member @AddParams

            # RecordType
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'RecordType'
                Value      = $Log.RecordType
            }
            $CustomObject | Add-Member @AddParams

            #region OPERATION
            # find operation friendly name
            $Row = $OperationsData | Where-Object { $_.Operation -eq $Log.Operations }
            $OperationFriendlyName = if (-not [string]::IsNullOrWhiteSpace($Row.CustomDescription)) {
                $Row.CustomDescription
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Row.FriendlyName)) {
                $Row.FriendlyName
            }
            else {
                $Log.Operations
            }
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'OperationFriendlyName'
                Value      = $OperationFriendlyName
            }
            $CustomObject | Add-Member @AddParams

            #region IP ADDRESSES
            $IpAddresses = [System.Collections.Generic.Hashset[string]]::new()
            if ( $AuditData.ClientIP ) {
                foreach ($Ip in $AuditData.ClientIP) {[void]$IpAddresses.Add($Ip)}
            }
            if ( $AuditData.ActorIpAddress ) {
                foreach ($Ip in $AuditData.ActorIpAddress) {[void]$IpAddresses.Add($Ip)}
            }
            if ( $AuditData.ClientIPAddress ) {
                foreach ($Ip in $AuditData.ClientIPAddress) {[void]$IpAddresses.Add($Ip)}
            }
            $IpString = $IpAddresses -join "`n"
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'IpAddresses'
                Value      = $IpString
            }
            $CustomObject | Add-Member @AddParams

            # process log type
            $RecordType = $Log.RecordType
            $Operations = $Log.Operations
            $OperationString = $RecordType + ' ' + $Operations
            $ResolveParams = @{
                Log = $Log
                AuditData = $AuditData
            }
            switch ( $OperationString ) {
                'AzureActiveDirectory Update user.' {
                    $EventObject = Resolve-AzureActiveDirectoryUpdateUser @ResolveParams
                }
                'ExchangeAdmin Set-ConditionalAccessPolicy' {
                    $EventObject = Resolve-ExchangeAdminSetConditionalAccessPolicy @ResolveParams
                }
                'ExchangeItemAggregated AttachmentAccess' {
                    $EventObject = Resolve-ExchangeItemAggregatedAttachmentAccess @ResolveParams
                }
                'ExchangeItemAggregated MailItemsAccessed' {
                    $EventObject = Resolve-ExchangeItemAggregatedMailItemsAccessed @ResolveParams
                }
                'ExchangeItem Create' {
                    $EventObject = Resolve-ExchangeItemSubject @ResolveParams
                }
                'ExchangeItem Send' {
                    $EventObject = Resolve-ExchangeItemSubject @ResolveParams
                }
                'ExchangeItem Update' {
                    $EventObject = Resolve-ExchangeItemUpdate @ResolveParams
                }
                'ExchangeItemGroup HardDelete' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @ResolveParams
                }
                'ExchangeItemGroup MoveToDeletedItems' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @ResolveParams
                }
                'ExchangeItemGroup SoftDelete' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @ResolveParams
                }
                'SharePoint SearchQueryPerformed' {
                    $EventObject = Resolve-SharePointSearchQueryPerformed @ResolveParams
                }
                'SharePointFileOperation FileAccessed' {
                    $EventObject = Resolve-SharePointFileOperationFileAccessed -Log $Log
                }
                'SharePointFileOperation FileModified' {
                    $EventObject = Resolve-SharePointFileOperationFileAccessed -Log $Log
                }
                'SharePointFileOperation FileModifiedExtended' {
                    $EventObject = Resolve-SharePointFileOperationFileAccessed -Log $Log
                }
                'SharePoint PageViewed' {
                    $EventObject = Resolve-SharePointPageViewed @ResolveParams
                }
                default {
                    $EventObject = [pscustomobject]@{
                        Summary = ''
                    }
                }
            }
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Summary'
                Value      = $EventObject.Summary
            }
            $CustomObject | Add-Member @AddParams

            # add to list
            $OutputTable.Add( $CustomObject )
        }

        # select just relevant properties and set column order
        $OutputTable = $OutputTable | Select-Object $DisplayProperties

        #region EXPORT SPREADSHEET
        # export spreadsheet
        $WorkSheetName = 'UnifiedAuditLogs'
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

        $IpAddressesColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'IpAddresses'}).Id | Convert-DecimalToExcelColumn
        $SummaryColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'Summary'}).Id | Convert-DecimalToExcelColumn

        #region CELL COLORING
        # if cell matches EXACTLY, make background RED
        $Strings = @(
            'Add app role assignment grant to user.'
            'Add app role assignment to service principal.'
            'Add application.'
            'Created new inbox rule in Outlook web app'
            'Consent to application.'
            'Update application – Certificates and secrets management'
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

        # if cell CONTAINS, make background RED
        $Strings = @(
            'New inbox rule'
            'Enable-InboxRule'
            'Modified inbox rule from Outlook web app'
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
        
        # if cell matches EXACTLY, make background YELLOW
        $Strings = @(
            # 'User started security info registration'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
                RuleType        = 'Equal'
                ConditionValue  = $String
                BackgroundColor = 'LightGoldenRodYellow'
            }
            Add-ConditionalFormatting @CFParams
        }
        
        # if cell CONTAINS text anywhere, make background BLUE
        $Strings = @(
            # 'AppOwnerOrganizationId: Microsoft'
        )
        foreach ( $String in $Strings ) {
            $CFParams = @{
                Worksheet       = $WorkSheet
                Address         = "${TableStartColumn}${TableStartRow}:${EndColumn}${EndRow}"
                RuleType        = 'ContainsText'
                ConditionValue  = $String
                BackgroundColor = 'LightBlue'
            }
            Add-ConditionalFormatting @CFParams
        }

        #region COLUMN WIDTH

        # resize Index column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Index' } ).Id 
        $Worksheet.Column($Column).Width = 8

        # resize DateTime column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id 
        $Worksheet.Column($Column).Width = 25

        # resize Raw column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8

        # # resize Tree column
        # $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Tree' } ).Id 
        # $Worksheet.Column($Column).Width = 12

        # resize UserIds column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'UserIds' } ).Id 
        $Worksheet.Column($Column).Width = 30

        # resize IpAddresses column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'IpAddresses' } ).Id 
        $Worksheet.Column($Column).Width = 16

        # resize OperationFriendlyName column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'OperationFriendlyName' } ).Id 
        $Worksheet.Column($Column).Width = 30

        # resize Summary column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Summary' } ).Id 
        $Worksheet.Column($Column).Width = 200

        #region FORMATTING

        # set text wrapping on specific columns
        $WrappingParams = @{
            Worksheet = $Worksheet
            Range     = "${SummaryColumn}${TableStartRow}:${SummaryColumn}${EndRow}"
            WrapText  = $true
        }
        Set-ExcelRange @WrappingParams
        $WrappingParams = @{
            Worksheet = $Worksheet
            Range     = "${IpAddressesColumn}${TableStartRow}:${IpAddressesColumn}${EndRow}"
            WrapText  = $true
        }
        Set-ExcelRange @WrappingParams

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