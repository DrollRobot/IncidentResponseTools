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
        [System.Collections.Generic.List[PSObject]] $Logs,

        [Parameter(Position=0, Mandatory, ParameterSetName='Xml')]
        [string] $XmlPath,

        [string] $TableStyle = 'Dark8',
        [boolean] $Xml = $true,
        [boolean] $WaitOnMessageTrace = $false,
        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        $ModulePath = $PSScriptRoot
        $Function = $MyInvocation.MyCommand.Name
        $ParameterSet = $PSCmdlet.ParameterSetName
        $RawDateProperty = 'CreationDate'
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $TitleDateFormat = "M/d/yy h:mmtt"
        $DateColumnHeader = 'DateTime'

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }

        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($Test) {$Global:IRTTestMode = $true}
        
        # import from xml
        if ($ParameterSet -eq 'Xml') {
            try {
                $ResolvedXmlPath = Resolve-ScriptPath -Path $XmlPath -File -FileExtension 'xml'
                [System.Collections.Generic.List[PSObject]]$Logs = Import-CliXml -Path $ResolvedXmlPath
            }
            catch {
                $_
                Write-Host @Red "${Function}: Error importing from ${XmlPath}."
                return
            }
        }

        # $Groups = Request-GraphGroups
        # $Roles = Request-DirectoryRoles
        # $RoleTemplates = Request-DirectoryRoleTemplates
        # $ServicePrincipals = Request-GraphServicePrincipals
        # $Users = Request-GraphUsers

        # import metadata
        if ($Logs[0].Metadata) {

            # remove metadata from beginning of list
            $Metadata = $Logs[0]
            $Logs.RemoveAt(0)

            $UserEmail = $Metadata.UserEmail
            $UserName = $Metadata.UserName
            $StartDate = $Metadata.StartDate
            $EndDate = $Metadata.EndDate
            $Days = $Metadata.Days
            # $Domain = $Metadata.Domain
            $FileNamePrefix = $Metadata.FileNamePrefix
        }
        else {
            Write-Host @Red "${Function}: No Metadata found."
        }

        # build file name
        $FileDateString = $EndDate.ToLocalTime().ToString($FileNameDateFormat)
        $ExcelOutputPath =  "${FileNamePrefix}_${Days}Days_${UserName}_${FileDateString}.xlsx"

        # build worksheet title
        $TitleEndDate = $EndDate.ToLocalTime().ToString($TitleDateFormat)
        $TitleStartDate = $StartDate.ToLocalTime().ToString($TitleDateFormat)
        $WorksheetTitle = "Unified audit logs for ${UserEmail}. Covers ${Days} days, ${TitleStartDate} to ${TitleEndDate}."

        # import all operations csv
        $AllOperationsFileName = 'unified_audit_log-all_operations.csv'
        $OperationsCsvPath = Join-Path -Path $ModulePath -ChildPath "\unified_audit_log-data\${AllOperationsFileName}"
        $OperationsData = Import-Csv -Path $OperationsCsvPath

        if ($Global:IRTTestMode) {
            # build set of operations from csv
            $OperationsFromCsv = [System.Collections.Generic.Hashset[PSCustomObject]]::new() 
            foreach ($Row in $OperationsData) {
                [void]$OperationsFromCsv.Add(
                    [PSCustomObject]@{
                        Workload = $Row.Workload
                        RecordType = $Row.RecordType
                        Operation = $Row.Operation
                    }
                )
            }
           $OperationsFromLog = [System.Collections.Generic.Hashset[PSCustomObject]]::new() 
        }
    }

    process {

        # convert audit data to powershell objects
        foreach ($Log in $Logs) {
            $Log.AuditData = $Log.AuditData | ConvertFrom-Json -Depth 10
        }

        #region ROW LOOP

        # process each log
        for ($i = 0; $i -lt ($Logs | Measure-Object).Count; $i++) { 
        
            $Log = $Logs[$i]
            $Row = [PSCustomObject]@{}

            # save operations to create complete list
            if ($Global:IRTTestMode) {
                [void]$OperationsFromLog.Add(
                    [PSCustomObject]@{
                        Workload = $Log.AuditData.Workload
                        RecordType = $Log.AuditData.RecordType
                        Operation = $Log.AuditData.Operation
                    }
                )
            }

            # extract auditdata
            $AuditData = $Log.AuditData # FIXME transition to using original log object converted from json

            # Raw
            $Raw = $Log | ConvertTo-Json -Depth 10
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Raw'
                Value      = $Raw
            }
            $Row | Add-Member @AddParams

            # Tree
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'Tree'
            #     Value      = $Log | Format-Tree -Depth 10 | Out-String # need to figure out how to output to pipeline, not host
            # }
            # $Row | Add-Member @AddParams

            # Date/Time
            $AddParams = @{
                MemberType  = 'NoteProperty'
                Name        = $DateColumnHeader
                Value       = Format-EventDateString $Log.$RawDateProperty
            }
            $Row | Add-Member @AddParams

            #region USERIDS
            # need to see if this works consistently across multiple log types
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
            $Row | Add-Member @AddParams

            # Workload
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Workload'
                Value      = $Log.AuditData.Workload
            }
            $Row | Add-Member @AddParams

            # RecordType
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'RecordType'
                Value      = $Log.RecordType
            }
            $Row | Add-Member @AddParams

            #region OPERATION NAME
            # find operation friendly name
            $OperationsRow = $OperationsData | Where-Object { $_.Operation -eq $Log.Operations }
            $OperationFriendlyName = if (-not [string]::IsNullOrWhiteSpace($OperationsRow.CustomDescription)) {
                $OperationsRow.CustomDescription
            }
            elseif (-not [string]::IsNullOrWhiteSpace($OperationsRow.FriendlyName)) {
                $OperationsRow.FriendlyName
            }
            else {
                $Log.Operations
            }
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'OperationFriendlyName'
                Value      = $OperationFriendlyName
            }
            $Row | Add-Member @AddParams

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
            $Row | Add-Member @AddParams

            # process log type
            $RecordType = $Log.RecordType
            $Operations = $Log.Operations
            $OperationString = $RecordType + ' ' + $Operations

            $OldParams = @{ # FIXME Update so separate auditdata isn't required.
                Log = $Log
                AuditData = $AuditData
            }
            $EmailParams = @{
                Log = $Log
                WaitOnMessageTrace = $WaitOnMessageTrace
                UserName = $UserName
            }
            switch ( $OperationString ) {
                'AzureActiveDirectory Update user.' {
                    $EventObject = Resolve-AzureActiveDirectoryUpdateUser @OldParams
                }
                'ExchangeAdmin Set-ConditionalAccessPolicy' {
                    $EventObject = Resolve-ExchangeAdminSetConditionalAccessPolicy @OldParams
                }
                'ExchangeItemAggregated AttachmentAccess' {
                    $EventObject = Resolve-ExchangeItemAggregatedAttachmentAccess @OldParams
                }
                'ExchangeItemAggregated MailItemsAccessed' {
                    $EventObject = Resolve-ExchangeItemAggregatedMailItemsAccessed @EmailParams
                }
                'ExchangeItem Create' {
                    $EventObject = Resolve-ExchangeItemSubject @OldParams
                }
                'ExchangeItem Send' {
                    $EventObject = Resolve-ExchangeItemSubject @OldParams
                }
                'ExchangeItem Update' {
                    $EventObject = Resolve-ExchangeItemUpdate @OldParams
                }
                'ExchangeItemGroup HardDelete' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @EmailParams
                }
                'ExchangeItemGroup MoveToDeletedItems' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @EmailParams
                }
                'ExchangeItemGroup SoftDelete' {
                    $EventObject = Resolve-ExchangeItemGroupDelete @EmailParams
                }
                'SharePoint SearchQueryPerformed' {
                    $EventObject = Resolve-SharePointSearchQueryPerformed @OldParams
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
                'SharePointFileOperation FilePreviewed' {
                    $EventObject = Resolve-SharePointFileOperationFileAccessed -Log $Log
                }
                'SharePointFileOperation FileUploaded' {
                    $EventObject = Resolve-SharePointFileOperationFileAccessed -Log $Log
                }
                'SharePoint PageViewed' {
                    $EventObject = Resolve-SharePointPageViewed @OldParams
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
            $Row | Add-Member @AddParams

            # add to list
            [void]$Rows.Add($Row)
        }

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
            $Workbook = $Rows | Export-Excel @ExcelParams
        }
        catch {
            Write-Error "${Function}: Unable to open new Excel document."
            if ( Get-YesNo "Try closing open files." ) {
                try {
                    $Workbook = $Rows | Export-Excel @ExcelParams
                }
                catch {
                    throw "${Function}: Unable to open new Excel document. Exiting."
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
            'Enable-InboxRule'
            'New-InboundConnector'
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

        # resize Raw column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Raw' } ).Id 
        $Worksheet.Column($Column).Width = 8

        # # resize Tree column
        # $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'Tree' } ).Id 
        # $Worksheet.Column($Column).Width = 12

        # resize DateTime column
        $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq $DateColumnHeader } ).Id 
        $Worksheet.Column($Column).Width = 26

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

        if ($Global:IRTTestMode) {
            $OperationsToAdd = [System.Collections.Generic.HashSet[PSCustomObject]]::new() 
            # find operations from log that are missing in csv
            foreach ($o in $OperationsFromLog) {
                if ($OperationsFromCsv.Add($o)) {
                    [void]$OperationsToAdd.Add($o)
                }
            }
            Write-Host @Red "Add to ${AllOperationsFileName}:"
            $OperationsToAdd | Format-Table | Out-Host
            Write-Host @Red "Exporting to: operations_to_add.csv"
            $OperationsToAdd | Export-Csv -Path "operations_to_add.csv" -NoTypeInformation
        }
                    
        # save and close
        Write-Host @Blue "${Function}: Exporting to: ${ExcelOutputPath}"
        $Workbook | Close-ExcelPackage -Show
    }
}