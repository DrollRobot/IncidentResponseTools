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
        [boolean] $IpInfo = $true,
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
        $RawDateProperty = 'CreationDate'
        $DateColumnHeader = 'DateTime'
        $Rows = [System.Collections.Generic.List[PSCustomObject]]::new()

        # colors
        $Blue = @{ ForegroundColor = 'Blue' }
        # $Green = @{ ForegroundColor = 'Green' }
        $Red = @{ ForegroundColor = 'Red' }
        # $Magenta = @{ ForegroundColor = 'Magenta' }
        $Yellow = @{ ForegroundColor = 'Yellow' }
        
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

        #region METADATA

        if ($Logs[0].Metadata) {

            # remove metadata from beginning of list
            $Metadata = $Logs[0]
            $Logs.RemoveAt(0)

            # $UserEmail = $Metadata.UserEmail
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
        $FileNameDateFormat = "yy-MM-dd_HH-mm"
        $FileDateString = $EndDate.ToLocalTime().ToString($FileNameDateFormat)
        $ExcelOutputPath =  "${FileNamePrefix}_${Days}Days_${UserName}_${FileDateString}.xlsx"

        # build worksheet title
        $TitleDateFormat = "M/d/yy h:mmtt"
        $TitleEndDate = $EndDate.ToLocalTime().ToString($TitleDateFormat)
        $TitleStartDate = $StartDate.ToLocalTime().ToString($TitleDateFormat)
        $WorksheetTitle = "Unified audit logs for ${Username}. Covers ${Days} days, ${TitleStartDate} to ${TitleEndDate}."

        # import alloperations csv
        $ModulePath = $PSScriptRoot
        $AllOperationsFileName = 'unified_audit_log-all_operations.csv'
        $OperationsCsvPath = Join-Path -Path $ModulePath -ChildPath "\unified_audit_log-data\${AllOperationsFileName}"
        $OperationsCsvData = Import-Csv -Path $OperationsCsvPath

        # test
        if ($Script:Test) {
            # build set of operations from csv to later output missing operations
            $OperationsFromCsv = [System.Collections.Generic.Hashset[string]]::new() 
            foreach ($Row in $OperationsCsvData) {
                [void]$OperationsFromCsv.Add("$($Row.Workload)|$($Row.RecordType)|$($Row.Operation)")
            }
           $OperationsFromLog = [System.Collections.Generic.Hashset[string]]::new()

            # start stopwatch
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
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
            # convert audit data to powershell objects
            $Log.AuditData = $Log.AuditData | ConvertFrom-Json -Depth 10

            # collect ip addresses
            if ($IpInfo) {
                if ( $Log.AuditData.ClientIP ) {
                    foreach ($Ip in $Log.AuditData.ClientIP) {[void]$IpInfoAddresses.Add($Ip)}
                }
                if ( $Log.AuditData.ActorIpAddress ) {
                    foreach ($Ip in $Log.AuditData.ActorIpAddress) {[void]$IpInfoAddresses.Add($Ip)}
                }
                if ( $Log.AuditData.ClientIPAddress ) {
                    foreach ($Ip in $Log.AuditData.ClientIPAddress) {[void]$IpInfoAddresses.Add($Ip)}
                } 
            }
        }

        #region QUERY IPS
        if ($IpInfo -and 
            $IpInfoPackage.Present -and
            ($IpInfoAddresses | Measure-Object).Count -gt 0
        ) {

            # query information for all IP addresses
            $IpQueryStart = $Stopwatch.Elapsed
            & $IpInfoPackage.Python ip_info --apis bulk --output_format none -ip_addresses $IpInfoAddresses
            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $IpQueryStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: ip_info query took ${ElapsedString}" | Out-Host
            }

            # add ip info to global colection
            $IpTableStart = $Stopwatch.Elapsed
            if (($Global:IRT_IpInfo | Measure-Object).Count -eq 0) {
                $Global:IRT_IpInfo = @{}
            }
            foreach ($Ip in $IpInfoAddresses) {
                # convert to object
                $Ip = [System.Net.IPAddress]$Ip
                # if ip doesn't exist in table, add it.
                if (-not $Global:IRT_IpInfo[$Ip]) {
                    $Output = & $IpInfoPackage.Python ip_info $Ip.ToString
                    $Global:IRT_IpInfo[$Ip] = $Output
                }
            }
            if ($Script:Test) {
                $ElapsedString = ($StopWatch.Elapsed - $IpTableStart).ToString('mm\:ss')
                Write-Host @Yellow "${Function}: adding ip info to global table took ${ElapsedString}" | Out-Host
            }
        }

        #region SECOND LOOP

        for ($i = 0; $i -lt ($Logs | Measure-Object).Count; $i++) { 
        
            $Log = $Logs[$i]
            $Row = [PSCustomObject]@{}

            # save operations to create complete list
            if ($Script:Test) {
                [void]$OperationsFromLog.Add(
                    "$($Log.AuditData.Workload)|$($Log.AuditData.RecordType)|$($Log.AuditData.Operation)"
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

            # Operation
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'Operation'
                Value      = $Log.AuditData.Operation
            }
            $Row | Add-Member @AddParams

            # # find operation friendly name # FIXME friendly names not useful?
            # $OperationsRow = $OperationsCsvData | Where-Object { $_.Operation -eq $Log.Operations }
            # $OperationFriendlyName = if (-not [string]::IsNullOrWhiteSpace($OperationsRow.CustomDescription)) {
            #     $OperationsRow.CustomDescription
            # }
            # elseif (-not [string]::IsNullOrWhiteSpace($OperationsRow.FriendlyName)) {
            #     $OperationsRow.FriendlyName
            # }
            # else {
            #     $Log.Operations
            # }
            # $AddParams = @{
            #     MemberType = 'NoteProperty'
            #     Name       = 'OperationFriendlyName'
            #     Value      = $OperationFriendlyName
            # }
            # $Row | Add-Member @AddParams

            #region IP ADDRESSES

            # collect ips
            $IpAddresses = [System.Collections.Generic.Hashset[System.Net.IPAddress]]::new()
            if ( $Log.AuditData.ClientIP ) {
                foreach ($i in $Log.AuditData.ClientIP) {[void]$IpAddresses.Add($i)}
            }
            if ( $Log.AuditData.ActorIpAddress ) {
                foreach ($i in $Log.AuditData.ActorIpAddress) {[void]$IpAddresses.Add($i)}
            }
            if ( $Log.AuditData.ClientIPAddress ) {
                foreach ($i in $Log.AuditData.ClientIPAddress) {[void]$IpAddresses.Add($i)}
            }
            # loop through rows, replace with ip info
            if (($IpAddresses | Measure-Object).Count -gt 0) {

                # build cell text
                $Strings = [System.Collections.Generic.List[string]]::new()
                $Strings.Add(($IpAddresses.ToString() | Sort-Object) -join ', ')

                # add info from table
                foreach ($Ip in $IpAddresses) {
                    $Strings.Add($Global:IRT_IpInfo[[System.Net.IPAddress]$Ip]) 
                }
            }
            $Cell = $Strings -join "`n`n"
            $AddParams = @{
                MemberType = 'NoteProperty'
                Name       = 'IpAddresses'
                Value      = $Cell
            }
            $Row | Add-Member @AddParams

            # add ip addresses to hashset for ip_info
            if ($IpInfo) {
                foreach ($i in $IpInfoAddresses) {[void]$IpInfoAddresses.Add($i)}
            }

            #region SUMMARY
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
                'ExchangeAdmin New-InboxRule' {
                    $EventObject = Resolve-ExchangeAdminInboxRule -Log $Log
                }
                'ExchangeAdmin Set-ConditionalAccessPolicy' {
                    $EventObject = Resolve-ExchangeAdminSetConditionalAccessPolicy -Log $Log
                }
                'ExchangeAdmin Set-InboxRule' {
                    $EventObject = Resolve-ExchangeAdminInboxRule -Log $Log
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
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileDownloaded' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileModified' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileModifiedExtended' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FilePreviewed' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileSyncDownloadedFull' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileSyncUploadedFull' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
                }
                'SharePointFileOperation FileUploaded' {
                    $EventObject = Resolve-SharePointFileOperation -Log $Log
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
        $OperationColumn = ($Worksheet.Tables[0].Columns | Where-Object {$_.Name -eq 'Operation'}).Id | Convert-DecimalToExcelColumn

        #region CELL COLORING

        foreach ($Row in $OperationsCsvData) {

            # color high risk operations red
            if ($Row.Risk -eq 'High') {
                $CFParams = @{
                    Worksheet       = $WorkSheet
                    Address         = "${OperationColumn}${TableStartRow}:${OperationColumn}${EndRow}"
                    RuleType        = 'ContainsText'
                    ConditionValue  = $Row.Operation
                    BackgroundColor = 'LightPink'
                }
                Add-ConditionalFormatting @CFParams
            }

            # color medium risk operations orange
            if ($Row.Risk -eq 'Medium') {
                $CFParams = @{
                    Worksheet       = $WorkSheet
                    Address         = "${OperationColumn}${TableStartRow}:${OperationColumn}${EndRow}"
                    RuleType        = 'ContainsText'
                    ConditionValue  = $Row.Operation
                    BackgroundColor = 'LightGoldenrodYellow'
                }
                Add-ConditionalFormatting @CFParams
            }
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
        $Worksheet.Column($Column).Width = 20

        # # resize OperationFriendlyName column # FIXME
        # $Column = ( $Worksheet.Tables[0].Columns | Where-Object { $_.Name -eq 'OperationFriendlyName' } ).Id 
        # $Worksheet.Column($Column).Width = 30

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

        if ($Script:Test) {
            $OperationsToAdd = [System.Collections.Generic.HashSet[PSCustomObject]]::new() 
            # find operations from log that are missing in csv
            foreach ($o in $OperationsFromLog) {
                if ($OperationsFromCsv.Add($o)) {
                    $Split = $o.Split('|')
                    [void]$OperationsToAdd.Add(
                        [PSCustomObject]@{
                            Workload = $Split[0]
                            RecordType = $Split[1]
                            Operation = $Split[2]
                        }
                    )
                }
            }
            # output for user
            if (($OperationsToAdd | Measure-Object).Count -gt 0) {
                Write-Host @Yellow "${Function}: Add to ${AllOperationsFileName}:" | Out-Host
                $OperationsToAdd | Format-Table | Out-Host
                Write-Host @Yellow "${Function}: Exporting to: operations_to_add.csv" | Out-Host
                $OperationsToAdd | Export-Csv -Path "operations_to_add.csv" -NoTypeInformation
            }
        }
                    
        # save and close
        Write-Host @Blue "${Function}: Exporting to: ${ExcelOutputPath}"
        $Workbook | Close-ExcelPackage -Show
    }
}