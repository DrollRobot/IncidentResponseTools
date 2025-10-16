function Request-IRTMessageTraceV1 {
    param(
        [string] $SenderAddress,
        [string] $RecipientAddress,

        [Parameter( Mandatory )]
        [datetime] $StartDate,

        [Parameter( Mandatory )]
        [datetime] $EndDate,

        [int] $ResultLimit = 50000,

        [switch] $Test
    )

    begin {

        #region BEGIN

        # constants
        # $Function = $MyInvocation.MyCommand.Name
        # $ParameterSet = $PSCmdlet.ParameterSetName
        if ($Test) {
            $Script:Test = $true
            # start stopwatch
            # $Stoconnwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }


        # colors
        $Blue = @{ForegroundColor = 'Blue'}
        # $Green = @{ForegroundColor = 'Green'}
        # $Magenta = @{ForegroundColor = 'Magenta'}
        # $Red = @{ForegroundColor = 'Red'}
        # $Yellow = @{ForegroundColor = 'Yellow'}

        $PageSize = 5000 # 5000 is max page size for message trace
        $Page       = 1
        $MoreToGet    = $true

        $Params = @{
            StartDate = $StartDate
            EndDate   = $EndDate
            PageSize   = $PageSize
        }
        if ( $SenderAddress ) {
            $Params['SenderAddress'] = $SenderAddress
        }
        if ( $RecipientAddress ) {
            $Params['RecipientAddress'] = $RecipientAddress
        }
    }

    process {

        # get all records
        $AllMessages = [System.Collections.Generic.List[psobject]]::new()
        while ($MoreToGet -and $AllMessages.Count -le $ResultLimit ) {

            $Params['Page'] = $Page

            # retrieve one page
            Write-Host @Blue "Requesting message trace page ${Page}"
            $PageResults = [psobject[]]@( Get-MessageTrace @Params )
            foreach ($i in $PageResults) {[void]$AllMessages.Add($i)}

            # stop if the page had less than max page size
            if (($PageResults | Measure-Object).Count -lt $PageSize ) {
                $MoreToGet = $false
            } else {
                $Page++
            }
        }

        return $AllMessages
    }
}