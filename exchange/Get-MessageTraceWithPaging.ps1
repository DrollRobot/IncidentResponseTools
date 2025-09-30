function Get-MessageTraceWithPaging {
    param(
        [string] $SenderAddress,
        [string] $RecipientAddress,

        [Parameter( Mandatory )]
        [datetime] $StartDate,

        [Parameter( Mandatory )]
        [datetime] $EndDate,

        [int] $PageLimit = 10
    )

    # variables
    $Blue = @{ ForegroundColor = 'Blue' }
    $AllMessages = [System.Collections.Generic.List[psobject]]::new()
    $PageSize = 5000 # 5000 is max page size for message trace
    $Page       = 1
    $HasData    = $true

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

    # get all records
    while ( $HasData -and $Page -le $PageLimit ) {

        $Params['Page'] = $Page

        # retrieve one page
        Write-Host @Blue "Requesting message trace page ${Page}"
        $PageResults = [psobject[]]@( Get-MessageTrace @Params )

        if ( $PageResults ) {
            $AllMessages.AddRange( $PageResults )
        } 

        # stop when the page wasn't completely filled
        if ( $PageResults.Count -lt $PageSize ) {
            $HasData = $false
        } else {
            $Page++
        }
    }

    return $AllMessages
}