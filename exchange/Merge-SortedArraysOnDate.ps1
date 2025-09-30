function Merge-SortedListsOnDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[psobject]] $ListOne,

        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[psobject]] $ListTwo,

        [string] $PropertyName,

        [Parameter(Mandatory, ParameterSetName = 'Ascending')]
        [switch] $Ascending,

        [Parameter(Mandatory, ParameterSetName = 'Descending')]
        [switch] $Descending
    )

    # initialize indexes and merged list
    $IndexOne  = 0
    $IndexTwo  = 0
    $MergedList = [System.Collections.Generic.List[psobject]]::new()

    # if descending
    if ($Descending) {
        while ($IndexOne -lt $ListOne.Count -and $IndexTwo -lt $ListTwo.Count) {
            if ($ListOne[$IndexOne].$PropertyName -gt $ListTwo[$IndexTwo].$PropertyName) {
                $MergedList.Add($ListOne[$IndexOne])
                $IndexOne++
            } else {
                $MergedList.Add($ListTwo[$IndexTwo])
                $IndexTwo++
            }
        }
    }
    # otherwise, ascending
    else {
        while ($IndexOne -lt $ListOne.Count -and $IndexTwo -lt $ListTwo.Count) {
            if ($ListOne[$IndexOne].$PropertyName -le $ListTwo[$IndexTwo].$PropertyName) {
                $MergedList.Add($ListOne[$IndexOne])
                $IndexOne++
            } else {
                $MergedList.Add($ListTwo[$IndexTwo])
                $IndexTwo++
            }
        }
    }

    # add any remaining elements
    while ($IndexOne -lt $ListOne.Count) {
        $MergedList.Add($ListOne[$IndexOne])
        $IndexOne++
    }
    while ($IndexTwo -lt $ListTwo.Count) {
        $MergedList.Add($ListTwo[$IndexTwo])
        $IndexTwo++
    }

    return $MergedList
}