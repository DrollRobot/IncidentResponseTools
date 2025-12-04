function Format-EventDateString {
    <#

    # FIXME Delete this file once converted over to using native excel date formatting.

    .SYNOPSIS
    Builds a date/time string that will be easily readable in a spreadsheet.
    #>
    [CmdletBinding()]
    param(
        [Parameter( Position = 0 , Mandatory )]
        [datetime] $DateTime,
        [string] $Format = 'MM/dd/yy hh:mm:sstt',
        [System.TimeZoneInfo] $TimeZoneInfo = [System.TimeZoneInfo]::Local
    )

    # build date string
    $BuildString = $DateTime.ToLocalTime().ToString($Format).ToLower()

    # create acronym from timezone full name
    if ($DateTime.IsDaylightSavingTime()) {
        $TimeZoneName = $TimeZoneInfo.DaylightName
    }
    else {
        $TimeZoneName = $TimeZoneInfo.StandardName
    }
    $TimeZoneAcronym = -join ($TimeZoneName -split ' ' | ForEach-Object { $_[0] })

    # add time zone acronym to string
    $EventDateString = $BuildString + ' ' + $TimeZoneAcronym

    # replace leading zeros with spaces for alignment
    if ($EventDateString[0] -eq '0') {
        $EventDateString = ' ' + $EventDateString.Substring(1)
    }
    if ($EventDateString[9] -eq '0') {
        $EventDateString = $EventDateString.Substring(0,9) + ' ' + $EventDateString.Substring(10)
    }

    return $EventDateString
}
