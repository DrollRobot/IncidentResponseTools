New-Alias -Name 'FindObject' -Value 'Find-GraphDirectoryObjects' -Force
New-Alias -Name 'FindObjects' -Value 'Find-GraphDirectoryObjects' -Force
New-Alias -Name 'Find-GraphDirectoryObject' -Value 'Find-GraphDirectoryObjects' -Force
function Find-GraphDirectoryObjects {
    param(
        [Parameter( Position = 0 )]
        [string] $Content,

        [string] $TenantId
    )

    begin {
        
        # variables
        $GuidPattern = "\b[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\b"
        $Cyan = @{
            ForegroundColor = 'Cyan'
        }
        $FgGreen = @{
            ForegroundColor = 'Green'
        }
        # $FgRed = @{
        #     ForegroundColor = 'Red'
        # }

        # get content from clipboard
        if ( -not $Content ) {
            Write-Host @FgGreen "`nNo content provided. Pulling from clipboard."
            $Content = Get-Clipboard
            if ( @( $Content ).Count -eq 0 ) {
                throw "No content provided, or found in clipboard. Exiting."
            }
            $DisplayLines = $Content -split "`r`n" | Select-Object -First 3
            $TruncatedLines = $DisplayLines | ForEach-Object {
                if ( $_.Length -gt 80 ) {
                    $_.Substring(0, 77) + "..." 
                }
                else { 
                    $_
                }
            }
            Write-Host $TruncatedLines
        }        
    }

    process {

        $Guids = $Content | Select-String -Pattern $GuidPattern -AllMatches | ForEach-Object { $_.Matches.Value }

        # remove duplicates
        $Guids = $Guids | Sort-Object -Unique
        
        Write-Host @Cyan "`nFound GUIDS:"
        $Guids
        
        foreach ( $Guid in $Guids ) {

            # variables
            $DirectoryObject = $null
            $ObjectType = $null
        
            Write-Host @Cyan "`nRunning Get-MgDirectoryObject for ${Guid}"

            try {

                $DirectoryObject = Get-MgDirectoryObject -DirectoryObjectId $Guid -ErrorAction Stop
        
                $ObjectType = $DirectoryObject.AdditionalProperties.'@odata.type' -replace '#', ''
                
                Write-Host "`nObjectType: ${ObjectType}"
            }
            catch {
                if ( $_ -match 'does not exist or one of its queried reference-property objects are not present' ) {
                    Write-Host "Unable to find object."
                }
                else {
                    $_
                }
            }

            switch ( $ObjectType ) {
                'microsoft.graph.user' {
                    $Object = Get-MgUser -UserId $Guid
                    $Object | Format-Table
                }
                'microsoft.graph.group' {
                    $Object = Get-MgGroup -GroupId $Guid
                    $Object | Format-Table
                }
                # microsoft.graph.serviceprincipal?
            }
        }
    }
}