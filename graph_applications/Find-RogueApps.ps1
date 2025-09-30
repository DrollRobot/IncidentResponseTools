New-Alias -Name 'RogueApps' -Value 'Find-RogueApps' -Force
function Find-RogueApps {

    begin {

        # variables
        $ESC = [char]27
        $End = "${ESC}[0m"
        $BrightGreen = "${ESC}[92m"
        $UserDisplayProperties = @(
            'AccountEnabled'
            'DisplayName'
            'UserPrincipalName'
            'Id' 
        )
        $HuntressProperties = @(
            'AppDisplayName'
            'Description'
            'Tags'
            'References'
        )
        $SyneProperties = @(
            'Name'
            'Description'
            'Categories'
            'References'
        )
        $FoundApps = $false

        # get all service principals
        $ServicePrincipals = Request-GraphServicePrincipals
	
        # get permission grants
        $PermissionGrants = Request-GraphOauth2Grants
	
        # get all users
        $Users = Request-GraphUsers
    }

    process {

        ### show settings
        Write-Host -ForegroundColor Blue "`nTenant App settings:"
        # user app registration/creation
        Write-Host "`n'AllowedToCreateApps' indicates whether users are allowed to create their own applications from scratch."
	    ( Get-MgPolicyAuthorizationPolicy ).DefaultUserRolePermissions | 
            Format-List AllowedToCreateApps | 
            Out-Host

        # user app consent
        Write-Host "If 'ManagePermissionGrantsForSelf.microsoft-user-default' is present, this indicates users are allowed to consent for 3rd party apps."
        $UserConsent = (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions | 
            Select-Object -ExpandProperty PermissionGrantPoliciesAssigned | 
            Where-Object { $_ -match "ManagePermissionGrantsForSelf.microsoft-user-default" }
        Write-Host "${BrightGreen}UserConsent:${End} ${UserConsent}"


        # get huntress data
        $HuntressUrl = "https://raw.githubusercontent.com/huntresslabs/rogueapps/main/public/rogueapps.json"
        $HuntressApps = Invoke-RestMethod -Uri $HuntressUrl

        # get syne data
        $SyneUrl = "https://raw.githubusercontent.com/randomaccess3/detections/refs/heads/main/M365_Oauth_Apps/MaliciousOauthAppDetections.json"
        $SyneApps = Invoke-RestMethod -Uri $SyneUrl

        # build combined list
        $SusAppIds = $HuntressApps.AppId + $SyneApps.Applications.AppId | Sort-Object -Unique
	
        # find risky apps
        $RiskyApps = $ServicePrincipals | Where-Object { $_.AppId -in $SusAppIds }
	
        foreach ( $RiskyApp in $RiskyApps ) {

            $FoundApps = $true
	
            # find permission grants for the app
            $AppGrants = $PermissionGrants | Where-Object { $_.ClientId -eq $RiskyApp.Id }
	
            # show app information
            Write-Host -ForegroundColor Blue "`nApp Information:"
            $HuntressInfo = $HuntressApps | Where-Object { $_.AppId -eq $RiskyApp.AppId }
            $SyneInfo = $SyneApps.Applications | Where-Object { $_.AppId -eq $RiskyApp.AppId }
            if ( $HuntressInfo ) {
                $HuntressInfo | Format-List $HuntressProperties | Out-Host
            }
            elseif ( $SyneInfo ) {
                $SyneInfo | Format-List $SyneProperties | Out-Host
            }
	
            # show users who have the app
            Write-Host -ForegroundColor Blue "Users who have this app:"
            $Users | 
                Where-Object { $_.Id -in $AppGrants.PrincipalId } | 
                Format-Table $UserDisplayProperties |
                Out-Host
        }

        if ( $FoundApps -eq $false ) {

            Write-Host "`nNo rogue apps found."
            Write-Host ""
        }
    }
}