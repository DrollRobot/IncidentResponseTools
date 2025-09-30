function Get-FullUserObject {
    <#
    .SYNOPSIS
    retrieves a user with a broad set of properties and augments with optional ones.

    .NOTES
    version: 1.0.5
    - add pipeline support (by object or by id/upn)
    - keep signInActivity in initial selection
    #>
    [CmdletBinding(DefaultParameterSetName='ByObject')]
    param(
        # pipe full user objects (e.g., from get-mguser)
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='ByObject')]
        [ValidateNotNull()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser] $UserObject,

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName='ById')]
        [Alias('Id')]
        [ValidateNotNullOrEmpty()]
        [string] $UserId,

        [Parameter(ParameterSetName='ByObject')]
        [Parameter(ParameterSetName='ById')]
        [switch] $NoRefresh
    )

    begin {
        # properties you can $select directly on /users
        $SelectProps = @(
            'id','userPrincipalName','displayName','accountEnabled',
            'ageGroup','businessPhones','city','companyName','consentProvidedForMinor',
            'country','createdDateTime','creationType','department',
            'employeeHireDate','employeeId','employeeLeaveDateTime','employeeOrgData',
            'employeeType','externalUserState','externalUserStateChangeDateTime',
            'faxNumber','givenName','identities','imAddresses','isResourceAccount',
            'jobTitle','lastPasswordChangeDateTime','legalAgeGroupClassification',
            'licenseAssignmentStates','mail','mailNickname','mobilePhone','officeLocation',
            'onPremisesDistinguishedName','onPremisesDomainName','onPremisesExtensionAttributes',
            'onPremisesImmutableId','onPremisesLastSyncDateTime','onPremisesProvisioningErrors',
            'onPremisesSamAccountName','onPremisesSecurityIdentifier','onPremisesSyncEnabled',
            'onPremisesUserPrincipalName','otherMails','passwordPolicies','passwordProfile',
            'postalCode','preferredDataLocation','preferredLanguage','provisionedPlans',
            'proxyAddresses','securityIdentifier','showInAddressList',
            'signInSessionsValidFromDateTime','state','streetAddress','surname',
            'usageLocation','userType','signInActivity'
        )

        # optional properties that can error or be null depending on licensing/mailbox/etc.
        $OptionalProps = @(
            'aboutMe','birthday','deviceEnrollmentLimit','hireDate','interests',
            'mailboxSettings','mailFolders','mySite','pastProjects','preferredName',
            'print','responsibilities','schools','skills'
        )
    }

    process {

        # if object is already full object, and -NoRefresh, don't query.
        if ($NoRefresh -and $PSCmdlet.ParameterSetName -eq 'ByObject' -and
            $UserObject.PSObject.Properties['AllProperties'] -and $UserObject.AllProperties) {
            Write-Output $UserObject
            return
        }

        # resolve the identifier for this pipeline item
        switch ($PSCmdlet.ParameterSetName) {
            'ById'     { $ResolvedId = $UserId }
            'ByObject' { $ResolvedId = $UserObject.Id }
            default    { $ResolvedId = $null }
        }

        if (-not $ResolvedId) {
            Write-Verbose "skipping item: could not resolve an id."
            return
        }

        # get base user with wide $select
        $GetParams = @{
            UserId      = $ResolvedId
            Property    = $SelectProps
            ErrorAction = 'Stop'
        }

        try {
            $ScriptUserObject = Get-MgUser @GetParams
        }
        catch {
            Write-Error "Get-MgUser failed for '$ResolvedId': $($_.Exception.Message)"
            if ($PSCmdlet.ParameterSetName -eq 'ByObject' -and $UserObject) {
                Write-Output $UserObject
            }
            return
        }

        # augment with optional properties (best-effort)
        foreach ($Property in $OptionalProps) {
            try {
                $OptionalParams = @{
                    UserId      = $ResolvedId
                    Property    = $Property
                    ErrorAction = 'Stop'
                }
                $TempUserObject = Get-MgUser @OptionalParams
                $ScriptUserObject.$Property = $TempUserObject.$Property
            }
            catch {
                Write-Verbose "Unable to retrieve property '$Property' for '$ResolvedId': $($_.Exception.Message)"
            }
        }

        # add property indicating object has all properties.
        $ScriptUserObject | Add-Member -NotePropertyName 'AllProperties' -NotePropertyValue $true -Force

        Write-Output $ScriptUserObject
    }
}