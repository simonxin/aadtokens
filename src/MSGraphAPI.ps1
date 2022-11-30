# This script contains functions for MSGraph API at https://graph.microsoft.com

# Returns the 50 latest signin entries or the given entry
# Jun 9th 2020
function Get-AzureSignInLog
{
    <#
    .SYNOPSIS
    Returns the 50 latest entries from Azure AD sign-in log or single entry by id

    .DESCRIPTION
    Returns the 50 latest entries from Azure AD sign-in log or single entry by id

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Get-AADIntAzureSignInLog

    createdDateTime              id                                   ipAddress      userPrincipalName             appDisplayName                   
    ---------------              --                                   ---------      -----------------             --------------                   
    2020-05-25T05:54:28.5131075Z b223590e-8ba1-4d54-be54-03071659f900 199.11.103.31  admin@company.onmicrosoft.com Azure Portal                     
    2020-05-29T07:56:50.2565658Z f6151a97-98cc-444e-a79f-a80b54490b00 139.93.35.110  user@company.com              Azure Portal                     
    2020-05-29T08:02:24.8788565Z ad2cfeff-52f2-442a-b8fc-1e951b480b00 11.146.246.254 user2@company.com             Microsoft Docs                   
    2020-05-29T08:56:48.7857468Z e0f8e629-863f-43f5-a956-a4046a100d00 1.239.249.24   admin@company.onmicrosoft.com Azure Active Directory PowerShell

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Get-AADIntAzureSignInLog

    createdDateTime              id                                   ipAddress      userPrincipalName             appDisplayName                   
    ---------------              --                                   ---------      -----------------             --------------                   
    2020-05-25T05:54:28.5131075Z b223590e-8ba1-4d54-be54-03071659f900 199.11.103.31  admin@company.onmicrosoft.com Azure Portal                     
    2020-05-29T07:56:50.2565658Z f6151a97-98cc-444e-a79f-a80b54490b00 139.93.35.110  user@company.com              Azure Portal                     
    2020-05-29T08:02:24.8788565Z ad2cfeff-52f2-442a-b8fc-1e951b480b00 11.146.246.254 user2@company.com             Microsoft Docs                   
    2020-05-29T08:56:48.7857468Z e0f8e629-863f-43f5-a956-a4046a100d00 1.239.249.24   admin@company.onmicrosoft.com Azure Active Directory PowerShell

    PS C:\>Get-AADIntAzureSignInLog -EntryId b223590e-8ba1-4d54-be54-03071659f900

    id                 : b223590e-8ba1-4d54-be54-03071659f900
    createdDateTime    : 2020-05-25T05:54:28.5131075Z
    userDisplayName    : admin company
    userPrincipalName  : admin@company.onmicrosoft.com
    userId             : 289fcdf8-af4e-40eb-a363-0430bc98d4d1
    appId              : c44b4083-3bb0-49c1-b47d-974e53cbdf3c
    appDisplayName     : Azure Portal
    ipAddress          : 199.11.103.31
    clientAppUsed      : Browser
    userAgent          : Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36
    ...
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$EntryId,
        [switch]$Export
        
    )
    Process
    {

        # Select one entry if provided
        if($EntryId)
        {
            $queryString = "`$filter=id eq '$EntryId'"
        }
        else
        {
            $queryString = "`$top=50&`$orderby=createdDateTime"
        }

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "auditLogs/signIns" -QueryString $queryString

        # Return full results
        if($Export)
        {
            return $results
        }
        elseif($EntryId) # The single entry
        {
            return $results
        }
        else # Print out only some info - the API always returns all info as $Select is not supported :(
        {
            $results | select createdDateTime,id,ipAddress,userPrincipalName,appDisplayName | ft
        }
    }
}

# Returns the 50 latest signin entries or the given entry
function Get-AzureAuditLog
{
    <#
    .SYNOPSIS
    Returns the 50 latest entries from Azure AD sign-in log or single entry by id

    .DESCRIPTION
    Returns the 50 latest entries from Azure AD sign-in log or single entry by id

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Get-AADIntAzureAuditLog

    id                                                            activityDateTime             activityDisplayName   operationType result  initiatedBy   
    --                                                            ----------------             -------------------   ------------- ------  -----------   
    Directory_9af6aff3-dc09-4ac1-a1d3-143e80977b3e_EZPWC_41985545 2020-05-29T07:57:51.4037921Z Add service principal Add           success @{user=; app=}
    Directory_f830a9d4-e746-48dc-944c-eb093364c011_1ZJAE_22273050 2020-05-29T07:57:51.6245497Z Add service principal Add           failure @{user=; app=}
    Directory_a813bc02-5d7a-4a40-9d37-7d4081d42b42_RKRRS_12877155 2020-06-02T12:49:38.5177891Z Add user              Add           success @{app=; user=}

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Get-AADIntAzureAuditLog

    id                                                            activityDateTime             activityDisplayName   operationType result  initiatedBy   
    --                                                            ----------------             -------------------   ------------- ------  -----------   
    Directory_9af6aff3-dc09-4ac1-a1d3-143e80977b3e_EZPWC_41985545 2020-05-29T07:57:51.4037921Z Add service principal Add           success @{user=; app=}
    Directory_f830a9d4-e746-48dc-944c-eb093364c011_1ZJAE_22273050 2020-05-29T07:57:51.6245497Z Add service principal Add           failure @{user=; app=}
    Directory_a813bc02-5d7a-4a40-9d37-7d4081d42b42_RKRRS_12877155 2020-06-02T12:49:38.5177891Z Add user              Add           success @{app=; user=}

    PS C:\>Get-AADIntAzureAuditLog -EntryId Directory_9af6aff3-dc09-4ac1-a1d3-143e80977b3e_EZPWC_41985545

    id                  : Directory_9af6aff3-dc09-4ac1-a1d3-143e80977b3e_EZPWC_41985545
    category            : ApplicationManagement
    correlationId       : 9af6aff3-dc09-4ac1-a1d3-143e80977b3e
    result              : success
    resultReason        : 
    activityDisplayName : Add service principal
    activityDateTime    : 2020-05-29T07:57:51.4037921Z
    loggedByService     : Core Directory
    operationType       : Add
    initiatedBy         : @{user=; app=}
    targetResources     : {@{id=66ce0b00-92ee-4851-8495-7c144b77601f; displayName=Azure Credential Configuration Endpoint Service; type=ServicePrincipal; userPrincipalName=; 
                          groupType=; modifiedProperties=System.Object[]}}
    additionalDetails   : {}
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$EntryId,
        [switch]$Export
        
    )
    Process
    {

        # Select one entry if provided
        if($EntryId)
        {
            $queryString = "`$filter=id eq '$EntryId'"
        }
        else
        {
            $queryString = "`$top=50&`$orderby=activityDateTime"
        }

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "auditLogs/directoryAudits" -QueryString $queryString

        # Return full results
        if($Export)
        {
            return $results
        }
        elseif($EntryId) # The single entry
        {
            return $results
        }
        else # Print out only some info - the API always returns all info as $Select is not supported :(
        {
            $results | select id,activityDateTime,activityDisplayName,operationType,result,initiatedBy | ft
        }
    }
}

# Gets the user's data
function Get-MSGraphUser
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $API = "users/$UserPrincipalName"
        $ApiVersion = "beta"
        $querystring = "`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,onPremisesDistinguishedName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSecurityIdentifier,refreshTokensValidFromDateTime,signInSessionsValidFromDateTime,usageLocation,provisionedPlans,proxyAddresses"

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -QueryString $querystring
        
        return $results
    }
}

# Gets the user's application role assignments
function Get-MSGraphUserAppRoleAssignments
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/appRoleAssignments" -ApiVersion v1.0

        return $results
    }
}

# Gets the user's owned devices
function Get-MSGraphUserOwnedDevices
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/ownedDevices" -ApiVersion v1.0

        return $results
    }
}

# Gets the user's registered devices
function Get-MSGraphUserRegisteredDevices
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/registeredDevices" -ApiVersion v1.0

        return $results
    }
}

# Gets the user's licenses
function Get-MSGraphUserLicenseDetails
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/licenseDetails" -ApiVersion v1.0 

        return $results
    }
}

# Gets the user's groups
function Get-MSGraphUserMemberOf
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/memberOf" -ApiVersion v1.0

        return $results
    }
}

# Gets the user's direct reports
function Get-MSGraphUserDirectReports
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$UserPrincipalName
    )
    Process
    {
        # Url encode for external users, replace # with %23
        $UserPrincipalName = $UserPrincipalName.Replace("#","%23")

        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "users/$UserPrincipalName/directReports" -ApiVersion v1.0 -QueryString "`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,onPremisesDistinguishedName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSecurityIdentifier,refreshTokensValidFromDateTime,signInSessionsValidFromDateTime,usageLocation,provisionedPlans,proxyAddresses"

        return $results
    }
}

# Gets the group's owners
function Get-MSGraphGroupOwners
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$GroupId
    )
    Process
    {
        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "groups/$GroupId/owners" -ApiVersion v1.0 -QueryString "`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,onPremisesDistinguishedName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSecurityIdentifier,refreshTokensValidFromDateTime,signInSessionsValidFromDateTime,usageLocation,provisionedPlans,proxyAddresses"

        return $results
    }
}

# Gets the group's members
function Get-MSGraphGroupMembers
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$False)]
        [String]$GroupId
    )
    Process
    {
        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "groups/$GroupId/members" -ApiVersion v1.0 -QueryString "`$top=500&`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,onPremisesDistinguishedName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSecurityIdentifier,refreshTokensValidFromDateTime,signInSessionsValidFromDateTime,usageLocation,provisionedPlans,proxyAddresses"

        return $results
    }
}


# Gets the aad roles 
function Get-MSGraphRoles
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$true)]
        [String]$RoleId
    )
    Process
    {
        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "directoryRoles" -ApiVersion v1.0 -QueryString "`$select=id,displayName,description,roleTemplateId"

        return $results
    }
}


# Gets the aad role members
function Get-MSGraphRoleMembers
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$true)]
        [String]$RoleId
    )
    Process
    {
        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "directoryRoles/$RoleId/members" -ApiVersion v1.0 -QueryString "`$select=businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,onPremisesDistinguishedName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSecurityIdentifier,refreshTokensValidFromDateTime,signInSessionsValidFromDateTime,usageLocation,provisionedPlans,proxyAddresses"

        return $results
    }
}

# Gets the tenant domains (all of them)
function Get-MSGraphDomains
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken
    )
    Process
    {
        $results=Call-MSGraphAPI -AccessToken $AccessToken -API "domains" -ApiVersion beta

        return $results
    }
}


# Gets the authorizationPolicy
function Get-TenantAuthPolicy
{
<#
    .SYNOPSIS
    Gets tenant's authorization policy.

    .DESCRIPTION
    Gets tenant's authorization policy, including user and guest settings.

    .PARAMETER AccessToken
    Access token used to retrieve the authorization policy.

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Get-AADIntTenantAuthPolicy

    id                                                : authorizationPolicy
    allowInvitesFrom                                  : everyone
    allowedToSignUpEmailBasedSubscriptions            : True
    allowedToUseSSPR                                  : True
    allowEmailVerifiedUsersToJoinOrganization         : False
    blockMsolPowerShell                               : False
    displayName                                       : Authorization Policy
    description                                       : Used to manage authorization related settings across the company.
    enabledPreviewFeatures                            : {}
    guestUserRoleId                                   : 10dae51f-b6af-4016-8d66-8c2a99b929b3
    permissionGrantPolicyIdsAssignedToDefaultUserRole : {microsoft-user-default-legacy}
    defaultUserRolePermissions                        : @{allowedToCreateApps=True; allowedToCreateSecurityGroups=True; allowedToReadOtherUsers=True}

#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    Process
    {      
        
        $results = Call-MSGraphAPI -AccessToken $AccessToken -API "policies/authorizationPolicy" 
        return $results
    }
}

# Gets the guest account restrictions
function Get-TenantGuestAccess
{
<#
    .SYNOPSIS
    Gets the guest access level of the user's tenant.

    .DESCRIPTION
    Gets the guest access level of the user's tenant.

    Inclusive:  Guest users have the same access as members
    Normal:     Guest users have limited access to properties and memberships of directory objects
    Restricted: Guest user access is restricted to properties and memberships of their own directory objects (most restrictive)

    .PARAMETER AccessToken
    Access token used to retrieve the access level.

    .Example
    Get-AADIntAccessTokenForMSGraph -SaveToCache
    PS C:\>Get-AADIntTenantGuestAccess

    Access Description                                                                        RoleId                              
    ------ -----------                                                                        ------                              
    Normal Guest users have limited access to properties and memberships of directory objects 10dae51f-b6af-4016-8d66-8c2a99b929b3
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    Process
    {

        $policy = Get-TenantAuthPolicy -AccessToken $AccessToken

        $roleId = $policy.guestUserRoleId

        
        switch($roleId)
        {
            "a0b1b346-4d3e-4e8b-98f8-753987be4970" {
                $attributes=[ordered]@{
                    "Access" =      "Full"
                    "Description" = "Guest users have the same access as members"
                }
                break
            }
            "10dae51f-b6af-4016-8d66-8c2a99b929b3" {
                $attributes=[ordered]@{
                    "Access" =      "Normal"
                    "Description" = "Guest users have limited access to properties and memberships of directory objects"
                }
                break
            }
            "2af84b1e-32c8-42b7-82bc-daa82404023b" {
                $attributes=[ordered]@{
                    "Access" =      "Restricted"
                    "Description" = "Guest user access is restricted to properties and memberships of their own directory objects (most restrictive)"
                }
                break
            }
        }

        $attributes["RoleId"] = $roleId

        return New-Object psobject -Property $attributes


    }
}

# Sets the guest account restrictions
function Set-TenantGuestAccess
{
<#
    .SYNOPSIS
    Sets the guest access level for the user's tenant.

    .DESCRIPTION
    Sets the guest access level for the user's tenant.

    Inclusive:  Guest users have the same access as members
    Normal:     Guest users have limited access to properties and memberships of directory objects
    Restricted: Guest user access is restricted to properties and memberships of their own directory objects (most restrictive)

    .PARAMETER AccessToken
    Access token used to retrieve the access level.

    .PARAMETER Level
    Guest access level. One of Inclusive, Normal, or Restricted.

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Set-AADIntTenantGuestAccess -Level Normal

    Access Description                                                                        RoleId                              
    ------ -----------                                                                        ------                              
    Normal Guest users have limited access to properties and memberships of directory objects 10dae51f-b6af-4016-8d66-8c2a99b929b3
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken,
        
        [Parameter(Mandatory=$True)]
        [ValidateSet('Full','Normal','Restricted')]
        [String]$Level
    )
    Process
    {

        switch($Level)
        {
            "Full"       {$roleId = "a0b1b346-4d3e-4e8b-98f8-753987be4970"; break}
            "Normal"     {$roleId = "10dae51f-b6af-4016-8d66-8c2a99b929b3"; break}
            "Restricted" {$roleId = "2af84b1e-32c8-42b7-82bc-daa82404023b"; break}
        }

        $body = @{
            "guestUserRoleId" = $roleId
        }
        
        # "{""guestUserRoleId"":""$roleId""}"


        Call-MSGraphAPI -AccessToken $AccessToken -API "policies/authorizationPolicy/authorizationPolicy" -Method "PATCH" -Body $body

        Get-TenantGuestAccess -AccessToken $AccessToken

    }
}


# Enables Msol PowerShell access
function Enable-TenantMsolAccess
{
<#
    .SYNOPSIS
    Enables Msol PowerShell module access for the user's tenant.

    .DESCRIPTION
    Enables Msol PowerShell module access for the user's tenant.

    .PARAMETER AccessToken
    Access token used to enable the Msol PowerShell access.

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Enable-AADIntTenantMsolAccess

#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    Process
    {


        $body = @{
            "blockMsolPowerShell" = "false"
        }

        # '{"blockMsolPowerShell":"false"}'

        Call-MSGraphAPI -AccessToken $AccessToken -API "policies/authorizationPolicy/authorizationPolicy" -Method "PATCH" -Body $body
    }
}

# Disables Msol PowerShell access
function Disable-TenantMsolAccess
{
<#
    .SYNOPSIS
    Disables Msol PowerShell module access for the user's tenant.

    .DESCRIPTION
    Disables Msol PowerShell module access for the user's tenant.

    .PARAMETER AccessToken
    Access token used to disable the Msol PowerShell access.

    .Example
    Get-AADIntAccessTokenForMSGraph
    PS C:\>Disable-AADIntTenantMsolAccess

#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    Process
    {

        $body = @{
            "blockMsolPowerShell" = "true"
        }

        # '{"blockMsolPowerShell":"true"}'

        Call-MSGraphAPI -AccessToken $AccessToken -API "policies/authorizationPolicy/authorizationPolicy" -Method "PATCH" -Body $body
    }
}


# get app service principals
function get-MSGraphServicePrincipal
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$ObjectId, # service principal objectId
        [Parameter(Mandatory=$false)]
        [String]$appid  # application Id
    )
    Process
    {

        

        # filter by appid if provided 
        if (![String]::IsNullOrEmpty($appid)){
             # Url encode for external users, replace # with %23
             $filter = "`$filter=appId eq '$appid'"
             $API="/servicePrincipals"
        } elseif(![String]::IsNullOrEmpty($ObjectId)) {
            $API="/servicePrincipals/$ObjectId"
            $filter = $null
        } else {
            $API="/servicePrincipals"
            $filter =$null
        }
        $ApiVersion = "v1.0"

        try {
            $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -QueryString $filter
            return $results
        }
        catch {
            write-error "failed to call msgraph. please check if the right clientId or ObjectId are provided"
            return $null
        }
    }
}



# get user consents permissions for giving user
function get-MSGraphoauth2permissions
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(ParameterSetName='UserPrincipalName',Mandatory=$True)]
        [String]$UserPrincipalName,
        [Parameter(ParameterSetName='clientId',Mandatory=$True)]
        [String]$clientId,
        [Parameter(ParameterSetName='clientId',Mandatory=$false)]
        [string]$resourceId,
        [Parameter(ParameterSetName='clientId',Mandatory=$false)]
        [switch]$adminconsentonly
    )
    Process
    {

        $ApiVersion = "v1.0"

        # try to get user consent if user name provided
        if (![String]::IsNullOrEmpty($UserPrincipalName)){
             # Url encode for external users, replace # with %23
            $UserPrincipalName = $UserPrincipalName.Replace("#","%23")
            $API = "/users/$UserPrincipalName/oauth2PermissionGrants"
            $filter = $NULL
        } else {
            $API = "/oauth2PermissionGrants"
            if ($adminconsentonly) {
                $filter = "`$filter=clientId eq '$clientId' AND consentType eq 'AllPrincipals'"
            } else {
                $filter = "`$filter=clientId eq '$clientId'"
            }
            
            if (![String]::IsNullOrEmpty($resourceId)){
                $filter = $filter + " AND resourceId eq '$resourceId'"
            }

        }


        try {
            $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -QueryString $filter
            return $results
        }
        catch {
            write-error "failed to call msgraph. please check if the right UserPrincipalName or clientId are provided"
            return $null
        }
    }
}


# grant admin consents permissions for target application
function add-MSGraphAdminconsent
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$clientId,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [String]$resourceId,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Principal','AllPrincipals')]
        [String]$consentType='AllPrincipals',
        [Parameter(Mandatory=$false)]
        [switch]$force
    )
    Process
    {

        $results=add-MSGraphUserconsent -clientId $clientId -resourceId $resourceid -scope $scope -consentType $consentType
        return $results
    }
}

# grant user consents permissions for target application
function add-MSGraphUserconsent
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$UserPrincipalName,
        [Parameter(Mandatory=$True)]
        [String]$clientId,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [String]$resourceId,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Principal','AllPrincipals')]
        [String]$consentType='Principal',
        [Parameter(Mandatory=$false)]
        [switch]$force
    )
    Process
    {

        $API = "/oauth2PermissionGrants"
        $ApiVersion="v1.0"
        # default permission is openid and profile for msgraph
        if ([String]::IsNullOrEmpty($scope)) {
            $scope="openid profile"
        }
        $missedpermissions = ""

        # default resourceId for MS graph API
        if ([String]::IsNullOrEmpty($scope)) {
            $resourceId="a58b0002-fd14-43d8-aa02-521cbb08493a"
        }
        
        if ($UserPrincipalName) {
            $user=Get-MSGraphUser -UserPrincipalName $UserPrincipalName

            if (!$user){
                write-error "not find user object for $UserPrincipalName"
                return $NULL
            }
        }

        # Check if the admin consent has been granted already
        $adminsonentpermissions = get-MSGraphoauth2permissions -AccessToken $AccessToken -clientid $clientId -resourceId $resourceId -adminconsentonly 
        if ($adminsonentpermissions) {

            $fullscopes = [string]::Join(" ",$adminsonentpermissions.scope).split(" ")
            $scopeitems =  $scope.split(" ")
            foreach ($scopeitem in $scopeitems){
                if (!($fullscopes -contains $scopeitem)) {

                    if ($missedpermissions -eq "") {
                        $missedpermissions=$scopeitem
                    } else {
                        $missedpermissions=$missedpermissions+" "+$scopeitem
                    }
                }
            }

            # skip add user consents if admin consents exists and no force swtich added
            if ($missedpermissions -eq "") {
                write-verbose "admin consents are granted already for client $clientId, please review the permissions if it is required to grant user consents"
                return $NULL
            } 
            
        }
        
        # add missed permissions only if not existing
        if ($consentType -eq 'AllPrincipals') {

            # only grant missed admin consent
            if ($adminsonentpermissions) {

                
                # if admin consent existing. do update only
                if ($missedpermissions -ne ""){
                    $newscopes = $($adminsonentpermissions.scope)+" "+$missedpermissions
                    $API = $API+"/"+$($adminsonentpermissions.id)
                    $oauth2gant = @{
                        "scope"=$newscopes
                    }

                    $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -method PATCH -body $oauth2gant
                }

            } else {

                $oauth2gant = @{
                    "clientId"= $clientId
                    "consentType"=$consentType
                    "resourceId"=$resourceId
                    "scope"=$missedpermissions
                }
                $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -method POST -body $oauth2gant

            }

        } else {

            # get all user consent permissions
            $oauth2permissions = get-MSGraphoauth2permissions -AccessToken $AccessToken -UserPrincipalName $UserPrincipalName 
            
            $clientpermission = $oauth2permissions | where {$_.clientid -eq $clientId -and $_.resourceId -eq $resourceId}
            if ($clientpermission -and $consentType -eq 'Principal') {
                
                if ($force) {
                    write-verbose "force update user permissions with new scope: $scope, it will remove the existing user permissions"
                    clear-MSGraphUserconsent -AccessToken $AccessToken -UserPrincipalName $UserPrincipalName -clientId $clientId -resourceId $resourceId                    

                } else {
                    write-verbose "user consents are granted already for client $clientId, skip the process"
                    write-verbose $clientpermission
                    return $NULL
                }
                
            }             
            
            if ($missedpermissions -eq "") {
            $oauth2gant = @{
                "clientId"= $clientId
                "consentType"=$consentType
                "resourceId"=$resourceId
                "scope"=$scope
                "principalId"=$user.id
            }

            } else {

            $oauth2gant = @{
                "clientId"= $clientId
                "consentType"=$consentType
                "resourceId"=$resourceId
                "scope"=$missedpermissions
                "principalId"=$user.id
            }
            }

            $results=Call-MSGraphAPI -AccessToken $AccessToken -API $API -ApiVersion $ApiVersion -method POST -body $oauth2gant
        
        }

        return $results
    }
}



# evaluate consents permissions for target user
function test-MSGraphUserconsent
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$UserPrincipalName,
        [Parameter(Mandatory=$True)]
        [String]$clientId,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [String]$resourceId
    )
    Process
    {

        # default permission is openid and profile for msgraph
        if ([String]::IsNullOrEmpty($scope)) {
            $scope="openid profile"
        }
        $missedpermissions = ""

        # default resourceId for MS graph API
        if ([String]::IsNullOrEmpty($resourceId)) {
            $resourceId="a58b0002-fd14-43d8-aa02-521cbb08493a"
        }
        
        $user=Get-MSGraphUser -UserPrincipalName $UserPrincipalName

        if (!$user){
            write-error "not find user object for $UserPrincipalName"
            return $false
        }

        # Check if the admin consent has been granted already
        $adminsonentpermissions = get-MSGraphoauth2permissions -AccessToken $AccessToken -clientid $clientId -resourceId $resourceId -adminconsentonly 
        if ($adminsonentpermissions) {

            $fullscopes = [string]::Join(" ",$adminsonentpermissions.scope).split(" ")
            $scopeitems =  $scope.split(" ")
            foreach ($scopeitem in $scopeitems){
                if (!($fullscopes -contains $scopeitem)) {

                    if ($missedpermissions -eq "") {
                        $missedpermissions=$scopeitem
                    } else {
                        $missedpermissions=$missedpermissions+" "+$scopeitem
                    }
                }
            }

            # skip add user consents if admin consents exists and no force swtich added
            if ($missedpermissions -eq "") {
                write-verbose "admin consents are granted already for client $clientId, please review the permissions if it is required to grant user consents"
                return $true
            } 
            
        }

        # get all user consent permissions
        $oauth2permissions = get-MSGraphoauth2permissions -AccessToken $AccessToken -UserPrincipalName $UserPrincipalName 
        
        $clientpermission = $oauth2permissions | where {$_.clientid -eq $clientId -and $_.resourceId -eq $resourceId}
        if ($clientpermission) {
            
            $fullscopes = [string]::Join(" ",$clientpermission.scope).split(" ")

            if ($missedpermissions -eq "") {
                $scopeitems =  $scope.split(" ")
            } else {
                $scopeitems =  $missedpermissions.split(" ")
                $missedpermissions = ""
            }

            foreach ($scopeitem in $scopeitems){
                if (!($fullscopes -contains $scopeitem)) {

                    if ($missedpermissions -eq "") {
                        $missedpermissions=$scopeitem
                    } else {
                        $missedpermissions=$missedpermissions+" "+$scopeitem
                    }
                }
            }

            # skip add user consents if admin consents exists and no force swtich added
            if ($missedpermissions -eq "") {
                write-verbose "user consents are granted already for client $clientId"
                return $true
            } else {
                write-verbose "user consents are granted but missing permissions: $missedpermissions"
                return $false
            }
          

        } else{

            write-verbose "no user consents are granted"
            return $false
        }
     
    }
}



# remove user consents permissions for target application
function clear-MSGraphUserconsent
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$UserPrincipalName,
        [Parameter(Mandatory=$True)]
        [String]$clientId,
        [Parameter(Mandatory=$false)]
        [String]$resourceId
    )
    Process
    {
        # default resourceId for MS graph API
        if ([String]::IsNullOrEmpty($scope)) {
            $resourceId="a58b0002-fd14-43d8-aa02-521cbb08493a"
        }
        
        # get all user consent permissions
        $oauth2permissions = get-MSGraphoauth2permissions -AccessToken $AccessToken -UserPrincipalName $UserPrincipalName 
        
        $clientpermissions = $oauth2permissions | where {$_.clientid -eq $clientId -and $_.resourceId -eq $resourceId}

        foreach ($clientpermission in $clientpermissions) {
            $API = "/oAuth2PermissionGrants/$($clientpermission.Id)"
            call-msgraphapi -api $API -apiversion "v1.0" -method delete
        }
    }
}