

# Get conditonal access policy using Azure Management API
# only support AADIAM
function Get-AADConditionalPolicies
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        $AccessToken,
        [Parameter(Mandatory=$false)]
        $top=100
    )
    Process
    {
        $response=Call-AzureAADIAMAPI -AccessToken $AccessToken -Command "Policies/Policies?top=$top&nextLink=null&appId=&includeBaseline=true"
        return $response.items
    }
}


# Checks whether the external user is unique or already exists in AAD
function Is-ExternalUserUnique
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        $AccessToken,
        [Parameter(Mandatory=$True)]
        [string]$EmailAddress
        
    )
    Process
    {
         $result = Call-AzureAADIAMAPI -AccessToken $AccessToken -Command "Users/IsUPNUniqueOrPending/$EmailAddress" 

         if ($result -eq "Unique") {
            return $true 
         } else {
            return $false
         }
    }
}

# Get the user roles based on a subscription scope
function Get-AzureRBACroles
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        $AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$subscriptionID,
        [Parameter(Mandatory=$false)]
        [String]$apiversion="2018-01-01-preview",
        [Parameter(Mandatory=$false)]
        [String]$filter  # such as type eq 'customrole'
    )
    Process
    {

        $command = "subscriptions/$subscriptionID/providers/Microsoft.Authorization/roleDefinitions"
       
        if (![string]::IsNullOrEmpty( $filter)) {
            $filterstring =  [System.Web.HttpUtility]::UrlEncode("`$filter=$filter").replace("+","%20").replace("%3d","=")
            $command=$command+"`?$filterstring"
        }

        try {
            $roles =  Call-AzureManagementAPI -AccessToken $AccessToken -Command  $command
               
            if($roles.properties) {
                return $roles.properties
            } else {
                return $NULL
            }
        }
        catch {

            write-verbose "failed to get RBAC roles"
        }

    }
}

# Gets Azure Tenant authentication methods
function Get-TenantAuthenticationMethods
{
<#
    .SYNOPSIS
    Gets Azure tenant authentication methods. 

    .DESCRIPTION
    Gets Azure tenant authentication methods. 

    
    .Example
    Get-AADIntAccessTokenForAADIAMAPI

    Tenant                               User Resource                             Client                              
    ------                               ---- --------                             ------                              
    6e3846ee-e8ca-4609-a3ab-f405cfbd02cd      74658136-14ec-4630-ad9b-26e160ff0fc6 d3590ed6-52b3-4102-aeff-aad2292ab01c

    PS C:\>Get-AADIntTenantAuthenticationMethods

    id                : 297c50d5-e789-40f7-8931-b3694713cb4d
    type              : 6
    state             : 0
    includeConditions : {@{type=group; id=9202b94b-5381-4270-a3cb-7fcf0d40fef1; isRequired=False; useForSignIn=True}}
    voiceSettings     : 
    fidoSettings      : @{allowSelfServiceSetup=False; enforceAttestation=False; keyRestrictions=}
    enabled           : True
    method            : FIDO2 Security Key

    id                : 3d2c4b8f-f362-4ce4-8f4b-cc8726b80106
    type              : 8
    state             : 1
    includeConditions : {@{type=group; id=all_users; isRequired=False; useForSignIn=True}}
    voiceSettings     : 
    fidoSettings      : 
    enabled           : False
    method            : Microsoft Authenticator passwordless sign-in

    id                : d7716fe0-7c2e-4b52-a5cd-394f8999176b
    type              : 5
    state             : 1
    includeConditions : {@{type=group; id=all_users; isRequired=False; useForSignIn=True}}
    voiceSettings     : 
    fidoSettings      : 
    enabled           : False
    method            : Text message

#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    Process
    {
        # Get the authentication methods
        $response =  Call-AzureAADIAMAPI -AccessToken $AccessToken -Command "AuthenticationMethods/AuthenticationMethodsPolicy"

        $methods = $response.authenticationMethods
        foreach($method in $methods)
        {
            $strType="unknown"
            switch($method.type)
            {
                6 {$strType = "FIDO2 Security Key"}
                8 {$strType = "Microsoft Authenticator passwordless sign-in"}
                5 {$strType = "Text message"}
            }

            $method | Add-Member -NotePropertyName "enabled" -NotePropertyValue ($method.state -eq 0)
            $method | Add-Member -NotePropertyName "method"  -NotePropertyValue $strType

        }

        return $methods
        
    }
}


# Gets Azure Tenant applications
function Get-TenantApplications
{
<#
    .SYNOPSIS
    Gets Azure tenant applications.

    .DESCRIPTION
    Gets Azure tenant applications.
    
    .Example
    Get-AADIntAccessTokenForAADIAMAPI -SaveToCache

    Tenant                               User Resource                             Client                              
    ------                               ---- --------                             ------                              
    6e3846ee-e8ca-4609-a3ab-f405cfbd02cd      https://management.core.windows.net/ d3590ed6-52b3-4102-aeff-aad2292ab01c

    PS C:\>Get-AADIntTenantApplications

    
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken
    )
    process
    {

        $body = @{
            "accountEnabled" =       $null
            "isAppVisible" =         $null
            "appListQuery"=          0
            "top" =                  999
            "loadLogo" =             $false
            "putCachedLogoUrlOnly" = $true
            "nextLink" =             ""
            "usedFirstPartyAppIds" = $null
            "__ko_mapping__" = @{
                "ignore" = @()
                "include" = @("_destroy")
                "copy" = @()
                "observe" = @()
                "mappedProperties" = @{
                    "accountEnabled" =       $true
                    "isAppVisible" =         $true
                    "appListQuery" =         $true
                    "searchText" =           $true
                    "top" =                  $true
                    "loadLogo" =             $true
                    "putCachedLogoUrlOnly" = $true
                    "nextLink" =             $true
                    "usedFirstPartyAppIds" = $true
                }
                "copiedProperties" = @{}
            }
        }

        # Get the applications
        $response =  Call-AzureAADIAMAPI -AccessToken $AccessToken -Command "ManagedApplications/List" -Body $body -Method Post

        return $response.appList
       
    }
}

# Get the status of AAD Connect
function Get-AADConnectStatus
{
<#
    .SYNOPSIS
    Shows the status of Azure AD Connect (AAD Connect).

    .DESCRIPTION
    Shows the status of Azure AD Connect (AAD Connect).

    .Example
    Get-AADIntAccessTokenForAADIAMAPI -SaveToCache
    PS C:\>Get-AADIntAADConnectStatus

    verifiedDomainCount              : 4
    verifiedCustomDomainCount        : 3
    federatedDomainCount             : 2
    numberOfHoursFromLastSync        : 0
    dirSyncEnabled                   : True
    dirSyncConfigured                : True
    passThroughAuthenticationEnabled : True
    seamlessSingleSignOnEnabled      : True
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        $AccessToken
    )
    Process
    {

        # Get the applications
        $response =  Call-AzureAADIAMAPI -AccessToken $AccessToken -Command "Directories/ADConnectStatus" 

        return $response
    }
}