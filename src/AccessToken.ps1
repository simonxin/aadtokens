# This contains functions for getting Azure AD access tokens

# Tries to get access token from cache unless provided as parameter
# Refactored Jun 8th 2020
function Get-AccessTokenFromCache
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$AccessToken,
        [Parameter(Mandatory=$True)]
        [String]$ClientID,
        [Parameter(Mandatory=$True)]
        [String]$Resource,
        [switch]$IncludeRefreshToken,
        [boolean]$Force=$false,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        # Check if we got the AccessToken as parameter
        if([string]::IsNullOrEmpty($AccessToken))
        {
            # Check if cache entry is empty
            if([string]::IsNullOrEmpty($Script:tokens["$ClientId-$Resource"]))
            {
                # Empty, so throw the exception
                Throw "No saved tokens found. Please call Get-AADIntAccessTokenFor<service> -SaveToCache"
            }
            else
            {
                $retVal=$Script:tokens["$ClientId-$Resource"]
            }
        }
        else
        {
            # Check that the audience of the access token is correct
            $audience=(Read-Accesstoken -AccessToken $AccessToken).aud

            # Strip the trailing slashes
            if($audience.EndsWith("/"))
            {
                $audience = $audience.Substring(0,$audience.Length-1)
            }
            if($Resource.EndsWith("/"))
            {
                $Resource = $Resource.Substring(0,$Resource.Length-1)
            }

            if(($audience -ne $Resource) -and ($Force -eq $False))
            {
                # Wrong audience
                Write-Verbose "ACCESS TOKEN HAS WRONG AUDIENCE: $audience. Exptected: $resource."
                Throw "The audience of the access token ($audience) is wrong. Should be $resource!"
            }
            else
            {
                # Just return the passed access token
                $retVal=$AccessToken
            }
        }

        # Check the expiration
        if(Is-AccessTokenExpired($retVal))
        {
            Write-Verbose "ACCESS TOKEN HAS EXPRIRED. Trying to get a new one with RefreshToken."
            $retVal = Get-AccessTokenWithRefreshToken -cloud $Cloud -Resource $Resource -ClientId $ClientID -RefreshToken $script:refresh_tokens["$ClientId-$Resource"] -TenantId (Read-Accesstoken -AccessToken $retVal).tid -SaveToCache $true -IncludeRefreshToken $IncludeRefreshToken
        }

        # Return
        return $retVal
    }
}

# Gets the access token for AAD Graph API
function Get-AccessTokenForAADGraph
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for AAD Graph

    .DESCRIPTION
    Gets OAuth Access Token for AAD Graph, which is used for example in Provisioning API.
    If credentials are not given, prompts for credentials (supports MFA).

    .Parameter Credentials
    Credentials of the user. If not given, credentials are prompted.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos ticket

    .Parameter KerberosTicket
    Kerberos token of the user.

    .Parameter UseDeviceCode
    Use device code flow.

    .Parameter Resource
    Resource, defaults to aad graph API 
    
    .Example
    Get-AADIntAccessTokenForAADGraph
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForAADGraph -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $resource = $script:AzureResources[$Cloud]['aad_graph_api'] # get AAD graph resource based on cloud
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1b730954-1685-4b74-9bfd-dac224a7b894" which is MS Graph API

        Get-AccessToken -cloud $Cloud -Credentials $Credentials -RedirectUri $RedirectUri -Resource $Resource -ClientId $clientId -SAMLToken $SAMLToken -Tenant $Tenant -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets the access token for MS Graph API
function Get-AccessTokenForMSGraph
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Microsoft Graph

    .DESCRIPTION
    Gets OAuth Access Token for Microsoft Graph, which is used in Graph API.
    If credentials are not given, prompts for credentials (supports MFA).

    .Parameter Credentials
    Credentials of the user. If not given, credentials are prompted.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user.

    .Example
    Get-AADIntAccessTokenForMSGraph
    
    .Example
    $cred=Get-Credential
    Get-AADIntAccessTokenForMSGraph -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["ms_graph_api"] # get MS graph resource based on cloud
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1b730954-1685-4b74-9bfd-dac224a7b894" which is MS Graph API

        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -SAMLToken $SAMLToken -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets the access token for enabling or disabling PTA
function Get-AccessTokenForPTA
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for PTA

    .DESCRIPTION
    Gets OAuth Access Token for PTA, which is used for example to enable or disable PTA.

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForPTA
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForPTA -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process    
    {

        $resource = $script:AzureResources[$Cloud]["cloudwebappproxy"] # get PTA resource based on cloud
        $clientId = $script:AzureKnwonClients["aadsync"] # set client Id = "cb1056e2-e479-49de-ae31-7812af012ed8" which is aad sync app


        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource $resource -RedirectUri $RedirectUri  -ClientId $clientId -SAMLToken $SAMLToken -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets the access token for Office Apps
function Get-AccessTokenForOfficeApps
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Office Apps

    .DESCRIPTION
    Gets OAuth Access Token for Office Apps.

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForOfficeApps
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForOfficeApps -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["officeapps"] # get office apps resource based on cloud
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1b730954-1685-4b74-9bfd-dac224a7b894" which is MS Graph API

        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource  $resource -RedirectUri $RedirectUri -ClientId  $clientId -SAMLToken $SAMLToken -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets the access token for Exchange Online
function Get-AccessTokenForEXO
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Exchange Online

    .DESCRIPTION
    Gets OAuth Access Token for Exchange Online

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForEXO
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForEXO -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["outlook"] # get outlook online apps resource based on cloud
        $clientId = $script:AzureKnwonClients["office"] # set client Id = "d3590ed6-52b3-4102-aeff-aad2292ab01c" which is office API client app

        # Office app has the required rights to Exchange Online
        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource $Resource -RedirectUri $RedirectUri -ClientId $clientId -SAMLToken $SAMLToken -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets the access token for Exchange Online remote PowerShell
function Get-AccessTokenForEXOPS
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Exchange Online remote PowerShell

    .DESCRIPTION
    Gets OAuth Access Token for Exchange Online remote PowerShell

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter Certificate
    x509 device certificate.
    
    .Example
    Get-AADIntAccessTokenForEXOPS
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForEXOPS -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["outlook"] # get outlook online apps resource based on cloud
        $clientId = $script:AzureKnwonClients["exo"] # set client Id = "a0c73c16-a7e3-4564-9a95-2bdf47383716" which is EXO Remote PowerShell

        # Office app has the required rights to Exchange Online
        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -SAMLToken $SAMLToken -KerberosTicket $KerberosTicket -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode -Domain $Domain
    }
}

# Gets the access token for SARA
# Jul 8th 2019
function Get-AccessTokenForSARA
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for SARA

    .DESCRIPTION
    Gets OAuth Access Token for Microsoft Support and Recovery Assistant (SARA)

    .Parameter KerberosTicket
    Kerberos token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token. 
    
    .Example
    Get-AADIntAccessTokenForSARA
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForSARA -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $resource = $script:AzureResources[$Cloud]["sara"] # get outlook online apps resource based on cloud
        $clientId = $script:AzureKnwonClients["office"] # set client Id = "d3590ed6-52b3-4102-aeff-aad2292ab01c" which is sara app
        # Office app has the required rights to Exchange Online
        Get-AccessToken -cloud $Cloud -Credentials $Credentials -Tenant $Tenant -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets an access token for OneDrive
# Nov 26th 2019
function Get-AccessTokenForOneDrive
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for OneDrive

    .DESCRIPTION
    Gets OAuth Access Token for OneDrive Sync client

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForOneDrive
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForOneDrive -Tenant "company" -Credentials $cred
#>
    [cmdletbinding()]
    Param(
    [Parameter(Mandatory=$True)]
        [String]$Tenant,
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {        
        $resource = $script:AzureResources[$Cloud]['suffixes']["sharepoint"] # get sharepoint apps resource based on cloud
        $clientId = $script:AzureKnwonClients["onedrive"] # set client Id = "ab9b8c07-8f02-4f72-87fa-80105867a763" which is one drive client app

        Get-AccessToken -cloud $Cloud -Tenant $Tenant -Resource "https://$Tenant-my$resource/" -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials  -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets an access token for Azure Core Management
# May 29th 2020
function Get-AccessTokenForAzureCoreManagement
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Azure Core Management

    .DESCRIPTION
    Gets OAuth Access Token for Azure Core Management

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token
    
    .Example
    Get-AADIntAccessTokenForOneOfficeApps
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForAzureCoreManagement -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $resource = $script:AzureResources[$Cloud]["windows_net_mgmt_api"] # get windows core management api resource
        $clientId = $script:AzureKnwonClients["office"]
     ##   $clientId = $script:AzureKnwonClients["graph_api"] # set client Id =  1b730954-1685-4b74-9bfd-dac224a7b894 which is azure graph api
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets an access token for SPO
# Jun 10th 2020
function Get-AccessTokenForSPO
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for SharePoint Online

    .DESCRIPTION
    Gets OAuth Access Token for SharePoint Online Management Shell, which can be used with any SPO requests.

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter Tenant
    The tenant name of the organization, ie. company.onmicrosoft.com -> "company"

    .Parameter Admin
    Get the token for admin portal
    
    .Example
    Get-AADIntAccessTokenForSPO
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForSPO -Credentials $cred -Tenant "company"
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [Parameter(Mandatory=$True)]
        [String]$Tenant,
        [switch]$SaveToCache,
        [switch]$Admin,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        if($Admin)
        {
            $prefix = "-admin"
        }

        $resource = $script:AzureResources[$Cloud]["sharepoint"] # get windows core management api resource
        $clientId = $script:AzureKnwonClients["spo_shell"] # set client Id =  "9bc3ab49-b65d-410a-85ad-de819febfddc" which is SPO management shell       

        Get-AccessToken -cloud $Cloud -Tenant $Tenant -Resource "https://$Tenant$prefix.$resource/" -RedirectUri $RedirectUri -ClientId  $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode 
    }
}

# Gets the access token for My Signins
function Get-AccessTokenForMySignins
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for My Signins

    .DESCRIPTION
    Gets OAuth Access Token for My Signins, which is used for example when registering MFA.
   
    .Example
    PS C:\>Get-AADIntAccessTokenForMySignins
#>
    [cmdletbinding()]
    Param(
        [switch]$SaveToCache,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        
        $resource = $script:AzureKnwonClients["aad_account"] # set client Id = "0000000c-0000-0000-c000-000000000000" which is AAD account    
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1b730954-1685-4b74-9bfd-dac224a7b894" which is graph api   

        return Get-AccessToken -cloud $Cloud -ClientId $clientId -Resource $resource -ForceMFA $true -SaveToCache $SaveToCache
    }
}


# Gets an access token for Azure AD Join
# Aug 26th 2020
function Get-AccessTokenForAADJoin
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Azure AD Join

    .DESCRIPTION
    Gets OAuth Access Token for Azure AD Join, allowing users' to register devices to Azure AD.

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.

    .Parameter BPRT
    Bulk PRT token, can be created with New-AADIntBulkPRTToken
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter Tenant
    The tenant name of the organization, ie. company.onmicrosoft.com -> "company"
    
    .Example
    Get-AADIntAccessTokenForAADJoin
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForAADJoin -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$False)]
        [Switch]$Device,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [Parameter(ParameterSetName='BPRT',Mandatory=$True)]
        [string]$BPRT,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [switch]$SaveToCache,        
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        if($Device)
        {
            Get-AccessTokenWithDeviceSAML -SAML $SAMLToken -SaveToCache $SaveToCache
        }
        else
        {

            $resource ="urn:ms-drs:"+$script:AzureResources[$Cloud]["devicemanagementsvc"] # get Device Registration Service
            $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1b730954-1685-4b74-9bfd-dac224a7b894" which is graph api 
            
    
           Get-AccessToken -cloud $Cloud -ClientID $clientId -Resource $resource -RedirectUri $RedirectUri -Tenant $Tenant -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode -ForceMFA $true -BPRT $BPRT
        }
    }
}

# Gets an access token for Intune MDM
# Aug 26th 2020
function Get-AccessTokenForIntuneMDM
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Intune MDM

    .DESCRIPTION
    Gets OAuth Access Token for Intune MDM, allowing users' to enroll their devices to Intune.

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter BPRT
    Bulk PRT token, can be created with New-AADIntBulkPRTToken

    .Parameter Tenant
    The tenant name of the organization, ie. company.onmicrosoft.com -> "company"

    .Parameter Certificate
    x509 device certificate.

    .Parameter TransportKeyFileName
    File name of the transport key

    .Parameter PfxFileName
    File name of the .pfx device certificate.

    .Parameter PfxPassword
    The password of the .pfx device certificate.

    .Parameter Resource
    The resource to get access token to, defaults to "https://enrollment.manage.microsoft.com/". To get access to AAD Graph API, use "https://graph.windows.net"
    
    .Example
    Get-AADIntAccessTokenForIntuneMDM
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForIntuneMDM -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [switch]$ForceMFA,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [Parameter(ParameterSetName='BPRT',Mandatory=$True)]
        [string]$BPRT,
        [switch]$SaveToCache,
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["mdm"] # set resource to Intune MDM enrollment 
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "29d9ed98-a469-4536-ade2-f981bc1d605e" which is Microsoft Authentication Broker (Azure MDM client)

       Get-AccessToken -cloud $Cloud -ClientID $clientId -Tenant $Tenant -Resource $Resource -RedirectUri $RedirectUri -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode -Certificate $Certificate -PfxFileName $PfxFileName -PfxPassword $PfxPassword -BPRT $BPRT -ForceMFA $ForceMFA -TransportKeyFileName $TransportKeyFileName
    }
}

# Gets an access token for Azure Cloud Shell
# Sep 9th 2020
function Get-AccessTokenForCloudShell
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Azure Cloud Shell

    .DESCRIPTION
    Gets OAuth Access Token for Azure Cloud Shell

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token
    
    .Example
    Get-AADIntAccessTokenForOneOfficeApps
    
    .Example
    PS C:\>$cred=Get-Credential
    PS C:\>Get-AADIntAccessTokenForCloudShell -Credentials $cred
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["windows_net_mgmt_api"] # get windows core management api resource
        $clientId = $script:AzureKnwonClients["android"] # set client Id = "0c1307d4-29d6-4389-a11c-5cbe7f65d7fa" which is android app

        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets an access token for Teams
# Oct 3rd 2020
function Get-AccessTokenForTeams
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Teams

    .DESCRIPTION
    Gets OAuth Access Token for Teams

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token
    
    .Example
    Get-AADIntAccessTokenForTeams
    
    .Example
    PS C:\>Get-AADIntAccessTokenForTeams -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]["spacesapi"] # set resource to skype/teams
        $clientId = $script:AzureKnwonClients["teams"] # set client Id = "1fec8e78-bce4-4aaf-ab1b-5451cc387264" which is teams app

        Get-AccessToken -cloud $Cloud -Resource $Resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}


# Gets an access token for Azure AD Management API
function Get-AccessTokenForAADIAMAPI
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for Azure AD IAM API

    .DESCRIPTION
    Gets OAuth Access Token for Azure AD IAM API

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token
    
    .Example
    Get-AADIntAccessTokenForAADIAMAPI
    
    .Example
    PS C:\>Get-AADIntAccessTokenForAADIAMAPI -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$forcemfa,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        # First get the access token for AADGraph
 
        $clientId = $script:AzureKnwonClients['office'] # client ID of azure portal
        $resource = $script:AzureResources[$Cloud]["azure_mgmt_api"] # get AAD graph resource based on cloud
        $aadiamapi = $script:AzureKnwonClients["adibizaux"] # aad iam
        # $clientId = $script:AzureKnwonClients["office"] # set client Id = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  which is office client app

        if ($forcemfa) {
            $AccessTokens = Get-AccessToken -cloud $Cloud -Resource  $resource -RedirectUri $RedirectUri -ClientId $clientId -forcemfa $true -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode -IncludeRefreshToken $true
        } else {
            $AccessTokens = Get-AccessToken -cloud $Cloud -Resource  $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode -IncludeRefreshToken $true
        }


        $AccessToken = Get-AccessTokenWithRefreshToken -cloud $Cloud -Resource $aadiamapi -ClientId $clientId -SaveToCache $SaveToCache -RefreshToken $AccessTokens[1] -TenantId (Read-Accesstoken $AccessTokens[0]).tid

        if(!$SaveToCache)
        {
            return $AccessToken
        }
    }
}

# Gets an access token for MS Commerce
# seems not work
function Get-AccessTokenForMSCommerce   
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for MS Commerce

    .DESCRIPTION
    Gets OAuth Access Token for MS Commerce

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForMSCommerce
    
    .Example
    PS C:\>Get-AADIntAccessTokenForMSCommerce -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $clientId = $script:AzureKnwonClients['mscommerce'] # get ms commerce api client id 
        $resource = $script:AzureKnwonClients["m365licent"] # set resource as M365 License Manager
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets an access token for MS Partner
# Sep 22nd 2021
function Get-AccessTokenForMSPartner
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for MS Partner

    .DESCRIPTION
    Gets OAuth Access Token for MS Partner

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForMSCommerce
    
    .Example
    PS C:\>Get-AADIntAccessTokenForMSPartner -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $clientId = $script:AzureKnwonClients["office"] # set client as office client id 
        $resource = $script:AzureKnwonClients["mspartner"] # set resource as mspartner API
        # The correct client id would be 4990cffe-04e8-4e8b-808a-1175604b879f but that flow doesn't work :(
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets an access token for admin.microsoft.com
# Sep 22nd 2021
function Get-AccessTokenForAdmin
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for admin.microsoft.com

    .DESCRIPTION
    Gets OAuth Access Token for admin.microsoft.com

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForAdmin
    
    .Example
    PS C:\>Get-AADIntAccessTokenForAdmin -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]['admin'] # get M365 admin portal resource
        $clientId = $script:AzureKnwonClients["office"] # set client Id = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  which is office client app
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientid -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}

# Gets an access token for onenote.com
# Feb 2nd 2022
function Get-AccessTokenForOneNote
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for onenote.com

    .DESCRIPTION
    Gets OAuth Access Token for onenote.com

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AADIntAccessTokenForAdmin
    
    .Example
    PS C:\>Get-AADIntAccessTokenForAdmin -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]['onenote'] # get onenote resource
        $clientId = $script:AzureKnwonClients["teams_client"] # set client Id = "1fec8e78-bce4-4aaf-ab1b-5451cc387264" which is teams client app        
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}



# Gets an access token for Microsoft Information Protection SDK
function Get-AccessTokenForMip
{
<#
    .SYNOPSIS
    Gets OAuth Access Token for onenote.com

    .DESCRIPTION
    Gets OAuth Access Token for onenote.com

    .Parameter Credentials
    Credentials of the user.

    .Parameter PRT
    PRT token of the user.

    .Parameter SAML
    SAML token of the user. 

    .Parameter UserPrincipalName
    UserPrincipalName of the user of Kerberos token

    .Parameter KerberosTicket
    Kerberos token of the user. 
    
    .Parameter UseDeviceCode
    Use device code flow.
    
    .Example
    Get-AccessTokenForMip
    
    .Example
    PS C:\>Get-AccessTokenForMip -SaveToCache
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Credentials',Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$True)]
        [String]$PRTToken,
        [Parameter(ParameterSetName='SAML',Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$KerberosTicket,
        [Parameter(ParameterSetName='Kerberos',Mandatory=$True)]
        [String]$Domain,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(ParameterSetName='DeviceCode',Mandatory=$True)]
        [switch]$UseDeviceCode,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $resource = $script:AzureResources[$Cloud]['mip'] # get onenote resource
        $clientId = $script:AzureKnwonClients["graph_api"] # set client Id = "1fec8e78-bce4-4aaf-ab1b-5451cc387264" which is teams client app        
        Get-AccessToken -cloud $Cloud -Resource $resource -RedirectUri $RedirectUri -ClientId $clientId -KerberosTicket $KerberosTicket -Domain $Domain -SAMLToken $SAMLToken -Credentials $Credentials -SaveToCache $SaveToCache -Tenant $Tenant -PRTToken $PRTToken -UseDeviceCode $UseDeviceCode
    }
}


# Functions to get OIDC user info based on Id token
function Get-UserInfo
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$true)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$true)]
        [String]$ClientId,
        [Parameter(Mandatory=$true)]
        [String]$clientSecret,
        [Parameter(Mandatory=$false)]
        [String]$scope="email", # add email scope by default 
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        if([String]::IsNullOrEmpty($Tenant))        
        {
            $Tenant = "common"
        }

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        # get specific access token for userinfo endpoint            
        $oidcuserinfotoken = Get-AccessToken -cloud $Cloud -ClientId $ClientId -clientSecret $clientSecret -Tenant $Tenant -RedirectUri $RedirectUri -scope $scope
        

        
        if(![string]::IsNullOrEmpty($oidcuserinfotoken)) {
            $headers = @{
                "Authorization" = "Bearer $oidcuserinfotoken"
            }

            # get response of userinfo endpoint

            $url = "$aadloginuri/common/openid/userinfo"
            Write-Verbose "get userinfo from endpoint: $url"
            $jsonResponse=Invoke-RestMethod -UseBasicParsing -Uri $url -Method GET -Headers $headers

            $jsonResponse

        } else {
            write-error "no valid access token for userinfo returned"
        }
    }
}


# Prompts for credentials and gets the access token
# Supports MFA, federation, etc.
function Get-Idtoken
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$False)]
        [bool]$ForceMFA=$false,
        [Parameter(Mandatory=$true)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$true)]
        [String]$ClientId,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [ValidateSet("id_token", "token")]
        [String]$tokentype="id_token",
        [Parameter(Mandatory=$false)]
        [String]$response_mode = "fragment", # use fragment as the response mode. 
        [Parameter(Mandatory=$false)]
        [String]$state = "1234",
        [Parameter(Mandatory=$false)]
        [String]$none = "56789",
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
        # Check the tenant
        if([String]::IsNullOrEmpty($Tenant))        
        {
            $Tenant = "common"
        }

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        $mdm = $script:AzureResources[$Cloud]['mdm']
        
        # Azure AD Join
        if($ClientId -eq "29d9ed98-a469-4536-ade2-f981bc1d605e" -and $Resource -ne $mdm) 
        {
                $RedirectUri="ms-aadj-redir://auth/drs"
        }

        # add scope
        if([String]::IsNullOrEmpty($scope))        
        {
            $scope = "openid profile"
        } else {
            $scope = "openid profile $scope"
        }
        $encodescope =  [System.Web.HttpUtility]::UrlEncode($scope)

        # Set variables
        $auth_redirect= $RedirectUri
        $client_id=     $ClientId # Usually should be graph_api

        $auth_redirect=[System.Web.HttpUtility]::UrlEncode($auth_redirect)
                                
        # Create the url
        $url="$aadloginuri/$Tenant/oauth2/v2.0/authorize?client_id=$client_id&response_type=$tokentype&redirect_uri=$auth_redirect&scope=$encodescope&reponse_mode=$response_mode&prompt=login&state=$state&nonce=$none"
       

       
        if($ForceMFA)
        {
            $url+="&amr_values=mfa"
        }

        write-verbose "oauth Url: $url"
 

        # Create the form and get output     
        $form = Create-LoginForm -Url $url -auth_redirect $RedirectUri

        

        # Show the form and wait for the return value
        if($form.ShowDialog() -ne "OK") {
             # Dispose the control
             $form.Controls[0].Dispose()
             Write-Verbose "Login cancelled"
             return $null
        }
  
        # get the output and extract the return messages
        $response = $form.Controls[0].url
        if ($response_mode -eq "fragment") {
            $queryOutput = [Web.HttpUtility]::ParseQueryString($response.Fragment.TrimStart("#"))
        } else {
            $queryOutput = [Web.HttpUtility]::ParseQueryString($response.Query)
        }

          
        # return null if the output contains error
        if(![string]::IsNullOrEmpty($queryOutput["error"])){
              Write-Error $queryOutput["error"]
              Write-Error $queryOutput["error_uri"]
              Write-Error $queryOutput["error_description"]     
  
              return $null
             
        }

        $tokens = @()
        if (!$queryOutput["id_token"] -and !$queryOutput["access_token"])  {
            Write-Verbose "no any of $tokentype detected from response"
            $form.Controls[0].Dispose()
            return $null

        } else {

            if ($queryOutput["id_token"] ) {

                Write-Verbose "id_token detected from response: "
                $tokens += $queryOutput["id_token"]

            }

            if ($queryOutput["access_token"] ) {

                Write-Verbose "access_token detected from response"
                $tokens += $queryOutput["access_token"]

            }

            $form.Controls[0].Dispose()
            return $tokens

        }
    }
}


# Gets the access token for provisioning API and stores to cache
# Refactored Jun 8th 2020
function Get-AccessToken
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(ParameterSetName='PRT',Mandatory=$False)]
        [String]$PRTToken,
        [Parameter(Mandatory=$False)]
        [String]$SAMLToken,
        [Parameter(Mandatory=$false)]
        [String]$Resource,
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$False)]
        [String]$clientSecret,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$False)]
        [String]$KerberosTicket,
        [Parameter(Mandatory=$False)]
        [String]$Domain,
        [Parameter(Mandatory=$False)]
        [bool]$SaveToCache,
        [Parameter(Mandatory=$False)]
        [bool]$IncludeRefreshToken=$false,
        [Parameter(Mandatory=$False)]
        [bool]$ForceMFA=$false,
        [Parameter(Mandatory=$False)]
        [bool]$UseDeviceCode=$false,
        [Parameter(Mandatory=$False)]
        [string]$BPRT,
        [Parameter(Mandatory=$False)]
        [string]$scope,
        [Parameter(Mandatory=$False)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory=$False)]
        [string]$PfxFileName,
        [Parameter(Mandatory=$False)]
        [securestring]$PfxPassword,
        [Parameter(Mandatory=$False)]
        [string]$TransportKeyFileName,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Begin
    {
        # List of clients requiring the same client id
        $requireClientId=@(
            "cb1056e2-e479-49de-ae31-7812af012ed8" # Pass-through authentication
            "c44b4083-3bb0-49c1-b47d-974e53cbdf3c" # Azure Admin web ui
            "1fec8e78-bce4-4aaf-ab1b-5451cc387264" # Teams
            "d3590ed6-52b3-4102-aeff-aad2292ab01c" # Office, ref. https://docs.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2
            "a0c73c16-a7e3-4564-9a95-2bdf47383716" # EXO Remote PowerShell
            "389b1b32-b5d5-43b2-bddc-84ce938d6737" # Office Management API Editor https://manage.office.com
            "ab9b8c07-8f02-4f72-87fa-80105867a763" # OneDrive Sync Engine
            "9bc3ab49-b65d-410a-85ad-de819febfddc" # SPO
            "29d9ed98-a469-4536-ade2-f981bc1d605e" # MDM
            "0c1307d4-29d6-4389-a11c-5cbe7f65d7fa" # Azure Android App
            "6c7e8096-f593-4d72-807f-a5f86dcc9c77" # MAM
            "4813382a-8fa7-425e-ab75-3b753aab3abb" # Microsoft authenticator
            "8c59ead7-d703-4a27-9e55-c96a0054c8d2"
            "c7d28c4f-0d2c-49d6-a88d-a275cc5473c7" # https://www.microsoftazuresponsorships.com/
            "04b07795-8ddb-461a-bbee-02f9e1bf7b46" # Azure CLI
            "ecd6b820-32c2-49b6-98a6-444530e5a77a" # Edge
            "1950a258-227b-4e31-a9cf-717495945fc2" # Microsoft Azure PowerShell
        )
    }
    Process
    {
        
        $aadgraph = $script:AzureResources[$Cloud]['aad_graph_api']
        $devicemanagementsvc = $script:AzureResources[$Cloud]['devicemanagementsvc']
        $mdm = $script:AzureResources[$Cloud]['mdm']

        # fullfil redirect Uri if not provided to generate access token
        if([string]::IsNullOrEmpty($RedirectUri))
        {
            $RedirectUri = Get-AuthRedirectUrl -ClientId $ClientId -Resource $Resource
        }


        if(![String]::IsNullOrEmpty($KerberosTicket)) # Check if we got the kerberos token
        {
            # Get token using the kerberos token
            $OAuthInfo = Get-AccessTokenWithKerberosTicket -KerberosTicket $KerberosTicket -Domain $Domain -Resource $Resource -ClientId $ClientId
            $access_token = $OAuthInfo.access_token
        }
        elseif(![String]::IsNullOrEmpty($PRTToken)) # Check if we got a PRT token
        {

            Write-Error "PRT token accessing is not implemented due to security considering"
            return $NULL

            # Get token using the PRT token
            #$OAuthInfo = Get-AccessTokenWithPRT -Cookie $PRTToken -Resource $Resource -ClientId $ClientId
            #$access_token = $OAuthInfo.access_token
        }
        elseif($UseDeviceCode) # Check if we want to use device code flow
        {
            # Get token using device code
            $OAuthInfo = Get-AccessTokenUsingDeviceCode -cloud $Cloud -Resource $Resource -ClientId $ClientId -Tenant $Tenant
            $access_token = $OAuthInfo.access_token
        }
        elseif(![String]::IsNullOrEmpty($BPRT)) # Check if we got a BPRT
        {
            # Get token using BPRT
            $OAuthInfo = @{
                "refresh_token" = $BPRT
                "access_token"  = Get-AccessTokenWithRefreshToken -cloud $Cloud -Resource "urn:ms-drs:$devicemanagementsvc" -ClientId "b90d5b8f-5503-4153-b545-b31cecfaece2" -TenantId "Common" -RefreshToken $BPRT
                }
            $access_token = $OAuthInfo.access_token
        }
        else
        {
            
            # Check if we got credentials
            if([string]::IsNullOrEmpty($Credentials) -and [string]::IsNullOrEmpty($SAMLToken))
            {
               
               
                # No credentials given, so prompt for credentials
                #if(  $ClientId -eq "d3590ed6-52b3-4102-aeff-aad2292ab01c" <# Office #> -or 
                #     $ClientId -eq "a0c73c16-a7e3-4564-9a95-2bdf47383716" <# EXO #>    -or 
                #    ($ClientId -eq "29d9ed98-a469-4536-ade2-f981bc1d605e" -and $Resource -eq $mdm) <# MDM #> -or
                #    $ClientId -eq "1b730954-1685-4b74-9bfd-dac224a7b894" <# graph API #>
                #)  
                #{

                $OAuthInfo = Prompt-Credentials -cloud $Cloud -Resource $Resource -ClientId $ClientId -clientSecret $clientSecret -Tenant $Tenant -ForceMFA $ForceMFA -redirecturi $RedirectUri -scope $scope
                
            }
            else
            {
                # Get OAuth info for user
                if(![string]::IsNullOrEmpty($SAMLToken))
                {
                    $OAuthInfo = Get-OAuthInfoUsingSAML -SAMLToken $SAMLToken -ClientId $ClientId -Resource $aadgraph 
                }
                else
                {

                  # call get oauth if the request contains a user name/password credential
                   if ($Credentials.username -like "*@*") {
                        if($requireClientId -contains $ClientId)
                        {
                            # Requires same clientId
                            $OAuthInfo = Get-OAuthInfo -Credentials $Credentials -ClientId $ClientId -Resource $aadgraph 
                        }
                        else
                        {
                            # "Normal" flow
                            $OAuthInfo = Get-OAuthInfo -Credentials $Credentials -Resource $aadgraph 
                        }
                    # call client crentail auth flow
                    } else {
                        $client_token= Get-AccessTokenwithclientcredentail -Credentials $Credentials -Resource $Resource -Tenant $tenant

                        $OAuthInfo = @{
                            "refresh_token" = ""
                            "access_token"  = $client_token.access_token
                        }
                    }
                }
            }

            if([String]::IsNullOrEmpty($OAuthInfo))
            {
                throw "Could not get OAuthInfo!"
            }
            
            # We need to get access token using the refresh token

            # Get the access token from response
            # $access_token = Get-AccessTokenWithRefreshToken -cloud $Cloud -Resource $Resource -ClientId $ClientId -TenantId $tenant_id -RefreshToken $RefreshToken -SaveToCache $SaveToCache
            
        }

        $refresh_token = $OAuthInfo.refresh_token
        $access_token = $OAuthInfo.access_token

        # Check whether we want to get the deviceid and (possibly) mfa in mra claim
        if(($Certificate -ne $null -and [string]::IsNullOrEmpty($PfxFileName)) -or ($Certificate -eq $null -and [string]::IsNullOrEmpty($PfxFileName) -eq $false))
        {
            try
            {
                Write-Verbose "Trying to get new tokens with deviceid claim."
                $deviceTokens = Set-AccessTokenDeviceAuth -AccessToken $access_token -RefreshToken $refresh_token -Certificate $Certificate -PfxFileName $PfxFileName -PfxPassword $PfxPassword -BPRT $([string]::IsNullOrEmpty($BPRT) -eq $False) -TransportKeyFileName $TransportKeyFileName
            }
            catch
            {
                Write-Warning "Could not get tokens with deviceid claim: $($_.Exception.Message)"
            }

            if($deviceTokens.access_token)
            {
                $access_token =  $deviceTokens.access_token
                $refresh_token = $deviceTokens.refresh_token

                $claims = Read-Accesstoken $access_token
                Write-Verbose "Tokens updated with deviceid: ""$($claims.deviceid)"" and amr: ""$($claims.amr)"""
            }
        }

        
        # Save the refresh token and other variables

        if($SaveToCache -and $OAuthInfo -ne $null -and $access_token -ne $null)
        {

            Write-Verbose "ACCESS TOKEN: SAVE TO CACHE"
            $script:tokens["$ClientId-$Resource"] =          $access_token
            $script:refresh_tokens["$ClientId-$Resource"] =  $refresh_token
        }

        # Return
        if([string]::IsNullOrEmpty($access_token))
        {
            Throw "Could not get Access Token!"
        }

        # Don't print out token if saved to cache!
        if($SaveToCache)
        {
            $pat = Read-Accesstoken -AccessToken $access_token
            $attributes=[ordered]@{
                "Tenant" =   $pat.tid
                "User" =     $pat.unique_name
                "Resource" = $Resource
                "Client" =   $ClientID
            }
            Write-Host "AccessToken saved to cache."
            return New-Object psobject -Property $attributes
        }
        else
        {
            if($IncludeRefreshToken) # Include refreshtoken
            {
                return @($access_token,$refresh_token)
            }
            else
            {
                return $access_token
            }
        }
    }
}

# Gets the access token using a refresh token
# Jun 8th 2020
function Get-AccessTokenWithRefreshToken
{
    [cmdletbinding()]
    Param(
        [String]$Resource,
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$True)]
        [String]$TenantId,
        [Parameter(Mandatory=$True)]
        [String]$RefreshToken,
        [Parameter(Mandatory=$False)]
        [bool]$SaveToCache = $false,
        [Parameter(Mandatory=$False)]
        [bool]$IncludeRefreshToken = $false,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        # Set the body for API call
        $body = @{
            "resource"=      $Resource
            "client_id"=     $ClientId
            "grant_type"=    "refresh_token"
            "refresh_token"= $RefreshToken
            "scope"=         "openid"
        }

        $aadlogin = $script:AzureResources[$Cloud]['aad_login']
        $aadlogincommon = $script:AzureResources[$Cloud]['aad_login_common']


        if($ClientId -eq "ab9b8c07-8f02-4f72-87fa-80105867a763") # OneDrive Sync Engine
        {
            $url = "$aadlogincommon/common/oauth2/token"
        }
        else
        {
            $url = "$aadlogin/$TenantId/oauth2/token"
        }

        # Debug
        Write-verbose "ACCESS TOKEN BODY: $($body | Out-String)"
        
        # Set the content type and call the API
        # try the refresh token and if the returned error is AADSTS50076, will force a new token with MFA enabled
        $contentType="application/x-www-form-urlencoded"
        try {
            $response=Invoke-RestMethod -UseBasicParsing -Uri $url -ContentType $contentType -Method POST -Body $body            
        }
        catch {
            $errormessage = $Error[0].ErrorDetails.Message | convertfrom-json

            # raise MFA request if the returned error code is 50076/50078/50079
            if ($errormessage.error_codes -contains "50076" -or $errormessage.error_codes -contains "50078" -or $errormessage.error_codes -contains "50079" ) {
                Write-Error "request access token need MFA. Please add force -ForceMFA at the end of your command and try again" 
            } else {
                Write-Error  $errormessage.error_description 
            }
        }
 

        # Debug
        Write-verbose "ACCESS TOKEN RESPONSE: $response"

        # Save the tokens to cache
        if($SaveToCache)
        {
            Write-Verbose "ACCESS TOKEN: SAVE TO CACHE"
            $Script:tokens["$ClientId-$Resource"] =         $response.access_token
            $Script:refresh_tokens["$ClientId-$Resource"] = $response.refresh_token
        }

        # Return
        if($IncludeRefreshToken)
        {
            return @($response.access_token, $response.refresh_token)
        }
        else
        {
            return $response.access_token    
        }
    }
}

# Gets access token using device code flow
function Get-AccessTokenUsingDeviceCode
{
    [cmdletbinding()]
    Param(
        
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$false)]
        [String]$resource,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

    # if resource is not defined, user the aad graph api as the default resource
       if (!$resource) {  
           $resource = $script:AzureResources[$Cloud]['aad_graph_api']
        }
        $aadlogin = $script:AzureResources[$Cloud]['aad_login']

        # Check the tenant
        if([string]::IsNullOrEmpty($Tenant))
        {
            $Tenant="Common"
        }

        # Create a body for the first request
        $body=@{
            "client_id" = $ClientId
            "resource" =  $Resource
        }

        # Invoke the request to get device and user codes
        $authResponse = Invoke-RestMethod -UseBasicParsing -Method Post -Uri " $aadlogin/$tenant/oauth2/devicecode?api-version=1.0" -Body $body

        Write-Host $authResponse.message

        $continue = $true
        $interval = $authResponse.interval
        $expires =  $authResponse.expires_in

        # Create body for authentication subsequent requests
        $body=@{
            "client_id" =  $ClientId
            "grant_type" = "urn:ietf:params:oauth:grant-type:device_code"
            "device_code" = $authResponse.device_code           
        }

        Write-verbose "ACCESS TOKEN BODY: $($body | Out-String)"
        Write-Verbose "try loop device code with a interval: $interval seconds and will be expired after $expires seconds"
        $total = 0

        # Loop while pending or until timeout exceeded
        while($continue)
        {
            Start-Sleep -Seconds $interval
            $total += $interval

            Write-verbose "waiting for $total seconds"

            if($total -gt $expires)
            {
                Write-Error "Timeout occurred"
                return
            }
                        
            # Try to get the response. Will give 40x while pending so we need to try&catch
            try
            {
                $response = Invoke-RestMethod -UseBasicParsing -Method POST -Uri "$aadlogin/$Tenant/oauth2/v2.0/token?api-version=1.0 " -Body $body -ErrorAction SilentlyContinue
            }
            catch
            {
                # This normal flow, always returns 40x unless successful
                $details=$_.ErrorDetails.Message | ConvertFrom-Json
                $continue = $details.error -eq "authorization_pending"
                Write-Verbose $details.error
                Write-Host "." -NoNewline

                if(!$continue)
                {
                    # Not authorization_pending so this is a real error :(
                    Write-Error $details.error_description
                    return
                } else {
                    continue
                }
            }

            # If we got response, all okay!
            if($response)
            {
                # Debug
                Write-verbose "ACCESS TOKEN RESPONSE: $response"

                return $response
            }
        }

    }
}

# get access token based on client credential
function Get-AccessTokenwithclientcredentail
{
    [cmdletbinding()]
    Param(
        
        [Parameter(Mandatory=$false)]
        [String]$resource,
        [Parameter(Mandatory=$true)]
        [String]$Tenant,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

    # if resource is not defined, user the aad graph api as the default resource
       if (!$resource) {  
           $resource = $script:AzureResources[$Cloud]["ms_graph_api"]
        }
        $aadlogin = $script:AzureResources[$Cloud]['aad_login']

        # define a scope with default of request resource
        $scope = $resource+"/.default"
        
        # Create a body for the first request
        $body=@{
            "client_id" = $credentials.GetNetworkCredential().username
            "grant_type" = "client_credentials"
            “client_secret" = $credentials.GetNetworkCredential().password
            "scope"=  $scope
        }

        Write-Verbose "ACCESS TOKEN BODY: $($body | Out-String)"

        # Invoke the request to get device and user codes
        $contentType = "application/x-www-form-urlencoded"
        $Response = Invoke-RestMethod -UseBasicParsing -Method Post -Uri " $aadlogin/$tenant/oauth2/v2.0/token"  -ContentType $contentType -Body $body

        # If we got response, all okay!
        if($response)
        {
            # Debug
            Write-verbose "ACCESS TOKEN RESPONSE: $response"
            return $response
        }

    }
}


# get access token with a shared app token (OBO behavier)
function Get-AccessTokenwithobo
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Mandatory=$true)]
        [String]$Tenant,        
        [Parameter(Mandatory=$true)]
        [String]$token,        
        [String]$scope, 
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

    # if resource is not defined, user the aad graph api as the default resource
       if (!$resource) {  
           $resource = $script:AzureResources[$Cloud]["ms_graph_api"]
        }
        $aadlogin = $script:AzureResources[$Cloud]['aad_login']


        if([String]::IsNullOrEmpty($scope))        
        {
            $scope = "openid profile"
        } 
        
        # Create a body for the first request
        $body=@{
            "client_id" = $credentials.GetNetworkCredential().username
            "grant_type" = "urn:ietf:params:oauth:grant-type:jwt-bearer"
            “client_secret" = $credentials.GetNetworkCredential().password
            "assertion" = $token
            "scope"=  $scope
            "requested_token_use"="on_behalf_of"
        }

        Write-Verbose "ACCESS TOKEN BODY: $($body | Out-String)"

        # Invoke the request to get device and user codes
        $contentType = "application/x-www-form-urlencoded"
        $Response = Invoke-RestMethod -UseBasicParsing -Method Post -Uri " $aadlogin/$tenant/oauth2/v2.0/token"  -ContentType $contentType -Body $body

        # If we got response, all okay!
        if($response)
        {
            # Debug
            Write-verbose "ACCESS TOKEN RESPONSE: $response"
            return $response
        }

    }
}

# Gets the access token using an authorization code
function Get-AccessTokenWithAuthorizationCode
{
    [cmdletbinding()]
    Param(
        [String]$Resource,
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$True)]
        [String]$TenantId,
        [Parameter(Mandatory=$True)]
        [String]$AuthorizationCode,
        [Parameter(Mandatory=$False)]
        [bool]$SaveToCache = $false,
        [Parameter(Mandatory=$False)]
        [bool]$IncludeRefreshToken = $false,
        [Parameter(Mandatory=$False)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$False)]
        [String]$CodeVerifier,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $aadlogin = $script:AzureResources[$Cloud]['aad_login']
        $aadlogincommon = $script:AzureResources[$Cloud]['aad_login_common']

        $headers = @{
        }

        # Set the body for API call
        $body = @{
            "resource"=      $Resource
            "client_id"=     $ClientId
            "grant_type"=    "authorization_code"
            "code"=          $AuthorizationCode
            "scope"=         "openid profile email"
        }
        if(![string]::IsNullOrEmpty($RedirectUri))
        {
            $body["redirect_uri"] = $RedirectUri
            $headers["Origin"] = $RedirectUri
        }

        if(![string]::IsNullOrEmpty($CodeVerifier))
        {
            $body["code_verifier"] = $CodeVerifier
            $body["code_challenge_method"] = "S256"
        }

        if($ClientId -eq "ab9b8c07-8f02-4f72-87fa-80105867a763") # OneDrive Sync Engine
        {
            $url = "$aadlogincommon/common/oauth2/token"
        }
        else
        {
            $url = "$aadlogin/$TenantId/oauth2/token"
        }
        
        # Debug
        write-verbose "ACCESS TOKEN BODY: $($body | Out-String)"
        
        # Set the content type and call the API
        $contentType = "application/x-www-form-urlencoded"
        $response =    Invoke-RestMethod -UseBasicParsing -Uri $url -ContentType $contentType -Method POST -Body $body -Headers $headers

        # Debug
        write-verbose "ACCESS TOKEN RESPONSE: $response"

        # Save the tokens to cache
        if($SaveToCache)
        {
            Write-Verbose "ACCESS TOKEN: SAVE TO CACHE"
            $Script:tokens["$ClientId-$Resource"] =         $response.access_token
            $Script:refresh_tokens["$ClientId-$Resource"] = $response.refresh_token
        }

        # Return
        return $response.access_token    
    }
}

# Gets the access token using device SAML token
# Feb 18th 2021
function Get-AccessTokenWithDeviceSAML
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$SAML,
        [Parameter(Mandatory=$False)]
        [bool]$SaveToCache,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $devicemanagementsvc = $script:AzureResources[$Cloud]['devicemanagementsvc']
        $resource = "urn:ms-drs:$devicemanagementsvc"
        $aadlogin = $script:AzureResources[$Cloud]['aad_login']

        $headers = @{
        }

         
        $ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894" #"dd762716-544d-4aeb-a526-687b73838a22"

        # Set the body for API call
        $body = @{
            "resource"=      $Resource
            "client_id"=     $ClientId
            "grant_type"=    "urn:ietf:params:oauth:grant-type:saml1_1-bearer"
            "assertion"=     Convert-TextToB64 -Text $SAML
            "scope"=         "openid"
        }
        
        # Debug
        write-verbose "ACCESS TOKEN BODY: $($body | Out-String)"
        
        # Set the content type and call the API
        $contentType = "application/x-www-form-urlencoded"
        $response =    Invoke-RestMethod -UseBasicParsing -Uri "$aadlogin/common/oauth2/token" -ContentType $contentType -Method POST -Body $body -Headers $headers

        # Debug
        write-verbose "ACCESS TOKEN RESPONSE: $response"

        # Save the tokens to cache
        if($SaveToCache)
        {
            Write-Verbose "ACCESS TOKEN: SAVE TO CACHE"
            $Script:tokens["$ClientId-$Resource"] =         $response.access_token
            $Script:refresh_tokens["$ClientId-$Resource"] = $response.refresh_token
        }
        else
        {
            # Return
            return $response.access_token    
        }
    }
}

# Logins to SharePoint Online and returns an IdentityToken
# TODO: Research whether can be used to get access_token to AADGraph
# TODO: Add support for Google?
# FIX: Web control stays logged in - clear cookies somehow?
# Aug 10th 2018
function Get-IdentityTokenByLiveId
{
<#
    .SYNOPSIS
    Gets identity_token for SharePoint Online for External user

    .DESCRIPTION
    Gets identity_token for SharePoint Online for External user using LiveId.

    .Parameter Tenant
    The tenant name to login in to WITHOUT .sharepoint.com part
    
    .Example
    PS C:\>$id_token=Get-AADIntIdentityTokenByLiveId -Tenant mytenant
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Tenant,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

    

        # Set variables
        $aadlogin = $script:AzureResources[$Cloud]['aad_login']
        $auth_redirect="$aadlogin/common/federation/oauth2" # When to close the form
        $sharepoint = $script:AzureResources[$Cloud]['sharepoint']
        $url="https://$Tenant.$sharepoint"

        # Create the form
        $form=Create-LoginForm -Url $url -auth_redirect $auth_redirect

        # Show the form and wait for the return value
        if($form.ShowDialog() -ne "OK") {
            Write-Verbose "Login cancelled"
            return $null
        }

        $web=$form.Controls[0]

        $code=$web.Document.All["code"].GetAttribute("value")
        $id_token=$web.Document.All["id_token"].GetAttribute("value")
        $session_state=$web.Document.All["session_state"].GetAttribute("value")

        return Read-Accesstoken($id_token)
    }
}

# Tries to generate access token using cached AADGraph token
# Jun 15th 2020
function Get-AccessTokenUsingAADGraph
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Resource,
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [switch]$SaveToCache,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        
        $aadgraph = $script:AzureResources[$Cloud]['aad_graph_api']
        
        # Try to get AAD Graph access token from the cache
        $AccessToken = Get-AccessTokenFromCache -AccessToken $null -Resource $aadgraph -ClientId "1b730954-1685-4b74-9bfd-dac224a7b894"

        # Get the tenant id
        $tenant = (Read-Accesstoken -AccessToken $AccessToken).tid
                
        # Get the refreshtoken
        $refresh_token=$script:refresh_tokens["1b730954-1685-4b74-9bfd-dac224a7b894-$aadgraph"]

        if([string]::IsNullOrEmpty($refresh_token))
        {
            Throw "No refreshtoken found! Use Get-AADIntAccessTokenForAADGraph with -SaveToCache switch."
        }

        # Create a new AccessToken for Azure AD management portal API
        $AccessToken = Get-AccessTokenWithRefreshToken -cloud $Cloud -Resource $Resource -ClientId $ClientId -TenantId $tenant -RefreshToken $refresh_token -SaveToCache $SaveToCache

        # Return
        $AccessToken
    }
}

# Apr 22th 2022
function Unprotect-EstsAuthPersistentCookie
{
<#
    .SYNOPSIS
    Decrypts and dumps users stored in ESTSAUTHPERSISTENT 

    .DESCRIPTION
    Decrypts and dumps users stored in ESTSAUTHPERSISTENT using login.microsoftonline.com/forgetUser

    .Parameter Cookie
    Value of ESTSAUTHPERSISTENT cookie
    
    .Example
    PS C:\>Unprotect-AADIntEstsAuthPersistentCookie -Cookie 0.ARMAqlCH3MZuvUCNgTAd4B7IRffhvoluXopNnz3s1gEl...

    name       : Some User
    login      : user@company.com
    imageAAD   : work_account.png
    imageMSA   : personal_account.png
    isLive     : False
    isGuest    : False
    link       : user@company.com
    authUrl    : 
    isSigned   : True
    sessionID  : 1fb5e6b3-09a4-4ceb-bcad-3d6d0ee89bf7
    domainHint : 
    isWindows  : False

    name       : Another User
    login      : user2@company.com
    imageAAD   : work_account.png
    imageMSA   : personal_account.png
    isLive     : False
    isGuest    : False
    link       : user2@company.com
    authUrl    : 
    isSigned   : False
    sessionID  : 1fb5e6b3-09a4-4ceb-bcad-3d6d0ee89bf7
    domainHint : 
    isWindows  : False
#>

    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline)]
        [String]$Cookie,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $aadlogin = $script:AzureResources[$Cloud]['aad_login']
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $aadloginsuffix = $aadlogin.split("//")[1].toString()
        
        $session.Cookies.Add((New-Object System.Net.Cookie("ESTSAUTHPERSISTENT", $Cookie, "/", ".$aadloginsuffix")))
        Invoke-RestMethod -UseBasicParsing -Uri "$aadlogin/forgetuser?sessionid=$((New-Guid).toString())" -WebSession $session
    }
}
