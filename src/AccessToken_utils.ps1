# This script contains functions for handling access tokens
# and some utility functions

# VARIABLES

# Unix epoch time (1.1.1970)
$epoch = Get-Date -Day 1 -Month 1 -Year 1970 -Hour 0 -Minute 0 -Second 0 -Millisecond 0

$DefaultAzureCloud = "AzureChina"
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms


# Clear the Forms.WebBrowser data
# 
$source=@"
[DllImport("wininet.dll", SetLastError = true)]
public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);

[DllImport("wininet.dll", SetLastError = true)]
public static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, System.Text.StringBuilder pchCookieData, ref uint pcchCookieData, int dwFlags, IntPtr lpReserved);
"@

##Create type from source
$WebBrowser = Add-Type -memberDefinition $source -passthru -name WebBrowser -ErrorAction SilentlyContinue
$INTERNET_OPTION_END_BROWSER_SESSION = 42
$INTERNET_OPTION_SUPPRESS_BEHAVIOR   = 81
$INTERNET_COOKIE_HTTPONLY            = 0x00002000
$INTERNET_SUPPRESS_COOKIE_PERSIST    = 3

# 

# referred known first-party app https://github.com/Seb8iaan/Microsoft-Owned-Enterprise-Applications/blob/main/Microsoft%20Owned%20Enterprise%20Applications%20Overview.md
$AzureKnwonClients = @{

    "msgraph" =             "00000003-0000-0000-c000-000000000000" # msgraph resource AppID
    "drs" =                 "01cb2876-7ebd-4aa4-9cc9-d28bd4d359a9" # Device Registration Service
    "graph_api"=            "1b730954-1685-4b74-9bfd-dac224a7b894" # MS Graph API
    "aadrm"=                "90f610bf-206d-4950-b61d-37fa6fd1b224" # AADRM
    "aad"=                  "00000002-0000-0000-c000-000000000000" # Azure Active Directory
    "mscommerce" =          "3d5cffa9-04da-4657-8cab-c7f074657cad" # MS Commerce
    "m365licent" =          "aeb86249-8ea3-49e2-900b-54cc8e308f85" # M365 License Manager
    "exo"=                  "a0c73c16-a7e3-4564-9a95-2bdf47383716" # EXO Remote PowerShell
    "skype"=                "d924a533-3729-4708-b3e8-1d2445af35e3" # Skype
    "o365portal"=           "00000006-0000-0ff1-ce00-000000000000" # Office portal
    "o365spo"=              "00000003-0000-0ff1-ce00-000000000000" # SharePoint Online
    "o365exo"=              "00000002-0000-0ff1-ce00-000000000000" # Exchange Online
    "dynamicscrm"=          "00000007-0000-0000-c000-000000000000" # Dynamics CRM
    "o365suiteux"=          "4345a7b9-9a63-4910-a426-35363201d503" # O365 Suite UX
    "aadsync"=              "cb1056e2-e479-49de-ae31-7812af012ed8" # Azure AD Sync
    "aadconnectv2"=         "6eb59a73-39b2-4c23-a70f-e2e3ce8965b1" # AAD Connect v2
    "synccli"=              "1651564e-7ce4-4d99-88be-0a65050d8dc3" # Sync client
    "azureadmin" =          "c44b4083-3bb0-49c1-b47d-974e53cbdf3c" # Azure Admin web ui
    "pta" =                 "cb1056e2-e479-49de-ae31-7812af012ed8" # Pass-through authentication
    "patnerdashboard" =     "4990cffe-04e8-4e8b-808a-1175604b879"  # Partner dashboard (missing on letter?)
    "webshellsuite" =       "89bee1f7-5e6e-4d8a-9f3d-ecd601259da7" # Office365 Shell WCSS-Client
    "teams" =               "1fec8e78-bce4-4aaf-ab1b-5451cc387264" # Teams
    "mspartner" =           "fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd" # Microsoft Partner Center
    "office" =              "d3590ed6-52b3-4102-aeff-aad2292ab01c" # Office, ref. https://docs.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2
    "office_online2" =      "57fb890c-0dab-4253-a5e0-7188c88b2bb4" # SharePoint Online Client
    "office_online" =       "bc59ab01-8403-45c6-8796-ac3ef710b3e3" # Outlook Online Add-in App
    "powerbi_contentpack" = "2a0c3efa-ba54-4e55-bdc0-770f9e39e9ee" # PowerBI content pack
    "aad_account" =         "0000000c-0000-0000-c000-000000000000" # https://account.activedirectory.windowsazure.com
    "sara" =                "d3590ed6-52b3-4102-aeff-aad2292ab01c" # Microsoft Support and Recovery Assistant (SARA)
    "office_mgmt" =         "389b1b32-b5d5-43b2-bddc-84ce938d6737" # Office Management API Editor https://manage.office.com 
    "onedrive" =            "ab9b8c07-8f02-4f72-87fa-80105867a763" # OneDrive Sync Engine
    "adibizaux" =           "74658136-14ec-4630-ad9b-26e160ff0fc6" # Azure portal AAD blade "ADIbizaUX"
    "msmamservice" =        "27922004-5251-4030-b22d-91ecd9a37ea4" # MS MAM Service API
    "teamswebclient" =      "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" # Teams web client
    "azuregraphclientint" = "7492bca1-9461-4d94-8eb8-c17896c61205" # Microsoft Azure Graph Client Library 2.1.9 Internal
    "azure_mgmt" =          "84070985-06ea-473d-82fe-eb82b4011c9d" # Windows Azure Service Management API
    "az" =                  "1950a258-227b-4e31-a9cf-717495945fc2" # AZ PowerShell Module
    "apple" =               "f8d98a96-0999-43f5-8af3-69971c7bb423" # Apple Internet Accounts
    "globaladmin" =         "7f59a773-2eaf-429c-a059-50fc5bb28b44" # https://docs.microsoft.com/en-us/rest/api/authorization/globaladministrator/elevateaccess#code-try-0
    "spo_shell" =           "9bc3ab49-b65d-410a-85ad-de819febfddc" # SPO Management Shell
    "aad_pin" =             "06c6433f-4fb8-4670-b2cd-408938296b8e" # AAD Pin redemption client
    "PIM" =                 "01fc33a7-78ba-4d2f-a4b7-768e336e890e" # Azure PIM: https://api.azrbac.azurepim.identitygovernance.azure.cn
    "mysignins" =           "19db86c3-b2b9-44cc-b339-36da233a3be2" # https://mysignins.microsoft.com
    "office_management" =   "00b41c95-dab0-4487-9791-b9d2c32c80f2" # Office 365 Management (mobile app)
    "azure_broker" =        "29d9ed98-a469-4536-ade2-f981bc1d605e" # Microsoft Authentication Broker (Azure MDM client)
    "WAM" =                 "6f7e0f60-9401-4f5b-98e2-cf15bd5fd5e3" # Microsoft.AAD.BrokerPlugin resource:https://cs.dds.microsoft.com
    "aad_cloud_ap" =        "38aa3b87-a06d-4817-b275-7a316988d93b" # Microsoft AAD Cloud AP
    "android" =             "0c1307d4-29d6-4389-a11c-5cbe7f65d7fa" # Azure Android App
    "intune" =              "6c7e8096-f593-4d72-807f-a5f86dcc9c77" # Intune MAM client resource:https://intunemam.microsoftonline.com
    "authenticator" =       "4813382a-8fa7-425e-ab75-3b753aab3abb" # Authenticator App resource:ff9ebd75-fe62-434a-a6ce-b3f0a8592eaf
    "teams_client" =        "1fec8e78-bce4-4aaf-ab1b-5451cc387264" # Teams client
    "wcd" =                 "de0853a1-ab20-47bd-990b-71ad5077ac7b" # Windows Configuration Designer (WCD)
    "aadj" =                "b90d5b8f-5503-4153-b545-b31cecfaece2" # AADJ CSP
    "EXO_powershell" =      "fb78d390-0c51-40cd-8e17-fdbfab77341b" # Microsoft Exchange REST API Based Powershell
    "aad_registeredapp" =   "18ed3507-a475-4ccb-b669-d66bc9f2a36e" # Microsoft_AAD_RegisteredApps
    "Intune_multi-tenant_management" =                    "3f1abb3f-12cc-42c3-ad06-5b608dc5fb67" # Microsoft Intune multi-tenant management UX extension
    "Entitlement_Management" =    "810dcf14-1858-4bf2-8134-4c369fa3235b" # Azure AD Identity Governance - Entitlement Management
}


# AccessToken resource strings
<#
$resources=@{
    "aad_graph_api"=         "https://graph.windows.net"
    "ms_graph_api"=          "https://graph.microsoft.com"
    "azure_mgmt_api" =       "https://management.azure.com"
    "windows_net_mgmt_api" = "https://management.core.windows.net/"
    "cloudwebappproxy" =     "https://proxy.cloudwebappproxy.net/registerapp"
    "officeapps" =           "https://officeapps.live.com"
    "outlook" =              "https://outlook.office365.com"
    "webshellsuite" =        "https://webshell.suite.office.com"
    "sara" =                 "https://api.diagnostics.office.com"
    "office_mgmt" =          "https://manage.office.com"
    "msmamservice" =         "https://msmamservice.api.application"
    "spacesapi" =            "https://api.spaces.skype.com"
}
#>

# Stored tokens (access & refresh)
$tokens=@{} # hash table
$refresh_tokens=@{} # hash table


## get valid AAD endpoint 
# referred doc: https://docs.microsoft.com/en-us/azure/china/resources-developer-guide
# referred doc: https://docs.microsoft.com/zh-cn/previous-versions/office/office-365-api/api/o365-china-endpoints

$AzureResources = @{
    "AzureChina" = @{
        "suffixes" = @{
            "acrLoginServerEndpoint" = '.azurecr.cn'
       ##    "attestationEndpoint" = '.attest.azure.cn'
       ##    "azureDatalakeAnalyticsCatalogAndJobEndpoint"= '.azuredatalakeanalytics.cn'
       ##     "azureDatalakeStoreFileSystemEndpoint" = '.azuredatalakestore.net'
            "keyvaultDns" = '.vault.azure.cn'
            "mariadbServerEndpoint" = '.mariadb.database.chinacloudapi.cn'
        ##     "mhsmDns" = '.managedhsm.azure.net'
            "mysqlServerEndpoint" = '.mysql.database.chinacloudapi.cn'
            "postgresqlServerEndpoint" = '.postgres.database.chinacloudapi.cn'
            "sqlServerHostname" =  '.database.chinacloudapi.cn'
            "storageEndpoint"  = '.core.chinacloudapi.cn'
            "storageSyncEndpoint"= '.file.core.chinacloudapi.cn'
            "AnalyticsServiceEndpoint" = '.asazure.chinacloudapi.cn'
            "cloudAppEndpoint"= '.chinacloudapp.cn'
            "trafficManagerEndpoint" = '.trafficmanager.cn'
            "hdinsight_suffix" =     ".azurehdinsight.cn"
            "compute"   =            ".chinacloudapp.cn"
            "sharepoint" =           ".sharepoint.cn"
        }
        "ms-pim"  =  "https://api.azrbac.azurepim.identitygovernance.azure.cn"
        "keyvault" =  "https://vault.azure.cn/"
        "storage"  =   "https://storage.azure.com/"
        "Cosmon" = "https://cosmos.azure.cn"
        "powerBI"= "https://analysis.chinacloudapi.cn/powerbi/api/"
        "onenote" =              "https://api.partner.office365.cn"        
        "admin" =                "https://portal.partner.microsoftonline.cn"       
        "portal" =               "https://portal.azure.cn"
        "aad_login" =            "https://login.partner.microsoftonline.cn"
        "aad_login_common" =     "https://login.chinacloudapi.cn"
        "aad_graph_api"=         "https://graph.chinacloudapi.cn"
        "ms_graph_api"=          "https://microsoftgraph.chinacloudapi.cn"
        "azure_mgmt_api" =       "https://management.chinacloudapi.cn/"
        "windows_net_mgmt_api" = "https://management.core.chinacloudapi.cn/"
        "appInsightsResourceId" =  'https://api.applicationinsights.azure.cn'
        "appInsightsTelemetryChannelResourceId" = 'https://dc.applicationinsights.azure.cn/v2/track'
        "batch" =  'https://batch.chinacloudapi.cn/'
        "gallery" = 'https://gallery.chinacloudapi.cn/'
        "logAnalytics" = 'https://api.loganalytics.azure.cn'
        "sqlManagement" = 'https://management.core.chinacloudapi.cn:8443/'
     ##   "synapseAnalytics" =   'https://dev.azuresynapse.net'
        "officeapps" =           "https://officeapps.live.com"
        "outlook" =              "https://partner.outlook.cn"
        "webshellsuite" =        "https://webshell.suite.partner.microsoftonline.cn"
        "sara" =                 "https://api.partner.office365.cn"
        "office_mgmt" =          "https://manage.office365.cn/"
        "sql_database" =         "https://management.database.chinacloudapi.cn"
      ##   "msmamservice" =         "https://msmamservice.api.application"
      ##  "spacesapi" =            "https://api.spaces.skype.com"
      ##  "intune" =               "https://api.manage.microsoftonline.cn/"
      ##  "cloudwebappproxy" =     "https://proxy.cloudwebappproxy.net/registerapp"
        "mdm" =                  "https://enrollment.manage.microsoftonline.cn/"
        "devicemanagementsvc" =  "enterpriseregistration.partner.microsoftonline.cn"
        "mip" =                  "https://syncservice.o365syncservice.com"
        "aad_iam" =              "https://main.iam.ad.ext.azure.cn"
        "o365exo" = "https://ps.compliance.protection.partner.outlook.cn"
    }
    "AzurePublic" = @{
        "suffixes" = @{
            "acrLoginServerEndpoint" = '.azurecr.io'
            "attestationEndpoint" = '.attest.azure.net'
            "azureDatalakeAnalyticsCatalogAndJobEndpoint"= '.azuredatalakeanalytics.net'
            "azureDatalakeStoreFileSystemEndpoint" = '.azuredatalakestore.net'
            "keyvaultDns" = '.vault.azure.net'
            "mariadbServerEndpoint" = '.mariadb.database.azure.com'
            "mhsmDns" = '.managedhsm.azure.net'
            "mysqlServerEndpoint" = '.mysql.database.azure.com'
            "postgresqlServerEndpoint" = '.postgres.database.azure.com'
            "sqlServerHostname" =  '.database.windows.net'
            "storageEndpoint"  = '.core.windows.net'
            "storageSyncEndpoint"= '.afs.azure.net'
            "synapseAnalyticsEndpoint" = '.dev.azuresynapse.net'
            "cloudAppEndpoint"= '.cloudapp.azure.com'
            "trafficManagerEndpoint" = '.trafficmanager.net'
            "hdinsight_suffix" =     ".azurehdinsight.net"
            "compute"   =            ".cloudapp.net"
            "sharepoint" =           ".sharepoint.com"
        }
        "ms-pim"  =  "https://api.azrbac.mspim.azure.com"
        "keyvault" =  "https://vault.azure.com/"
        "storage"  =   "https://storage.azure.com/"
        "Cosmon" = "https://cosmos.azure.com"
        "powerBI" = "https://analysis.windows.net/powerbi/api/"
        "onenote" =              "https://onenote.com"        
        "admin" =                "https://admin.microsoft.com"
        "portal" =               "https://portal.azure.com"
        "aad_login" =            "https://login.microsoftonline.com"
        "aad_login_common" =     "https://login.windows.net"
        "aad_graph_api"=         "https://graph.windows.net"
        "ms_graph_api"=          "https://graph.microsoft.com"
        "azure_mgmt_api" =       "https://management.azure.com"
        "windows_net_mgmt_api" = "https://management.core.windows.net/"
        "cloudwebappproxy" =     "https://proxy.cloudwebappproxy.net/registerapp"
        "appInsightsResourceId" = 'https://api.applicationinsights.io'
        "appInsightsTelemetryChannelResourceId" = 'https://dc.applicationinsights.azure.com/v2/track'
        "attestationResourceId" = 'https://attest.azure.net'
        "batch" =  'https://batch.core.windows.net/'
        "gallery" = 'https://gallery.azure.com/'
        "logAnalytics" = 'https://api.loganalytics.io'
        "sqlManagement" = 'https://management.core.windows.net:8443/'
        "synapseAnalyticsResourceId" =  'https://dev.azuresynapse.net'
        "officeapps" =           "https://officeapps.live.com"
        "outlook" =              "https://outlook.office.com"
        "webshellsuite" =        "https://webshell.suite.office.com"
        "sara" =                 "https://api.diagnostics.office.com"
        "office_mgmt" =          "https://manage.office.com"
        "sql_database" =         "https://management.database.windows.net"
        "msmamservice" =         "https://msmamservice.api.application"
        "spacesapi" =            "https://api.spaces.skype.com"
        "autodiscover"  =        "https://autodiscover-s.outlook.com"
        "intune" =               "https://api.manage.microsoft.com/"
        "mdm" =                  "https://enrollment.manage.microsoft.com/"
        "devicemanagementsvc" =  "enterpriseregistration.windows.net"
        "mip" =                  "https://syncservice.o365syncservice.com"  
        "aad_iam" =              "https://main.iam.ad.ext.azure.com"
        "o365exo" = "https://ps.outlook.com"
    }

    
}

# get known client
function Get-AADknownclient
{
    Param(
        [Parameter(Mandatory=$false)]
        [String]$clientname
    )
    Process
    {
        if([String]::IsNullOrEmpty($clientname)){
           return $AzureKnwonClients
        } else {
            if ($AzureKnwonClients[$clientname]) {

               $redirectUri = Get-AuthRedirectUrl -ClientId $AzureKnwonClients[$clientname]

               return @{
                clientId = $AzureKnwonClients[$clientname]
                clientname = $clientname
                redirectUri =  $redirectUri 
               }

            } else {
                return $null
            }
    
       
        }

    }
}



# get default cloud instance
function Set-DefaultCloud
{
    Param(
    [Parameter(Mandatory=$True)]
    [ValidateSet("AzureChina", "AzurePublic")]
    [String]$cloud
    )
    Process
    {

        $script:DefaultAzureCloud=$Cloud
        write-output "updated default cloud to $Cloud"
    }
}

# set default cloud instance
function Get-DefaultCloud
{
    $Cloud=$script:DefaultAzureCloud
    write-output "Current default cloud is $Cloud"
   
    write-output "Resource Endpoints for current cloud instance: "
    write-output $script:AzureResources[$Cloud]
}



## UTILITY FUNCTIONS FOR API COMMUNICATIONS

# Return user's login information
function Get-LoginInformation
{
<#
    .SYNOPSIS
    Returns authentication information of the given user or domain

    .DESCRIPTION
    Returns authentication of the given user or domain

    .Example
    Get-AADIntLoginInformation -Domain outlook.com

    Tenant Banner Logo                   : 
    Authentication Url                   : https://login.live.com/login.srf?username=nn%40outlook.com&wa=wsignin1.0&wtrealm=urn%3afederation%3aMicrosoftOnline&wctx=
    Pref Credential                      : 6
    Federation Protocol                  : WSTrust
    Throttle Status                      : 0
    Cloud Instance                       : microsoftonline.com
    Federation Brand Name                : MSA Realms
    Domain Name                          : live.com
    Federation Metadata Url              : https://nexus.passport.com/FederationMetadata/2007-06/FederationMetadata.xml
    Tenant Banner Illustration           : 
    Consumer Domain                      : True
    State                                : 3
    Federation Active Authentication Url : https://login.live.com/rst2.srf
    User State                           : 2
    Account Type                         : Federated
    Tenant Locale                        : 
    Domain Type                          : 2
    Exists                               : 5
    Has Password                         : True
    Cloud Instance audience urn          : urn:federation:MicrosoftOnline
    Federation Global Version            : -1

    .Example
    Get-AADIntLoginInformation -UserName someone@company.com

    Tenant Banner Logo                   : https://secure.aadcdn.microsoftonline-p.com/c1c6b6c8-okmfqodscgr7krbq5-p48zooi4b7m9g2zcpryoikta/logintenantbranding/0/bannerlogo?ts=635912486993671038
    Authentication Url                   : 
    Pref Credential                      : 1
    Federation Protocol                  : 
    Throttle Status                      : 1
    Cloud Instance                       : microsoftonline.com
    Federation Brand Name                : Company Ltd
    Domain Name                          : company.com
    Federation Metadata Url              : 
    Tenant Banner Illustration           : 
    Consumer Domain                      : 
    State                                : 4
    Federation Active Authentication Url : 
    User State                           : 1
    Account Type                         : Managed
    Tenant Locale                        : 0
    Domain Type                          : 3
    Exists                               : 0
    Has Password                         : True
    Cloud Instance audience urn          : urn:federation:MicrosoftOnline
    Desktop Sso Enabled                  : True
    Federation Global Version            : 

   
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Domain',Mandatory=$True)]
        [String]$Domain,

        [Parameter(ParameterSetName='User',Mandatory=$True)]
        [String]$UserName,

        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
        if([string]::IsNullOrEmpty($UserName))
        {
            $isDomain = $true
            $UserName = "nn@$Domain"
        }


        # Gather login information using different APIs
        $realm1=Get-UserRealm -UserName $UserName  -Cloud $Cloud     # common/userrealm API 1.0
        $realm2=Get-UserRealmExtended -UserName $UserName -Cloud $Cloud    # common/userrealm API 2.0
        $realm3=Get-UserRealmV2 -UserName $UserName  -Cloud $Cloud        # GetUserRealm.srf (used in the old Office 365 login experience)
        $realm4=Get-CredentialType -UserName $UserName  -Cloud $Cloud     # common/GetCredentialType (used in the "new" Office 365 login experience)

        # Create a return object
        $attributes = @{
            "Account Type" = $realm1.account_type # Managed or federated
            "Domain Name" = $realm1.domain_name
            "Cloud Instance" = $realm1.cloud_instance_name
            "Cloud Instance audience urn" = $realm1.cloud_audience_urn
            "Federation Brand Name" = $realm2.FederationBrandName
            "Tenant Locale" = $realm2.TenantBrandingInfo.Locale
            "Tenant Banner Logo" = $realm2.TenantBrandingInfo.BannerLogo
            "Tenant Banner Illustration" = $realm2.TenantBrandingInfo.Illustration
            "State" = $realm3.State
            "User State" = $realm3.UserState
            "Exists" = $realm4.IfExistsResult
            "Throttle Status" = $realm4.ThrottleStatus
            "Pref Credential" = $realm4.Credentials.PrefCredential
            "Has Password" = $realm4.Credentials.HasPassword
            "Domain Type" = $realm4.EstsProperties.DomainType
            "Federation Protocol" = $realm1.federation_protocol
            "Federation Metadata Url" = $realm1.federation_metadata_url
            "Federation Active Authentication Url" = $realm1.federation_active_auth_url
            "Authentication Url" = $realm2.AuthUrl
            "Consumer Domain" = $realm2.ConsumerDomain
            "Federation Global Version" = $realm3.FederationGlobalVersion
            "Desktop Sso Enabled" = $realm4.EstsProperties.DesktopSsoEnabled
        }
      
        # Return
        return New-Object psobject -Property $attributes
    }
}

# Return user's authentication realm from common/userrealm using API 1.0
function Get-UserRealm
{
<#
    .SYNOPSIS
    Returns authentication realm of the given user

    .DESCRIPTION
    Returns authentication realm of the given user using common/userrealm API 1.0

    .Example 
    Get-AADIntUserRealm -UserName "user@company.com"

    ver                 : 1.0
    account_type        : Managed
    domain_name         : company.com
    cloud_instance_name : microsoftonline.com
    cloud_audience_urn  : urn:federation:MicrosoftOnline

    .Example 
    Get-AADIntUserRealm -UserName "user@company.com"

    ver                        : 1.0
    account_type               : Federated
    domain_name                : company.com
    federation_protocol        : WSTrust
    federation_metadata_url    : https://sts.company.com/adfs/services/trust/mex
    federation_active_auth_url : https://sts.company.com/adfs/services/trust/2005/usernamemixed
    cloud_instance_name        : microsoftonline.com
    cloud_audience_urn         : urn:federation:MicrosoftOnline

    
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UserName, 

        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
      

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        # Call the API
        $userRealm=Invoke-RestMethod -UseBasicParsing -Uri ("$aadloginuri/common/userrealm/$UserName"+"?api-version=1.0")

        # Verbose
        Write-Verbose "USER REALM $($userRealm | Out-String)"

        # Return
        $userRealm
    }
}

# Return user's authentication realm from common/userrealm using API 2.0
function Get-UserRealmExtended
{
<#
    .SYNOPSIS
    Returns authentication realm of the given user

    .DESCRIPTION
    Returns authentication realm of the given user using common/userrealm API 2.0

    .Example
    Get-AADIntUserRealmExtended -UserName "user@company.com"

    NameSpaceType       : Managed
    Login               : user@company.com
    DomainName          : company.com
    FederationBrandName : Company Ltd
    TenantBrandingInfo  : {@{Locale=0; BannerLogo=https://secure.aadcdn.microsoftonline-p.com/xxx/logintenantbranding/0/bannerlogo?
                          ts=111; TileLogo=https://secure.aadcdn.microsoftonline-p.com/xxx/logintenantbranding/0/til
                          elogo?ts=112; BackgroundColor=#FFFFFF; BoilerPlateText=From here
                          you can sign-in to Company Ltd services; UserIdLabel=firstname.lastname@company.com;
                          KeepMeSignedInDisabled=False}}
    cloud_instance_name : microsoftonline.com

    .Example 
    Get-AADIntUserRealmExtended -UserName "user@company.com"

    NameSpaceType       : Federated
    federation_protocol : WSTrust
    Login               : user@company.com
    AuthURL             : https://sts.company.com/adfs/ls/?username=user%40company.com&wa=wsignin1.
                          0&wtrealm=urn%3afederation%3aMicrosoftOnline&wctx=
    DomainName          : company.com
    FederationBrandName : Company Ltd
    TenantBrandingInfo  : 
    cloud_instance_name : microsoftonline.com
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UserName,

        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
      
        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        # Call the API
        $userRealm=Invoke-RestMethod -UseBasicParsing -Uri ("$aadloginuri/common/userrealm/$UserName"+"?api-version=2.0")

        # Verbose
        Write-Verbose "USER REALM $($userRealm | Out-String)"

        # Return
        $userRealm
    }
}

# Return user's authentication realm from GetUserRealm.srf (used in the old Office 365 login experience)
function Get-UserRealmV2
{
<#
    .SYNOPSIS
    Returns authentication realm of the given user

    .DESCRIPTION
    Returns authentication realm of the given user using GetUserRealm.srf (used in the old Office 365 login experience)

    .Example
    Get-AADIntUserRealmV3 -UserName "user@company.com"

    State               : 4
    UserState           : 1
    Login               : user@company.com
    NameSpaceType       : Managed
    DomainName          : company.com
    FederationBrandName : Company Ltd
    CloudInstanceName   : microsoftonline.com

    .Example 
    Get-AADIntUserRealmV2 -UserName "user@company.com"

    State                   : 3
    UserState               : 2
    Login                   : user@company.com
    NameSpaceType           : Federated
    DomainName              : company.com
    FederationGlobalVersion : -1
    AuthURL                 : https://sts.company.com/adfs/ls/?username=user%40company.com&wa=wsignin1.
                              0&wtrealm=urn%3afederation%3aMicrosoftOnline&wctx=
    FederationBrandName     : Company Ltd
    CloudInstanceName       : microsoftonline.com
    
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UserName,

        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
      
        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        # Call the API
        $userRealm=Invoke-RestMethod -UseBasicParsing -Uri ("$aadloginuri/GetUserRealm.srf?login=$UserName")

        # Verbose
        Write-Verbose "USER REALM: $($userRealm | Out-String)"

        # Return
        $userRealm
    }
}

# Return user's authentication type information from common/GetCredentialType
function Get-CredentialType
{
<#
    .SYNOPSIS
    Returns authentication information of the given user

    .DESCRIPTION
    Returns authentication of the given user using common/GetCredentialType (used in the "new" Office 365 login experience)

    .Example
    Get-AADIntUserRealmExtended -UserName "user@company.com"

    Username       : user@company.com
    Display        : user@company.com
    IfExistsResult : 0
    ThrottleStatus : 1
    Credentials    : @{PrefCredential=1; HasPassword=True; RemoteNgcParams=; FidoParams=; SasParams=}
    EstsProperties : @{UserTenantBranding=System.Object[]; DomainType=3}
    FlowToken      : 
    apiCanary      : AQABAAA..A

    NameSpaceType       : Managed
    Login               : user@company.com
    DomainName          : company.com
    FederationBrandName : Company Ltd
    TenantBrandingInfo  : {@{Locale=0; BannerLogo=https://secure.aadcdn.microsoftonline-p.com/xxx/logintenantbranding/0/bannerlogo?
                          ts=111; TileLogo=https://secure.aadcdn.microsoftonline-p.com/xxx/logintenantbranding/0/til
                          elogo?ts=112; BackgroundColor=#FFFFFF; BoilerPlateText=From here
                          you can sign-in to Company Ltd services; UserIdLabel=firstname.lastname@company.com;
                          KeepMeSignedInDisabled=False}}
    cloud_instance_name : microsoftonline.com

    .Example 
    Get-AADIntUserRealmExtended -UserName "user@company.com"

    Username       : user@company.com
    Display        : user@company.com
    IfExistsResult : 0
    ThrottleStatus : 1
    Credentials    : @{PrefCredential=4; HasPassword=True; RemoteNgcParams=; FidoParams=; SasParams=; FederationRed
                     irectUrl=https://sts.company.com/adfs/ls/?username=user%40company.com&wa=wsignin1.0&wtreal
                     m=urn%3afederation%3aMicrosoftOnline&wctx=}
    EstsProperties : @{UserTenantBranding=; DomainType=4}
    FlowToken      : 
    apiCanary      : AQABAAA..A
   
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UserName,
        [Parameter(Mandatory=$False)]
        [String]$FlowToken,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud


    )
    Process
    {
        # Create a body for REST API request
        $body = @{
            "username"=$UserName
            "isOtherIdpSupported"="true"
	        "checkPhones"="true"
	        "isRemoteNGCSupported"="false"
	        "isCookieBannerShown"="false"
	        "isFidoSupported"="false"
            "originalRequest"=""
            "flowToken"=$FlowToken
        }
      
        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        # Call the API
        $userRealm=Invoke-RestMethod -UseBasicParsing -Uri ("$aadloginuri/common/GetCredentialType") -ContentType "application/json; charset=UTF-8" -Method POST -Body ($body|ConvertTo-Json)

        # Verbose
        Write-Verbose "CREDENTIAL TYPE: $($userRealm | Out-String)"

        # Return
        $userRealm
    }
}



# Return cloud instance
function Get-TenantCloud
{
<#
    .SYNOPSIS
    Returns cloud instance based on tenantId

    .Example
    Get-TenantCloud -tenantId <tenantId>
  
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$tenantId,

        [Parameter(Mandatory=$false)]
        [String]$cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        $uri = "$aadloginuri/$tenantId/.well-known/openid-configuration"
     
        try {
            $openIdConfig=Invoke-RestMethod -UseBasicParsing $uri
        }
        catch {
            Write-Verbose "failed to get cloud instance from tenant: $tenantId"
            return $NULL
        }

         if ($openIdConfig.cloud_instance_name -like   "partner.microsoftonline.cn") {
            return  "AzureChina"

        } elseif ($openIdConfig.cloud_instance_name -like   "microsoftonline.com") {
            return  "AzurePublic"

        } else { 
            write-verbose "currently supported cloud is Azure China and Azure Public. $($openIdConfig.cloud_instance_name) is not supported yet"
            return $NULL       
        }

    }
}


# Return OpenID configuration for the domain
function Get-OpenIDConfiguration
{
<#
    .SYNOPSIS
    Returns OpenID configuration of the given domain or user

    .DESCRIPTION
    Returns OpenID configuration of the given domain or user

    .Example
    Get-AADIntOpenIDConfiguration -UserName "user@company.com"

    .Example
    Get-AADIntOpenIDConfiguration -Domain company.com

    authorization_endpoint                : https://login.microsoftonline.com/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/oauth2/v2.0/authorize
    token_endpoint                        : https://login.microsoftonline.com/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/oauth2/v2.0/token
    token_endpoint_auth_methods_supported : {client_secret_post, private_key_jwt, client_secret_basic}
    jwks_uri                              : https://login.microsoftonline.com/common/discovery/keys
    response_modes_supported              : {query, fragment, form_post}
    subject_types_supported               : {pairwise}
    id_token_signing_alg_values_supported : {RS256}
    http_logout_supported                 : True
    frontchannel_logout_supported         : True
    end_session_endpoint                  : https://login.microsoftonline.com/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/oauth2/logout
    response_types_supported              : {code, id_token, code id_token, token id_token...}
    scopes_supported                      : {openid}
    issuer                                : https://sts.windows.net/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/
    claims_supported                      : {sub, iss, cloud_instance_name, cloud_instance_host_name...}
    microsoft_multi_refresh_token         : True
    check_session_iframe                  : https://login.microsoftonline.com/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/oauth2/checkses
                                            sion
    userinfo_endpoint                     : https://login.microsoftonline.com/5b62a25d-60c6-40e6-aace-8a43e8b8ba4a/openid/userinfo
    tenant_region_scope                   : EU
    cloud_instance_name                   : microsoftonline.com
    cloud_graph_host_name                 : graph.windows.net
    msgraph_host                          : graph.microsoft.com
    rbac_url                              : https://pas.windows.net

    
   
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Domain',Mandatory=$true)]
        [String]$Domain,

        [Parameter(ParameterSetName='User',Mandatory=$true)]
        [String]$UserName,

        [Parameter(Mandatory=$false)]
        [String]$appid,

        [Parameter(Mandatory=$false)]
        [String]$cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']
        # Call the API
        if([String]::IsNullOrEmpty($Domain))
        {
            $Domain = $UserName.split("@")[1].ToString()
        }

        # include the appId configuration if it has custom claim maps
        if ([String]::IsNullOrEmpty($appid)) {

            $uri = "$aadloginuri/$Domain/.well-known/openid-configuration"

        } else {
            $uri = "$aadloginuri/$Domain/.well-known/openid-configuration?appid=$appid"
        }
        
        try {
            $openIdConfig=Invoke-RestMethod -UseBasicParsing $uri
        }
        catch {
            return $null
        }

        # Return
        $openIdConfig
    }
}

# Check if the accout is dual-federation enabled, if yes, return tenant ID for all domains. Otherwise, return the tenant ID and target cloud
function get-DualFedTenant
{
<#
    .SYNOPSIS
    Returns the dual-federation info for giving domain or user SPN   
#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Domain',Mandatory=$true)]
        [String]$Domain,

        [Parameter(ParameterSetName='User',Mandatory=$true)]
        [String]$UserName
    )
    Process
    {

        # Call the API
        if([String]::IsNullOrEmpty($Domain))
        {
            $Domain = $UserName.split("@")[1].ToString()
        }
   
        $aztenantId = Get-TenantID -Domain $domain -cloud "AzurePublic"
        $mktenantId = Get-TenantID -Domain $domain -cloud "AzureChina"
        
        if ([String]::IsNullOrEmpty($aztenantId) -or [String]::IsNullOrEmpty($mktenantId) -or ($aztenantId -eq $mktenantId)) {
            $isdualfeddomain = $false
        } else {
            $isdualfeddomain = $true
        }

        if ($isdualfeddomain) {
            $globalazureTenant = $aztenantId
            $chinaazureTenant = $mktenantId
         
        } else {

            if([String]::IsNullOrEmpty($aztenantId)){
                $globalazureTenant = ""
                $chinaazureTenant = $mktenantId
            } else {
                $globalazureTenant = $aztenantId
                $chinaazureTenant = ""
            }
        }

        $tenantdualfed = @{
            "Is Dual-Fed enabled" = $isdualfeddomain
            "Azure Commercial Tenant" =   $globalazureTenant
            "Azure China Tenant" = $chinaazureTenant
        }

        $tenantdualfed
    }
}

# Get the tenant ID for the given user/domain/accesstoken
function Get-TenantID
{
<#
    .SYNOPSIS
    Returns TenantID of the given domain, user, or AccessToken

    .DESCRIPTION
    Returns TenantID of the given domain, user, or AccessToken

    .Example
    Get-AADIntTenantID -UserName "user@company.com"

    .Example
    Get-AADIntTenantID -Domain company.com

    .Example
    Get-AADIntTenantID -AccessToken $at

#>
    [cmdletbinding()]
    Param(
        [Parameter(ParameterSetName='Domain',Mandatory=$true)]
        [String]$Domain,

        [Parameter(ParameterSetName='User',Mandatory=$true)]
        [String]$UserName,

       [Parameter(ParameterSetName='AccessToken', Mandatory=$false)]
        [String]$AccessToken,

        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        if([String]::IsNullOrEmpty($AccessToken))
        {

            if([String]::IsNullOrEmpty($Domain))
            {
                $Domain = $UserName.split("@")[1].ToString()
            }
    
          Try
          {
                $OpenIdConfig = Get-OpenIDConfiguration -Domain $Domain -cloud $Cloud
                $TenantId = $OpenIdConfig.authorization_endpoint.Split("/")[3]
                $currentaadloginuri = "https://$($OpenIdConfig.authorization_endpoint.Split("/")[2])"
          }
          catch
          {
               return $null
           }
          
        }
        else
        {
            $TenantId=(Read-Accesstoken($AccessToken)).tid
        }

        if ($aadloginuri -eq $currentaadloginuri) {
            return $TenantId
        } else {
            # Return null if the tenant ID does not match with the current cloud
            return $NULL

        }
    }
}

# Check if the access token has expired
function Is-AccessTokenExpired
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

        
    )
    Process
    {
        # Read the token
        $token = Read-Accesstoken($AccessToken)
        $now=(Get-Date).ToUniversalTime()

        # Get the expiration time
        $exp=$epoch.Date.AddSeconds($token.exp)

        # Compare and return
        $retVal = $now -ge $exp

        return $retVal
    }
}

# Check if the access token signature is valid
# May 20th 2020
function Is-AccessTokenValid
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$AccessToken

    )
    Process
    {
        # Token sections
        $sections =  $AccessToken.Split(".")
        $header =    $sections[0]
        $payload =   $sections[1]
        $signature = $sections[2]

        $signatureValid = $false

        # Fill the header with padding for Base 64 decoding
        while ($header.Length % 4)
        {
            $header += "="
        }

        # Convert the token to string and json
        $headerBytes=[System.Convert]::FromBase64String($header)
        $headerArray=[System.Text.Encoding]::ASCII.GetString($headerBytes)
        $headerObj=$headerArray | ConvertFrom-Json

        # Get the signing key
        $KeyId=$headerObj.kid
        write-verbose "PARSED TOKEN HEADER: $($headerObj | Format-List | Out-String)"

        # The algorithm should be RSA with SHA-256, i.e. RS256
        if($headerObj.alg -eq "RS256")
        {
            # Get the public certificate
            $publicCert = Get-APIKeys -KeyId $KeyId
            write-verbose "TOKEN SIGNING CERT: $publicCert"
            $certBin=[convert]::FromBase64String($publicCert)

            # Construct the JWT data to be verified
            $dataToVerify="{0}.{1}" -f $header,$payload
            $dataBin = [text.encoding]::UTF8.GetBytes($dataToVerify)

            # Remove the Base64 URL encoding from the signature and add padding
            $signature=$signature.Replace("-","+").Replace("_","/")
            while ($signature.Length % 4)
            {
                $signature += "="
            }
            $signBytes = [convert]::FromBase64String($signature)

            # Extract the modulus and exponent from the certificate
            for($a=0;$a -lt $certBin.Length ; $a++)
            {
                # Read the bytes    
                $byte =  $certBin[$a] 
                $nByte = $certBin[$a+1] 

                # We are only interested in 0x02 tag where our modulus is hidden..
                if($byte -eq 0x02 -and $nByte -band 0x80)
                {
                    $a++
                    if($nbyte -band 0x02)
                    {
                        $byteCount = [System.BitConverter]::ToInt16($certBin[$($a+2)..$($a+1)],0)
                        $a+=3
                    }
                    elseif($nbyte -band 0x01)
                    {
                        $byteCount = $certBin[$($a+1)]
                        $a+=2
                    }

                    # If the first byte is 0x00, skip it
                    if($certBin[$a] -eq 0x00)
                    {
                        $a++
                        $byteCount--
                    }

                    # Now we have the modulus!
                    $modulus = $certBin[$a..$($a+$byteCount-1)]

                    # Next byte value is the exponent
                    $a+=$byteCount
                    if($certBin[$a++] -eq 0x02)
                    {
                        $byteCount = $certBin[$a++]
                        $exponent =  $certBin[$a..$($a+$byteCount-1)]
                        write-verbose "MODULUS:  $(Convert-ByteArrayToHex -Bytes $modulus)"
                        write-verbose "EXPONENT: $(Convert-ByteArrayToHex -Bytes $exponent)"
                        break
                    }
                    else
                    {
                        write-verbose "Error getting modulus and exponent"
                    }
                }
            }

            if($exponent -and $modulus)
            {
                # Create the RSA and other required objects
                $rsa = New-Object -TypeName System.Security.Cryptography.RSACryptoServiceProvider
                $rsaParameters = New-Object -TypeName System.Security.Cryptography.RSAParameters
    
                # Set the verification parameters
                $rsaParameters.Exponent = $exponent
                $rsaparameters.Modulus = $modulus
                $rsa.ImportParameters($rsaParameters)
                
                $signatureValid = $rsa.VerifyData($dataBin, $signBytes,[System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)

                $rsa.Dispose() 
                  
            }
                
        }
        else
        {
            Write-Error "Access Token signature algorithm $($headerObj.alg) not supported!"
        }

        return $signatureValid 
    }
}



# Gets OAuth information using SAML token
function Get-OAuthInfoUsingSAML
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$SAMLToken,
        [Parameter(Mandatory=$True)]
        [String]$Resource,
        [Parameter(Mandatory=$False)]
        [String]$ClientId="1b730954-1685-4b74-9bfd-dac224a7b894",
        [Parameter(Mandatory=$false)]
        [String]$scope,        
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Begin
    {
        # Create the headers. We like to be seen as Outlook.
        $headers = @{
            "User-Agent" = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; Tablet PC 2.0; Microsoft Outlook 16.0.4266)"
        }
    }
    Process
    {

        # default scope
        if([String]::IsNullOrEmpty($scope))        
        {
            $scope = "openid profile"
        }         

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        $encodedSamlToken= [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($SAMLToken))
        # Debug
        write-verbose "SAML TOKEN: $samlToken"
        write-verbose "ENCODED SAML TOKEN: $encodedSamlToken"

        # Create a body for API request
        $body = @{
            "resource"=$Resource
            "client_id"=$ClientId
            "grant_type"="urn:ietf:params:oauth:grant-type:saml1_1-bearer"
            "assertion"=$encodedSamlToken
            "scope"="$scope"
        }

        # Debug
        write-verbose "FED AUTHENTICATION BODY: "
        foreach($key in $body.keys) {
            write-verbose "$key`: $($body[$key])"
        }

        # Set the content type and call the Microsoft Online authentication API
        $contentType="application/x-www-form-urlencoded"
        try
        {
            $jsonResponse=Invoke-RestMethod -UseBasicParsing -Uri "$aadloginuri/common/oauth2/v2.0/token" -ContentType $contentType -Method POST -Body $body -Headers $headers
        }
        catch
        {
            $e = $_.Exception
            $memStream = $e.Response.GetResponseStream()
            $readStream = New-Object System.IO.StreamReader($memStream)
            while ($readStream.Peek() -ne -1) {
                Write-Error $readStream.ReadLine()
            }
            $readStream.Dispose();
        }

        return $jsonResponse
    }
}

# Return OAuth information for the given user
function Get-OAuthInfo
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Mandatory=$false)]
        [String]$tenant,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [String]$resource,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeRefreshToken=$false,
        [Parameter(Mandatory=$false)]
        [String]$ClientId="1b730954-1685-4b74-9bfd-dac224a7b894",
        [Parameter(Mandatory=$false)]
        [String]$clientsecret, # required when the client ID is not first-party app
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {

        # default scope
        $scopevalue = get-oauthscopes -resource $resource -scope $scope -authflow 'code' -IncludeRefreshToken $IncludeRefreshToken


        # get tenant
        if([String]::IsNullOrEmpty($tenant))        
        {
            $tenant = "common"
        } 



        # Get the user realm
        $userRealm = Get-UserRealm($Credentials.UserName)

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        # Check the authentication type
        if($userRealm.account_type -eq "Unknown")
        {
            Write-Error "User type  of $($Credentials.Username) is Unknown!"
            return $null
        }
        elseif($userRealm.account_type -eq "Managed")
        {
            # If authentication type is managed, we authenticate directly against Microsoft Online
            # with user name and password to get access token

            # Create a body for REST API request

            if ($script:AzureKnwonClients.Values -contains $ClientId) {
                # first-party application
                $body = @{
                    "client_id"=$ClientId
                    "grant_type"="password"
                    "username"=$Credentials.UserName
                    "password"=$Credentials.GetNetworkCredential().Password
                    "scope"="$scopevalue"
                }

            } else {

                # customer app needs client secrets
                if([String]::IsNullOrEmpty($clientsecret))  {
                    write-verbose "recommanded to use clientsecret for confidential application access"
                    
                    $body = @{
                        "client_id"=$ClientId
                        "grant_type"="password"
                        "username"=$Credentials.UserName
                        "password"=$Credentials.GetNetworkCredential().Password
                        "scope"="$scopevalue"
                    }

                } else {
                    $body = @{
                        "client_id"=$ClientId
                        "client_secret"=$clientsecret
                        "grant_type"="password"
                        "username"=$Credentials.UserName
                        "password"=$Credentials.GetNetworkCredential().Password
                        "scope"="$scopevalue"
                    }
                 }   
    
            }

            # Debug
            write-verbose "AUTHENTICATION BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }
           
            # Set the content type and call the Microsoft Online authentication API
            $contentType="application/x-www-form-urlencoded"
            $jsonResponse=Invoke-RestMethod -UseBasicParsing -Uri "$aadloginuri/$tenant/oauth2/v2.0/token" -ContentType $contentType -Method POST -Body $body

        }
        else
        {
            # If authentication type is Federated, we must first authenticate against the identity provider
            # to fetch SAML token and then get access token from Microsoft Online

            # Get the federation metadata url from user realm
            $federation_metadata_url=$userRealm.federation_metadata_url

            # Call the API to get metadata
            [xml]$response=Invoke-RestMethod -UseBasicParsing -Uri $federation_metadata_url 

            # Get the url of identity provider endpoint.
            # Note! Tested only with AD FS - others may or may not work
            $federation_url=($response.definitions.service.port | where name -eq "UserNameWSTrustBinding_IWSTrustFeb2005Async").address.location

            # login.live.com
            # TODO: Fix
            #$federation_url=$response.EntityDescriptor.RoleDescriptor[1].PassiveRequestorEndpoint.EndpointReference.Address

            # Set credentials and other needed variables
            $username=$Credentials.UserName
            $password=$Credentials.GetNetworkCredential().Password
            $created=(Get-Date).ToUniversalTime().toString("yyyy-MM-ddTHH:mm:ssZ").Replace(".",":")
            $expires=(Get-Date).AddMinutes(10).ToUniversalTime().toString("yyyy-MM-ddTHH:mm:ssZ").Replace(".",":")
            $message_id=(New-Guid).ToString()
            $user_id=(New-Guid).ToString()

            # Set headers
            $headers = @{
                "SOAPAction"="http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue"
                "Host"=$federation_url.Split("/")[2]
                "client-request-id"=(New-Guid).toString()
            }

            # Debug
            write-verbose "FED AUTHENTICATION HEADERS: $($headers | Out-String)"
            
            # Create the SOAP envelope
            $envelope=@"
                <s:Envelope xmlns:s='http://www.w3.org/2003/05/soap-envelope' xmlns:a='http://www.w3.org/2005/08/addressing' xmlns:u='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd'>
	                <s:Header>
		                <a:Action s:mustUnderstand='1'>http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
		                <a:MessageID>urn:uuid:$message_id</a:MessageID>
		                <a:ReplyTo>
			                <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>
		                </a:ReplyTo>
		                <a:To s:mustUnderstand='1'>$federation_url</a:To>
		                <o:Security s:mustUnderstand='1' xmlns:o='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd'>
			                <u:Timestamp u:Id='_0'>
				                <u:Created>$created</u:Created>
				                <u:Expires>$expires</u:Expires>
			                </u:Timestamp>
			                <o:UsernameToken u:Id='uuid-$user_id'>
				                <o:Username>$username</o:Username>
				                <o:Password>$password</o:Password>
			                </o:UsernameToken>
		                </o:Security>
	                </s:Header>
	                <s:Body>
		                <trust:RequestSecurityToken xmlns:trust='http://schemas.xmlsoap.org/ws/2005/02/trust'>
			                <wsp:AppliesTo xmlns:wsp='http://schemas.xmlsoap.org/ws/2004/09/policy'>
				                <a:EndpointReference>
					                <a:Address>urn:federation:MicrosoftOnline</a:Address>
				                </a:EndpointReference>
			                </wsp:AppliesTo>
			                <trust:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</trust:KeyType>
			                <trust:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</trust:RequestType>
		                </trust:RequestSecurityToken>
	                </s:Body>
                </s:Envelope>
"@
            # Debug
            write-verbose "FED AUTHENTICATION: $envelope"

            # Set the content type and call the authentication service            
            $contentType="application/soap+xml"
            [xml]$xmlResponse=Invoke-RestMethod -UseBasicParsing -Uri $federation_url -ContentType $contentType -Method POST -Body $envelope -Headers $headers

            # Get the SAML token from response and encode it with Base64
            $samlToken=$xmlResponse.Envelope.Body.RequestSecurityTokenResponse.RequestedSecurityToken.Assertion.OuterXml
            $encodedSamlToken= [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($samlToken))

            $jsonResponse = Get-OAuthInfoUsingSAML -SAMLToken $samlToken -Resource $Resource -ClientId $ClientId
        }
        
        if ($jsonResponse.access_token) {
            # Debug
            write-verbose "AUTHENTICATION JSON: $($jsonResponse | Out-String)"

            # Return
            $oauthifno = @{
                id_token = $jsonResponse.id_token
                access_token = $jsonResponse.access_token
                refresh_token = $jsonResponse.refresh_token
            }

            return $oauthifno            

        } else {

            return $null
        }

    }
}

# Parse access token and return it as PS object
function Read-Accesstoken
{
<#
    .SYNOPSIS
    Extract details from the given Access Token

    .DESCRIPTION
    Extract details from the given Access Token and returns them as PS Object

    .Parameter AccessToken
    The Access Token.
    
    .Example
    PS C:\>$token=Get-AADIntReadAccessTokenForAADGraph
    PS C:\>Parse-AADIntAccessToken -AccessToken $token

    aud                 : https://graph.windows.net
    iss                 : https://sts.windows.net/f2b2ba53-ed2a-4f4c-a4c3-85c61e548975/
    iat                 : 1589477501
    nbf                 : 1589477501
    exp                 : 1589481401
    acr                 : 1
    aio                 : ASQA2/8PAAAALe232Yyx9l=
    amr                 : {pwd}
    appid               : 1b730954-1685-4b74-9bfd-dac224a7b894
    appidacr            : 0
    family_name         : company
    given_name          : admin
    ipaddr              : 107.210.220.129
    name                : admin company
    oid                 : 1713a7bf-47ba-4826-a2a7-bbda9fabe948
    puid                : 100354
    rh                  : 0QfALA.
    scp                 : user_impersonation
    sub                 : BGwHjKPU
    tenant_region_scope : NA
    tid                 : f2b2ba53-ed2a-4f4c-a4c3-85c61e548975
    unique_name         : admin@company.onmicrosoft.com
    upn                 : admin@company.onmicrosoft.com
    uti                 : -EWK6jMDrEiAesWsiAA
    ver                 : 1.0

    .Example
    PS C:\>Parse-AADIntAccessToken -AccessToken $token -Validate

    Read-Accesstoken : Access Token is expired
    At line:1 char:1
    + Read-Accesstoken -AccessToken $at -Validate -verbose
    + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : NotSpecified: (:) [Write-Error], WriteErrorException
        + FullyQualifiedErrorId : Microsoft.PowerShell.Commands.WriteErrorException,Read-Accesstoken

    aud                 : https://graph.windows.net
    iss                 : https://sts.windows.net/f2b2ba53-ed2a-4f4c-a4c3-85c61e548975/
    iat                 : 1589477501
    nbf                 : 1589477501
    exp                 : 1589481401
    acr                 : 1
    aio                 : ASQA2/8PAAAALe232Yyx9l=
    amr                 : {pwd}
    appid               : 1b730954-1685-4b74-9bfd-dac224a7b894
    appidacr            : 0
    family_name         : company
    given_name          : admin
    ipaddr              : 107.210.220.129
    name                : admin company
    oid                 : 1713a7bf-47ba-4826-a2a7-bbda9fabe948
    puid                : 100354
    rh                  : 0QfALA.
    scp                 : user_impersonation
    sub                 : BGwHjKPU
    tenant_region_scope : NA
    tid                 : f2b2ba53-ed2a-4f4c-a4c3-85c61e548975
    unique_name         : admin@company.onmicrosoft.com
    upn                 : admin@company.onmicrosoft.com
    uti                 : -EWK6jMDrEiAesWsiAA
    ver                 : 1.0
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline)]
        [String]$AccessToken,
        [Parameter()]
        [Switch]$ShowDate,
        [Parameter()]
        [Switch]$Validate

    )
    Process
    {
        # Token sections
        $sections =  $AccessToken.Split(".")
        $header =    $sections[0]
        $payload =   $sections[1]
        $signature = $sections[2]

        # Convert the token to string and json
        $payloadString = Convert-B64ToText -B64 $payload
        $payloadObj=$payloadString | ConvertFrom-Json

        if($ShowDate)
        {
            # Show dates
            $payloadObj.exp=($epoch.Date.AddSeconds($payloadObj.exp)).toString("yyyy-MM-ddTHH:mm:ssZ").Replace(".",":")
            $payloadObj.iat=($epoch.Date.AddSeconds($payloadObj.iat)).toString("yyyy-MM-ddTHH:mm:ssZ").Replace(".",":")
            $payloadObj.nbf=($epoch.Date.AddSeconds($payloadObj.nbf)).toString("yyyy-MM-ddTHH:mm:ssZ").Replace(".",":")
        }

        if($Validate)
        {
            # Check the signature
            if((Is-AccessTokenValid -AccessToken $AccessToken))
            {
                Write-Verbose "Access Token signature successfully verified"
            }
            else
            {
                Write-Error "Access Token signature could not be verified"
            }

            # Check the timestamp
            if((Is-AccessTokenExpired -AccessToken $AccessToken))
            {
                Write-Error "Access Token is expired"
            }
            else
            {
                Write-Verbose "Access Token is not expired"
            }

        }

        # Debug
        write-verbose "PARSED ACCESS TOKEN: $($payloadObj | Out-String)"
        
        # Return
        $payloadObj
    }
}


# Prompts scope based on author flows
# limited to use a single resource
function get-oauthscopes
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$Resource,
        [Parameter(Mandatory=$false)]
        [String]$scope, # multiple scope required to seperate with '' or ,
        [Parameter(Mandatory=$true)]
        [ValidateSet("code", "password","client_credentials","device_code","jwt-bearer","legacy","obo","refreshtoken")]
        [String]$authflow,
        [Parameter(Mandatory=$False)]
        [bool]$IncludeRefreshToken=$false
    )
    Process
    {
        $scopevalue = ''
    
        if ([string]::IsNullOrEmpty($scope) -or $scope -eq '') {
            $scope = '.default'
        }
        $scopeitems = $scope.split(' ,')
        $resource=$resource.TrimEnd('/')
        write-verbose "set scope for auth flow: $authflow"

        if ([string]::IsNullOrEmpty($Resource)) {
            $resource = $script:AzureResources[$Cloud]["ms_graph_api"] # get MS graph resource based on cloud
        }

        write-verbose "current scope $scope"


        if ($authflow -like 'client_credentials') {
            if ($resource -like "http://*" -or $resource -like 'https://*') {
                $scopevalue = $scopevalue + " $($resource.TrimEnd('/'))/.default"
            } else {
                $scopevalue = $scopevalue + " API://$resource/.default"
            }                            
        } elseif($authflow -like 'legacy') {
   
            foreach ( $item in  $scopeitems) {
                $scopevalue = $scopevalue + " $($item.split("/")[-1])"
            }
        } else {
   
                foreach ( $item in  $scopeitems) {

                        if ($item -like "http://*" -or $item -like 'https://*' -or $item -like "api:*" -or $item -like "spn:*")  {
                            $scopevalue = $scopevalue + " $item" 
                        } else {
                            # bind with resource with the scopes

                            if ($item -like 'user_impersonation') {
                                $itemvalue = '.default'
                            } else {
                                $itemvalue = $item
                            }

                            if ($resource -like "http://*" -or $resource -like 'https://*') {
                                $scopevalue = $scopevalue + " $resource/$itemvalue"
                            } else {
                                $scopevalue = $scopevalue + " api://$resource/$itemvalue"
                            }
                        }
                    }
        }
           

        # default scope
        if ($IncludeRefreshToken -and $authflow -ne 'refreshtoken') {
            $scopevalue =  $scopevalue + " offline_access"
        } 
        
        return $scopevalue.TrimStart(' ')
    
    }
}

# Prompts for credentials and gets the access token
# Supports MFA, federation, etc.
function Prompt-Credentials
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$Resource,
        [Parameter(Mandatory=$true)]
        [String]$ClientId="1b730954-1685-4b74-9bfd-dac224a7b894" <# graph_api #>,
        [Parameter(Mandatory=$False)]
        [String]$clientSecret,
        [Parameter(Mandatory=$False)]
        [String]$Tenant,
        [Parameter(Mandatory=$False)]
        [bool]$ForceMFA=$false,
        [Parameter(Mandatory=$false)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [bool]$IncludeRefreshToken=$false,
        [Parameter(Mandatory=$false)]
        [String]$prompt="login",  # values like login, consent, admin_consent, none      
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

        if([string]::IsNullOrEmpty($RedirectUri))
        {
            $RedirectUri = Get-AuthRedirectUrl -ClientId $ClientId -Resource $Resource
        }

        # Set variables
        $auth_redirect= $RedirectUri
        $client_id=     $ClientId # Usually should be graph_api

        if ($auth_redirect -like "urn:ietf:wg:oauth:2.0:oob*") {
            $auth_redirect=[System.Web.HttpUtility]::UrlEncode($auth_redirect)
        }

        # Create the url
        $request_id=(New-Guid).ToString()
        if ($script:AzureKnwonClients.Values -contains $ClientId) {
        
            if([string]::IsNullOrEmpty($scope))
            {
                $scope = 'openid'
            }
            $scopevalue = get-oauthscopes -Resource $Resource -scope $scope -authflow 'legacy' -IncludeRefreshToken $IncludeRefreshToken
            $encodescope =  [System.Web.HttpUtility]::UrlEncode($scopevalue)

            $url="$aadloginuri/$Tenant/oauth2/authorize?resource=$resource&client_id=$client_id&response_type=code&haschrome=1&redirect_uri=$auth_redirect&client-request-id=$request_id&prompt=$prompt&scope=$encodescope"
        } else {

                            
            $scopevalue = get-oauthscopes -Resource $Resource -scope $scope -authflow 'code' -IncludeRefreshToken $IncludeRefreshToken
            $encodescope =  [System.Web.HttpUtility]::UrlEncode($scopevalue)

            $url="$aadloginuri/$Tenant/oauth2/v2.0/authorize?client_id=$client_id&response_type=code&haschrome=1&redirect_uri=$auth_redirect&client-request-id=$request_id&prompt=$prompt&scope=$encodescope"
        }
        write-verbose "oauth Url: $url"
       
        if($ForceMFA)
        {
            $url+="&amr_values=mfa"
        }

        # Azure AD Join
        if($ClientId -eq "29d9ed98-a469-4536-ade2-f981bc1d605e" -and $Resource -ne $mdm) 
        {
                $RedirectUri="ms-aadj-redir://auth/drs"
        }

        # Create the form and get output     
        $output = Show-OAuthWindow -Url $url

        # return null if the output contains error
        if(![string]::IsNullOrEmpty($output["error"])){
            Write-Error $output["error"]
            Write-Error $output["error_uri"]
            Write-Error $output["error_description"]     

            $form.Controls[0].Dispose()
            return $null
           
        }

        if ($output["code"]) {
            # Create a body for REST API request


            if([string]::IsNullOrEmpty($clientSecret)) {
                $body = @{
                    client_id=$client_id
                    grant_type="authorization_code"
                    code=$output["code"]
                    redirect_uri=$RedirectUri
                    scope = $scopevalue
                }
            } else {
                $body = @{
                    client_id=$client_id
                    client_secret=$clientSecret
                    grant_type="authorization_code"
                    code=$output["code"]
                    redirect_uri=$RedirectUri
                    scope = $scopevalue
                }
            }

            # verbose output for token request
            # Debug
            write-verbose "AUTHENTICATION BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }

            # Set the content type and call the Microsoft Online authentication API
            $contentType="application/x-www-form-urlencoded"
            $jsonResponse=Invoke-RestMethod -UseBasicParsing -Uri "$aadloginuri/$Tenant/oauth2/v2.0/token" -ContentType $contentType -Method POST -Body $body

            # return 
            $jsonResponse
        } else {
            Write-Verbose "no authorization code available from login attempts"
            $form.Controls[0].Dispose()
            return $null
        }
    }
}


# prompt Azure oAuth login window
Function Show-OAuthWindow
  {
    param (
      [Parameter(Mandatory=$true)]   
      [System.Uri] $Url
    )


   # write-verbose "oauth Url: $url"

    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420; Height=600; Url=($url) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri
        If ($Global:Uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
    }

    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
  
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440; Height=640}
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
  
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
      $output["$key"] = $queryOutput[$key]
    }

    # Dispose the control
    $form.Controls[0].Dispose()
      
    $output
}
 


function Clear-WebBrowser
{
    [cmdletbinding()]
    Param(
    )
    Process
    {
        
        # Clear the cache
        [IntPtr] $optionPointer = [IntPtr]::Zero
        $s =                      [System.Runtime.InteropServices.Marshal]::SizeOf($INTERNET_OPTION_END_BROWSER_SESSION)
        $optionPointer =          [System.Runtime.InteropServices.Marshal]::AllocCoTaskMem($s)
        [System.Runtime.InteropServices.Marshal]::WriteInt32($optionPointer, ([ref]$INTERNET_SUPPRESS_COOKIE_PERSIST).Value)
        $status =                 $WebBrowser::InternetSetOption([IntPtr]::Zero, $INTERNET_OPTION_SUPPRESS_BEHAVIOR, $optionPointer, $s)
        write-verbose "Clearing Web browser cache. Status:$status"
        [System.Runtime.InteropServices.Marshal]::Release($optionPointer)|out-null

        # Clear the current session
        $status = $WebBrowser::InternetSetOption([IntPtr]::Zero, $INTERNET_OPTION_END_BROWSER_SESSION, [IntPtr]::Zero, 0)
        write-verbose "Clearing Web browser. Status:$status"
    }
}

function Get-WebBrowserCookies
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Url
    )
    Process
    {
        $dataSize = 1024
        $cookieData = [System.Text.StringBuilder]::new($dataSize)
        $status = $WebBrowser::InternetGetCookieEx($Url,$null,$cookieData, [ref]$dataSize, $INTERNET_COOKIE_HTTPONLY, [IntPtr]::Zero)
        write-verbose "GETCOOKIEEX Status: $status, length: $($cookieData.Length)"
        if(!$status)
        {
            $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            write-verbose "GETCOOKIEEX ERROR: $LastError"
        }

        if($cookieData.Length -gt 0)
        {
            $cookies = $cookieData.ToString()
            write-verbose "Cookies for $url`: $cookies"
            Return $cookies
        }
        else
        {
            Write-Warning "Cookies not found for $url"
        }

    }
}


## GENERAL ADMIN API FUNCTIONS

# Gets Office 365 instance names (used when getting ip addresses)
function Get-EndpointInstances
{
<#
    .SYNOPSIS
    Get Office 365 endpoint instances

    .DESCRIPTION
    Get Office 365 endpoint instances
  
    .Example
    PS C:\>Get-AADIntEndpointInstances

    instance     latest    
    --------     ------    
    Worldwide    2018100100
    USGovDoD     2018100100
    USGovGCCHigh 2018100100
    China        2018100100
    Germany      2018100100

#>

    [cmdletbinding()]
    Param()
    Process
    {
        $clientrequestid=(New-Guid).ToString();
        Invoke-RestMethod -UseBasicParsing -Uri "https://endpoints.office.com/version?clientrequestid=$clientrequestid"
    }
}

# Gets Office 365 ip addresses for specific instance
function Get-EndpointIps
{
<#
    .SYNOPSIS
    Get Office 365 endpoint ips and urls

    .DESCRIPTION
    Get Office 365 endpoint ips and urls

    .Parameter Instance
    The instance which ips and urls are returned. Defaults to WorldWide.
  
    .Example
    PS C:\>Get-AADIntEndpointIps

    id                     : 1
    serviceArea            : Exchange
    serviceAreaDisplayName : Exchange Online
    urls                   : {outlook.office.com, outlook.office365.com}
    ips                    : {13.107.6.152/31, 13.107.9.152/31, 13.107.18.10/31, 13.107.19.10/31...}
    tcpPorts               : 80,443
    expressRoute           : True
    category               : Optimize
    required               : True

    id                     : 2
    serviceArea            : Exchange
    serviceAreaDisplayName : Exchange Online
    urls                   : {smtp.office365.com}
    ips                    : {13.107.6.152/31, 13.107.9.152/31, 13.107.18.10/31, 13.107.19.10/31...}
    tcpPorts               : 587
    expressRoute           : True
    category               : Allow
    required               : True

    .Example
    PS C:\>Get-AADIntEndpointIps -Instance Germany

    id                     : 1
    serviceArea            : Exchange
    serviceAreaDisplayName : Exchange Online
    urls                   : {outlook.office.de}
    ips                    : {51.4.64.0/23, 51.5.64.0/23}
    tcpPorts               : 80,443
    expressRoute           : False
    category               : Optimize
    required               : True

    id                     : 2
    serviceArea            : Exchange
    serviceAreaDisplayName : Exchange Online
    urls                   : {r1.res.office365.com}
    tcpPorts               : 80,443
    expressRoute           : False
    category               : Default
    required               : True
#>
    [cmdletbinding()]
    Param(
        [Parameter()]
        [ValidateSet('Worldwide','USGovDoD','USGovGCCHigh','China','Germany')]
        [String]$Instance="China"
    )
    Process
    {
        $clientrequestid=(New-Guid).ToString();
        Invoke-RestMethod -UseBasicParsing -Uri ("https://endpoints.office.com/endpoints/$Instance"+"?clientrequestid=$clientrequestid")
    }
}

# Gets username from authorization header
# Apr 4th 2019
function Get-UserNameFromAuthHeader
{
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Auth
    )
    
    Process
        {
        $type = $Auth.Split(" ")[0]
        $data = $Auth.Split(" ")[1]

        if($type -eq "Basic")
        {
            ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($data))).Split(":")[0]
        }
        else
        {
            (Read-Accesstoken -AccessToken $data).upn
        }
    }
}

# Creates authorization header from Credentials or AccessToken
# Apr 4th 2019
function Create-AuthorizationHeader
{
    Param(
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credentials,
        [Parameter()]
        [String]$AccessToken,
        [Parameter()]
        [String]$Resource,
        [Parameter()]
        [String]$ClientId

    )

    Process
    {
    
        if($Credentials -ne $null)
        {
            $userName = $Credentials.UserName
            $password = $Credentials.GetNetworkCredential().Password
            $auth = "Basic $([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($userName):$($password)")))"
        }
        else
        {
            # Get from cache if not provided
            $AccessToken = Get-AccessTokenFromCache -AccessToken $AccessToken -Resource $Resource -ClientId $ClientId
            $auth = "Bearer $AccessToken"
        }

        return $auth
    }
}


# Gets Microsoft online services' public keys
# May 18th 2020
function Get-APIKeys
{
    [cmdletbinding()]
    Param(
        [Parameter()]
        [String]$KeyId,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        $keys=Invoke-RestMethod -UseBasicParsing -Uri "$aadloginuri/common/discovery/keys"

        if($KeyId)
        {
            $keys.keys | Where-Object -Property kid -eq $KeyId | Select-Object -ExpandProperty x5c
        }
        else
        {
            $keys.keys
        }
        
    }
}

# Gets the AADInt credentials cache
function Get-Cache
{
<#
    .SYNOPSIS
    Dumps AADInternals credentials cache
    .DESCRIPTION
    Dumps AADInternals credentials cache
    
    .EXAMPLE
    Get-AADIntCache | Format-Table
    Name              ClientId                             Audience                             Tenant                               IsExpired HasRefreshToken
    ----              --------                             --------                             ------                               --------- ---------------
    admin@company.com 1b730954-1685-4b74-9bfd-dac224a7b894 https://graph.windows.net            82205ae4-4c4e-4db5-890c-cb5e5a98d7a3     False            True
    admin@company.com 1b730954-1685-4b74-9bfd-dac224a7b894 https://management.core.windows.net/ 82205ae4-4c4e-4db5-890c-cb5e5a98d7a3     False            True
#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        $cacheKeys = $script:tokens.keys

        # Loop through the cache elements
        foreach($key in $cacheKeys)
        {
            $accessToken=$script:tokens[$key]

        
                if([string]::IsNullOrEmpty($accessToken))
                {
                    Write-Warning "Access token with key ""$key"" not found!"
                    $script:tokens.remove($key)
                    
                } else {
    
                    $parsedToken = Read-Accesstoken -AccessToken $accessToken
                    $ClientId = $parsedToken.appid
                    $cloud = $($key.split("-")[0])
                    $refreshkey = "$cloud-$ClientId"                    
    
                    $attributes = [ordered]@{
                        "Name" =            $parsedToken.unique_name
                        "ObjectId" =        $parsedToken.oid
                        "ClientId" =        $parsedToken.appid
                        "Audience" =        $parsedToken.aud
                        "Tenant" =          $parsedToken.tid
                        "IsExpired" =       Is-AccessTokenExpired -AccessToken $accessToken
                        "HasRefreshToken" = $script:refresh_tokens.Contains($refreshkey)
                        "AuthMethods" =     $parsedToken.amr
                        "Device" =          $parsedToken.deviceid
                        "cloud" =           $cloud                        
                    }
    
                    New-Object psobject -Property $attributes
                }

            
        }
        
    }
}

# Clears the AADInt credentials cache
function Clear-Cache
{
<#
    .SYNOPSIS
    Clears AADInternals credentials cache

    .DESCRIPTION
    Clears AADInternals credentials cache
    
    .EXAMPLE
    Clear-AADIntCache
#>
    [cmdletbinding()]
    Param()
    Process
    {
        $script:tokens =         @{}
        $script:refresh_tokens = @{}
    }
}

# Gets other domains of the given tenant
function Get-TenantDomains
{
<#
    .SYNOPSIS
    Gets other domains from the tenant of the given domain

    .DESCRIPTION
    Uses Exchange Online autodiscover service to retrive other 
    domains from the tenant of the given domain. 

    The given domain SHOULD be Managed, federated domains are not always found for some reason. 
    If nothing is found, try to use <domain>.onmicrosoft.com

    .Example
    Get-AADIntTenantDomains -Domain company.com

    company.com
    company.fi
    company.co.uk
    company.onmicrosoft.com
    company.mail.onmicrosoft.com

#>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Domain,

        [Parameter(Mandatory=$false)]
        [String]$cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $autodiscoveruri = $script:AzureResources[$Cloud]["autodiscover"]
        # Create the body
        $body=@"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:exm="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:ext="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
	<soap:Header>
		<a:Action soap:mustUnderstand="1">http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetFederationInformation</a:Action>
		<a:To soap:mustUnderstand="1">https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc</a:To>
		<a:ReplyTo>
			<a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>
		</a:ReplyTo>
	</soap:Header>
	<soap:Body>
		<GetFederationInformationRequestMessage xmlns="http://schemas.microsoft.com/exchange/2010/Autodiscover">
			<Request>
				<Domain>$Domain</Domain>
			</Request>
		</GetFederationInformationRequestMessage>
	</soap:Body>
</soap:Envelope>
"@
        # Create the headers
        $headers=@{
            "Content-Type" = "text/xml; charset=utf-8"
            "SOAPAction" =   '"http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetFederationInformation"'
            "User-Agent" =   "AutodiscoverClient"
        }
        # Invoke
        $response = Invoke-RestMethod -UseBasicParsing -Method Post -uri "$autodiscoveruri/autodiscover/autodiscover.svc" -Body $body -Headers $headers

        # Return
        $response.Envelope.body.GetFederationInformationResponseMessage.response.Domains.Domain | Sort-Object
    }
}

# Gets the auth_redirect url for the given client and resource
function Get-AuthRedirectUrl
{

    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$false)]
        [String]$Resource,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        # default
        # use nativeclient as default redirect Uri 

        $azportal=$AzureResources[$Cloud]["portal"]
        $mdm=$AzureResources[$Cloud]["mdm"]

        $redirect_uri = "urn:ietf:wg:oauth:2.0:oob"

        if($ClientId -eq "1b730954-1685-4b74-9bfd-dac224a7b894") # MS graph api
        {
            $redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        }elseif($ClientId -eq "1fec8e78-bce4-4aaf-ab1b-5451cc387264")     # Teams
        {
            $redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        }
        elseif($ClientId -eq "9bc3ab49-b65d-410a-85ad-de819febfddc") # SPO
        {
            $redirect_uri = "https://oauth.spops.microsoft.com/"
        }
        elseif($ClientId -eq "c44b4083-3bb0-49c1-b47d-974e53cbdf3c") # Azure admin interface
        {
            $redirect_uri = "$azportal/signin/index"
        }
        elseif($ClientId -eq "0000000c-0000-0000-c000-000000000000") # Azure AD Account
        {
            $redirect_uri = "https://account.activedirectory.windowsazure.com/"
        }
        elseif($ClientId -eq "19db86c3-b2b9-44cc-b339-36da233a3be2") # My sign-ins
        {
            $redirect_uri = "https://mysignins.microsoft.com"
        }
        elseif($ClientId -eq "29d9ed98-a469-4536-ade2-f981bc1d605e" -and $Resource -ne $mdm) # Azure AD Join
        {
            $redirect_uri = "ms-aadj-redir://auth/drs"
        }
        elseif($ClientId -eq "0c1307d4-29d6-4389-a11c-5cbe7f65d7fa") # Azure Android App
        {
            $redirect_uri = "https://azureapp"
        }
        elseif($ClientId -eq "33be1cef-03fb-444b-8fd3-08ca1b4d803f") # OneDrive Web
        {
            $redirect_uri = "https://admin.onedrive.com/"
        }
        elseif($ClientId -eq "ab9b8c07-8f02-4f72-87fa-80105867a763") # OneDrive native client
        {
            $redirect_uri = "https://login.windows.net/common/oauth2/nativeclient"
        }
        elseif($ClientId -eq "3d5cffa9-04da-4657-8cab-c7f074657cad") # MS Commerce
        {
            $redirect_uri = "http://localhost/m365/commerce"
        }
        elseif($ClientId -eq "4990cffe-04e8-4e8b-808a-1175604b879f") # MS Partner - this flow doesn't work as expected :(
        {
            $redirect_uri = "https://partner.microsoft.com/aad/authPostGateway"
        }
        elseif($ClientId -eq "fb78d390-0c51-40cd-8e17-fdbfab77341b") # Microsoft Exchange REST API Based Powershell
        {
            $redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        }
		elseif($ClientId -eq "3b511579-5e00-46e1-a89e-a6f0870e2f5a") 
        {
            $redirect_uri = "https://windows365.microsoft.com/signin-oidc"
        }
        elseif($ClientId -eq "a0c73c16-a7e3-4564-9a95-2bdf47383716") # EXO PS
        {
            $redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"  ## !! seems not working now
        }
        

        return $redirect_uri
    }
}


# request admin consent
function Get-AdminConsent
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$true)]
        [String]$Tenant,
        [Parameter(Mandatory=$true)]
        [String]$RedirectUri,
        [Parameter(Mandatory=$false)]
        [String]$scope,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {
        if([String]::IsNullOrEmpty($scope))        
        {
            $scope = "openid"
        }
        $encodescope =  [System.Web.HttpUtility]::UrlEncode($scope)
     

        $aadloginuri = $script:AzureResources[$Cloud]['aad_login']

        # Set variables
        $auth_redirect= $RedirectUri
        $client_id=     $ClientId # Usually should be graph_api

        if ($auth_redirect -like "urn:ietf:wg:oauth:2.0:oob*") {
            $auth_redirect=[System.Web.HttpUtility]::UrlEncode($auth_redirect)
        }

        $url="$aadloginuri/$Tenant/v2.0/adminconsent?client_id=$client_id&redirect_uri=$auth_redirect&scope=$encodescope&state=12345"
    
        write-verbose "admin consent Url: $url"
       
        $output = Show-OAuthWindow -Url $url

        # return null if the output contains error
        if(![string]::IsNullOrEmpty($output["error"])){
            Write-Error $output["error"]
            Write-Error $output["error_uri"]
            Write-Error $output["error_description"]     

            $form.Controls[0].Dispose()
            return $null
           
        } else {

            return $output 

        }

    }
}



# Creates an interactive login form based on given url and auth_redirect.
function Create-LoginForm
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Url,
        [Parameter(Mandatory=$True)]
        [String]$auth_redirect,
        [Parameter(Mandatory=$False)]
        [String]$Headers,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud

    )
    Process
    {
        # Check does the registry key exists
        $regPath="HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
        if(!(Test-Path -Path $regPath )){
            Write-Warning "WebBrowser control emulation registry key not found!"
            $answer = Read-Host -Prompt "Would you like to create the registry key? (Y/N)"
            if($answer -eq "Y")
            {
                New-Item -ItemType directory -Path $regPath -Force
            }
        }

        $azureportal = $script:AzureResources[$Cloud]['portal']

        # Check the registry value for WebBrowser control emulation. Should be IE 11
        $reg=Get-ItemProperty -Path $regPath

        if([String]::IsNullOrEmpty($reg.'powershell.exe') -or [String]::IsNullOrEmpty($reg.'powershell_ise.exe'))
        {
            Write-Warning "WebBrowser control emulation not set for PowerShell or PowerShell ISE!"
            $answer = Read-Host -Prompt "Would you like set the emulation to IE 11? Otherwise the login form may not work! (Y/N)"
            if($answer -eq "Y")
            {
                Set-ItemProperty -Path $regPath -Name "powershell_ise.exe" -Value 0x00002af9
                Set-ItemProperty -Path $regPath -Name "powershell.exe" -Value 0x00002af9
                Write-Host "Emulation set. Restart PowerShell/ISE!"
                return
            }
        }

        # Create the form and add a WebBrowser control to it
 
        $form = New-Object Windows.Forms.Form
        $form.Width = 560
        $form.Height = 680
        $form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.TopMost = $true

        $web = New-Object Windows.Forms.WebBrowser
        $web.Size = $form.ClientSize
        $web.Anchor = "Left,Top,Right,Bottom"
        $form.Controls.Add($web)

        $field = New-Object Windows.Forms.TextBox
        $field.Visible = $false
        $form.Controls.Add($field)
		$field.Text = $auth_redirect

        # Clear WebBrowser control cache
        Clear-WebBrowser

         # Add an event listener to track down where the browser is
         $web.add_Navigated({
            # If the url matches the redirect url, close with OK.
            $curl=$_.Url.ToString()
			$auth_redirect = $form.Controls[1].Text							   
            Write-Debug "NAVIGATED TO: $($curl)"
            if($curl.StartsWith($auth_redirect)) {

                # Hack for Azure Portal Login. Jul 11th 2019 
                # Check whether the body has the Bearer
                if(![String]::IsNullOrEmpty($form.Controls[0].Document.GetElementsByTagName("script")))
                {
                    $script=$form.Controls[0].Document.GetElementsByTagName("script").outerhtml
                    if($script.Contains('"oAuthToken":')){
                        $s=$script.IndexOf('"oAuthToken":')+13
                        $e=$script.IndexOf('}',$s)+1
                        $oAuthToken=$script.Substring($s,$e-$s) | ConvertFrom-Json
                        $at=$oAuthToken.authHeader.Split(" ")[1]
                        $rt=$oAuthToken.refreshToken
                        $script:AccessToken = @{"access_token"=$at; "refresh_token"=$rt}
                        Write-Debug "ACCESSTOKEN $script:accessToken"
                    }
                    elseif($curl.StartsWith($azureportal))
                    {
                        Write-Debug "WAITING FOR THE TOKEN!"
                        # Do nothing, wait for it..
                        return
                    }
                }
                
                # Add the url to the hidden field
                #$form.Controls[1].Text = $curl

                $form.DialogResult = "OK"
                $form.Close()
                Write-Debug "PROMPT CREDENTIALS URL: $curl"
            } # Automatically logs in -> need to logout first
            elseif($curl.StartsWith($url)) {
                # All others
                Write-Verbose "Returned to the starting url, someone already logged in?"
            }
        })

        
        # Add an event listener to track down where the browser is going
        $web.add_Navigating({
            $curl=$_.Url.ToString()
            Write-Verbose "NAVIGATING TO: $curl"
            # SharePoint login

         if($curl -eq $auth_redirect)
           {
                $_.Cancel=$True
                $form.DialogResult = "OK"
                $form.Close()
           }
        })
        
#        $web.ScriptErrorsSuppressed = $True

        # Set the url
        if([String]::IsNullOrEmpty($Headers))
        {
            $web.Navigate($url)
        }
        else
        {
            $web.Navigate($url,"",$null,$Headers)
        }

        # Return
        return $form
    }
}
