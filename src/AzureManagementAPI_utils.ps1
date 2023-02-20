
# Calls the Azure AD IAM API
function Call-AzureAADIAMAPI
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        $Body,
        [Parameter(Mandatory=$false)]
        $AccessToken,
        [Parameter(Mandatory=$True)]
        $Command,
        [Parameter(Mandatory=$False)]
        [ValidateSet('Put','Get','Post','Delete')]
        [String]$Method="Get",
        [Parameter(Mandatory=$False)]
        [String]$Version = "2.0",
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $aadiam = $script:AzureResources[$Cloud]['aad_iam'].trimend("/")  # get AAD IAM endpoint

        $clientId = $script:AzureKnwonClients['graph_api'] # client ID of azure portal

        $aadiamapi = $script:AzureKnwonClients["adibizaux"] # aad iam


        
        if ([string]::IsNullOrEmpty($AccessToken)) {
            $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $aadiamapi -cloud $cloud 
            
            if ([string]::IsNullOrEmpty($AccessToken)) {
            
                write-verbose "no valid AAD IAM access token detected in cache. try to request a new access token"
                $accesstoken = Get-AccessTokenForAADIAMAPI -SaveToCache  -cloud $cloud 
               # $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $msgraphapi
            }
            
        } 


        # Check if the giving token is expired or not existing
        if($(Is-AccessTokenExpired($AccessToken)) -or [string]::IsNullOrEmpty($AccessToken))
            {
                write-verbose "AccessToken has expired or not valid"
                throw "AccessToken has expired or no invalid token exists"
            }
                
        

        $headers=@{
            "Authorization" = "Bearer $AccessToken"
            "X-Requested-With" = "XMLHttpRequest"
            "x-ms-client-request-id" = (New-Guid).ToString()
        }
        
        if (![string]::IsNullOrEmpty($body)) {
            # convert to json format
            if ($(test-json $body)) {
                $jsonbody = $body
                $bodyobj = $body | convertfrom-json

                $hash = @{}
                $bodyobj.psobject.properties | foreach{$hash[$_.Name]= $_.Value}
                $body = $hash
        
            } else {
                $jsonbody = $body |  ConvertTo-Json -Depth 5 
            }

        } else {
            $jsonbody = $NULL
        }

        if ($command -like "$aadiam/api/*") {

            $url = $comamnd

        } else {
            $url = "$aadiam/api/$command`?api-version=$Version"
        }


        write-verbose "call AAD IAM API $url"

        if ($body) {
            write-verbose "FED AAD IAM requests BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }
        }

        # Call the API
        try {
            $response = Invoke-RestMethod -UseBasicParsing -Uri  $url  -ContentType "application/json; charset=utf-8" -Headers $headers -Method $Method -Body $jsonbody
        }
        catch {

            $e = $_.Exception
            $memStream = $e.Response.GetResponseStream()
            $readStream = New-Object System.IO.StreamReader($memStream)
            while ($readStream.Peek() -ne -1) {
                Write-Error $readStream.ReadLine()
            }
            $readStream.Dispose();

        }

        # Return
        if($response.StatusCode -eq $null)
        {
            return $response
        }
    }
}

# Calls the Azure Management API
function Call-AzureManagementAPI
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        $Body,
        [Parameter(Mandatory=$false)]
        $AccessToken,
        [Parameter(Mandatory=$false)]
        $clientId,
        [Parameter(Mandatory=$false)]
        [String]$redirectUri,
        [Parameter(Mandatory=$false)]
        [String]$tenant,
        [Parameter(ParameterSetName='Command',Mandatory=$True)]
        [string]$Command,
        [Parameter(ParameterSetName='resourceId',Mandatory=$true)]
        [String]$resourceId,  # target resource Id like  /subscriptions/{subscriptionId}}/resourceGroups/{resourcegroup}/providers/Microsoft.Compute/virtualMachines/{vmname}
        [Parameter(ParameterSetName='resourceId',Mandatory=$false)]
        [String]$operation,  # opertion like restart
        [Parameter(ParameterSetName='resourceId',Mandatory=$false)]
        [String]$apiversion,  # API-version
        [Parameter(Mandatory=$false)]
        [switch]$headerresponse,  # API-version
        [Parameter(Mandatory=$False)]
        [ValidateSet('Put','Get','Post','Delete')]
        [String]$Method="Get",
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {


        # if tenant cloud does not match with default cloud, will use tenant cloud  
        if(![string]::IsNullOrEmpty($Tenant))
        {
            $tenantcloud = Get-TenantCloud -tenantId $Tenant
            write-verbose "tenant cloud is detected as $tenantcloud"
            if(![string]::IsNullOrEmpty($tenantcloud)) {
                $cloud =  $tenantcloud
            }
        }

        $azuremanagement = $script:AzureResources[$Cloud]['azure_mgmt_api'].trimend("/") # get Azure Management API

        if ([string]::IsNullOrEmpty($clientId)) {
            $clientId = $script:AzureKnwonClients["graph_api"]
        }
    
   

        if ([string]::IsNullOrEmpty($AccessToken)) {
            $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $azuremanagement  -cloud $cloud 
            
            if ([string]::IsNullOrEmpty($AccessToken)) {
            
                write-verbose "no valid Azure Management access token detected in cache. try to request a new access token"
                $accesstoken = Get-AccessTokenForAzureManagement -SaveToCache  -cloud $cloud -Tenant $tenant -RedirectUri $RedirectUri -clientid $clientId

            }
            
        } 

        # Check if the giving token is expired or not existing
        if($(Is-AccessTokenExpired($AccessToken)) -or [string]::IsNullOrEmpty($AccessToken))
            {
                write-verbose "AccessToken has expired or not valid"
                throw "AccessToken has expired or no invalid token exists"
            }
                

        
        $headers=@{
            "Authorization" = "Bearer $AccessToken"
            "X-Requested-With" = "XMLHttpRequest"
            "x-ms-client-request-id" = (New-Guid).ToString()
        }

        if (![string]::IsNullOrEmpty($body)) {
            # convert to json format
            if ($(test-json $body)) {
                $jsonbody = $body
                $bodyobj = $body | convertfrom-json # -AsHashtable only works in powershell 7

                $hash = @{}
                $bodyobj.psobject.properties | foreach{$hash[$_.Name]= $_.Value}
                $body = $hash
        
            } else {
                $jsonbody = $body |  ConvertTo-Json -Depth 5 
            }

        } else {
            $jsonbody = $NULL
        }


        # if the command is not provided, we will need to combine the command based on the provided subscription/resource type/resourcegroup/resource
        if ([string]::IsNullOrEmpty($Command)) {


            $command = $resourceId.trimstart("/")
            # add resource group in command if it existing
            if (![string]::IsNullOrEmpty($operation)){
                
                $command =  $command+"/"+$operation
            }

            if ([string]::IsNullOrEmpty($apiversion)){
                
                $apiversion = get-azuremanagementapiversion -AccessToken $AccessToken  -resourceId $resourceId
                write-verbose "use Apiversion $apiversion for resource $resourceId"
            }
               $command =  $command+"`?api-version="+$apiversion 
               $url="$azuremanagement/$command"

        } else {
            $command = $command.trimstart("/")
            # try to fullfil API version
            if (!($command -like "*api-version=2*")) {
                $apiversion = get-azuremanagementapiversion -AccessToken $AccessToken -resourceId $Command
                write-verbose "command has no API version added, Use Apiversion $apiversion for Azure Management API namespace $resourcetype"
                if ($command.split('?').Length -gt 1){
                    $command  = $command+"`&api-version="+ $apiversion
                } else {
                    $command  = $command+"`?api-version="+ $apiversion
                }
            }

            if ($command -like "https://*" -or $command -like "http://*") {
                $url=$command
            } else {
                $url="$azuremanagement/$command"

            }
            
        }


        # Call the API
        write-verbose "call Azure Management API $url"

        if ($body) {
            write-verbose "FED Azure Management API requests BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }           
        }

        # Call the API
        try {
            if ($headerresponse) {
                $response=Invoke-WebRequest -UseBasicParsing -Uri $url -Method $Method -Headers $headers  -ContentType "application/json; charset=utf-8" -body $jsonbody
           
                write-verbose "return headers and conent"
                $responseresult = @{
                    value = $response.content | convertfrom-json
                    header = $response.headers
                }
                return $responseresult
            
            
            } else {
                $response=Invoke-RestMethod -UseBasicParsing -Uri $url -Method $Method -Headers $headers  -ContentType "application/json; charset=utf-8" -body $jsonbody
   
                write-verbose "Azure Management REST API call is successful with a response code 200"
                if ($response.value) {
                    return $response.value
                } else {
                    return $response
                }
   
            }
        
           
        }
        catch {
            write-verbose "Azure Management REST API call failed with error details below:"
            if (!$_.Exception) {
                $e = $_.Exception
                $memStream = $e.Response.GetResponseStream()
                $readStream = New-Object System.IO.StreamReader($memStream)
                while ($readStream.Peek() -ne -1) {
                    Write-Error $readStream.ReadLine()
                }
                $readStream.Dispose()

            } else {
                throw $_
            }
            return $NULL

        }

    }
}




# get the lastest API version based on resource namespace and resource type

# Calls the Azure Management API
function get-azuremanagementapiversion
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        $AccessToken,
        [Parameter(ParameterSetName='command',Mandatory=$True)]
        $command,
        [Parameter(ParameterSetName='resourceId',Mandatory=$true)]
        [String]$resourceId, 
        [Parameter(ParameterSetName='resourceId',Mandatory=$false)]
        [String]$operation,  
        [Parameter(Mandatory=$false)]
        [string]$apiversion="2021-04-01",
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        $azuremanagement = $script:AzureResources[$Cloud]['azure_mgmt_api'] # get Azure Management API
        
              
        $headers=@{
            "Authorization" = "Bearer $AccessToken"
            "X-Requested-With" = "XMLHttpRequest"
            "x-ms-client-request-id" = (New-Guid).ToString()
        }

        # if command is used, need to split the command to get subscription Id and resource namespace
        if (![string]::IsNullOrEmpty($command)) {
            $resourceIditems = extract-azureresourceID $command

        } else {

            $resourceIditems = extract-azureresourceID $resourceId
        }
        
        $subscriptionID = $resourceIditems.subscriptionID
        $resourceprovider = $resourceIditems.resourceprovider
        $resourcetype = $resourceIditems.resourcetype


        # if no subsription Id provided, try to get API with tenant scope
        if (![string]::IsNullOrEmpty($subscriptionId)) {

            $url="$azuremanagement/subscriptions/$subscriptionID/providers/$resourceprovider`?api-version=$apiversion"
        } else {
            $url="$azuremanagement/providers/$resourceprovider`?api-version=$apiversion"
        }

        Write-Verbose "query API version from $url"
        $response=Invoke-RestMethod -UseBasicParsing -Uri $url -Method GET -Headers $headers

        # Call the API
        try {
            $response=Invoke-RestMethod -UseBasicParsing -Uri $url -Method GET -Headers $headers
        }
        catch {
            $e = $_.Exception
            $memStream = $e.Response.GetResponseStream()
            $readStream = New-Object System.IO.StreamReader($memStream)
            while ($readStream.Peek() -ne -1) {
                Write-Error $readStream.ReadLine()
            }
            $readStream.Dispose();
        }
    
        if ($response.resourcetypes) {

            $currenttype =  $response.resourcetypes | where {$_.resourcetype -like $resourcetype}
            # return first available version
            if ( $currenttype ) {
                return $currenttype.apiVersions[0]
            } else {

                # try the API version of the root resource type
                $roottype = $resourcetype.Split("/")[0]
                $currenttype =  $response.resourcetypes | where {$_.resourcetype -like $roottype}
                return $currenttype.apiVersions[0]
            }
        } else {
            return $NULL
        }
        
    }
}