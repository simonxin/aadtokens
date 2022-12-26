# This script contains utility functions for PIM at https://api.azrbac.azurepim.identitygovernance.azure.cn



# Calls the provisioning SOAP API
function Call-MSPIMAPI
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$clientId,
        [Parameter(Mandatory=$true)]
        [String]$API,
        [Parameter(Mandatory=$False)]
        [ValidateSet('v1','v2','v3')]
        [String]$ApiVersion="v2",
        [Parameter(Mandatory=$False)]
        [ValidateSet('Put','Get','Post','Delete','PATCH')]
        [String]$Method="Get",
        [Parameter(Mandatory=$False)]
        $Body,
        [Parameter(Mandatory=$False)]
        $Headers,
        [Parameter(Mandatory=$False)]
        [String]$QueryString,
        [Parameter(Mandatory=$False)]
        [int]$MaxResults=200,
        [Parameter(Mandatory=$false)]
        [String]$Cloud=$script:DefaultAzureCloud
    )
    Process
    {

        # set msgraph api

        $mspimapi = $script:AzureResources[$Cloud]["ms-pim"].trimend("/") 

        if ([string]::IsNullOrEmpty($clientId)) {
            $clientId = $script:AzureKnwonClients["graph_api"]
        }

        if ([string]::IsNullOrEmpty($AccessToken)) {
            $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $msgraphapi 
            
            if ([string]::IsNullOrEmpty($AccessToken)) {
            
                write-verbose "no valid PIM access token detected in cache. try to request a new access token"
                $accesstoken = get-accesstokenformsgraph -SaveToCache
               # $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $msgraphapi
            }
            
        } 

        # Check if the giving token is expired or not existing
        if($(Is-AccessTokenExpired($AccessToken)) -or [string]::IsNullOrEmpty($AccessToken))
            {
                write-verbose "AccessToken has expired or not valid"
                throw "AccessToken has expired or no invalid token exists"
            }
                


        # Set the required variables
        # $TenantID = (Read-Accesstoken $AccessToken).tid

        if($Headers -eq $null)
        {
            $Headers=@{}
        }
        $Headers["Authorization"] = "Bearer $AccessToken"
        
                
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


        # format API
        $API = $API.TrimStart("/")

        # Create the url
        # add query string if exists
        if([String]::IsNullOrEmpty($QueryString)) {
                $url = "$mspimapi/api/$($ApiVersion)/$($API)"
            } else {
                $url = "$mspimapi/api/$($ApiVersion)/$($API)?$QueryString"
        }
        
        write-verbose "call PIM API with method $Method"

        if ($body) {
            write-verbose "FED PIM API requests BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }
        }

        # Call the API
        try {
            $error.Clear()
            $response = Invoke-RestMethod -UseBasicParsing -Uri $url -ContentType "application/json" -Method $Method  -Headers $Headers -Body $jsonbody
        }
        catch {
           $error
           return $null
        }

        # Check if we have more items to fetch
        if($response.psobject.properties.name -match '@odata.nextLink')
        {
            $items=$response.value.count

            # Loop until finished or MaxResults reached
            while(($url = $response.'@odata.nextLink') -and $items -lt $MaxResults)
            {
                # Return
                $response.value
                     
                $response = Invoke-RestMethod -UseBasicParsing -Uri $url -ContentType "application/json" -Method $Method -Headers $Headers -Body $jsonbody
                $items+=$response.value.count
            }

            # Return
            $response.value
            
        }
        else
        {

            # Return
            if($response.psobject.properties.name -match "Value")
            {
                return $response.value 
            }
            else
            {
                return $response
            }
        }

    }
}

