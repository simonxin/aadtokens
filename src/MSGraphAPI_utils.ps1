# This script contains utility functions for MSGraph API at https://graph.microsoft.com



# Calls the provisioning SOAP API
function Call-MSGraphAPI
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false)]
        [String]$clientId,
        [Parameter(Mandatory=$false)]
        [String]$redirectUri,
        [Parameter(Mandatory=$false)]
        [String]$tenant,
        [Parameter(Mandatory=$false)]
        [String]$API,
        [Parameter(Mandatory=$False)]
        [ValidateSet('beta','v1.0')]
        [String]$ApiVersion="beta",
        [Parameter(Mandatory=$False)]
        [ValidateSet('Put','Get','Post','Delete','PATCH','update')]
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

        $msgraphapi = $script:AzureResources[$Cloud]["ms_graph_api"].trimend("/") 

        if ([string]::IsNullOrEmpty($clientId)) {
            $clientId = $script:AzureKnwonClients["graph_api"]
        }   


        if ([string]::IsNullOrEmpty($AccessToken)) {
            $accesstoken = get-accesstokenfromcache  -ClientID $clientId  -resource $msgraphapi -cloud $cloud
            
            if ([string]::IsNullOrEmpty($AccessToken)) {
            
                write-verbose "no valid MS Graph access token detected in cache. try to request a new access token"
                $accesstoken = get-accesstokenformsgraph -SaveToCache -cloud $Cloud -Tenant $tenant -RedirectUri $RedirectUri -clientid $clientId
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
        # use /me as ms graph API if not provided
        if ([string]::IsNullOrEmpty($API)) {
            $url = "$msgraphapi/$($ApiVersion)/me"
        } else { 

            # add query string if exists
            if([String]::IsNullOrEmpty($QueryString)) {
                $url = "$msgraphapi/$($ApiVersion)/$($API)"
            } else {
                $url = "$msgraphapi/$($ApiVersion)/$($API)?$QueryString"
            }
        }


        write-verbose "call MS Graph API with method $Method"

        if ($body) {
            write-verbose "FED MS Graph API requests BODY: "
            foreach($key in $body.keys) {
                write-verbose "$key`: $($body[$key])"
            }
        }

        # Call the API
        try {
            $response = Invoke-RestMethod -UseBasicParsing -Uri $url -ContentType "application/json" -Method $Method  -Headers $Headers -Body $jsonbody
           
        }
        catch {

            $e = $_.Exception
            $memStream = $e.Response.GetResponseStream()
            $readStream = New-Object System.IO.StreamReader($memStream)
            while ($readStream.Peek() -ne -1) {
                Write-Error $readStream.ReadLine()
            }
            $readStream.Dispose()

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

