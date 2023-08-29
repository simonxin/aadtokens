


# get local prt token
function get-prttoken
{
    [CmdletBinding()]
    param(
        [switch]$devicecreds
    )
    Process {

        $deviceinfo = Get-aadjdeviceinfo

        if ($deviceinfo.tenantId) {

          $cloud = get-tenantcloud -tenant $($deviceinfo.tenantId)

        } else {
           write-verbose "no device info found, please check if the current device is AAD joined or registered"
           return $null
        }

        $aadlogin = $script:AzureResources[$Cloud]["aad_login"] # get aad login url

        # get nonce
        $response = Invoke-RestMethod -UseBasicParsing -Method Post -Uri "$aadlogin/Common/oauth2/token" -Body "grant_type=srv_challenge"
        $nonce = $response.Nonce

        write-verbose "get nonce: $nonce"
        

        # Create the process
        
        $browserCore = "$($env:windir)\BrowserCore\browsercore.exe"

        if (!(test-path $browserCore ) ) {
            $browserCore  = "$($env:programfiles)\Windows Security\BrowserCore\browsercore.exe"
        }

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo.FileName = $browserCore
        $p.StartInfo.UseShellExecute = $false
        $p.StartInfo.RedirectStandardInput = $true
        $p.StartInfo.RedirectStandardOutput = $true
        $p.StartInfo.CreateNoWindow = $true


        $body = @"
{
    "method":"GetCookies",
    "uri":"$aadlogin/common/oauth2/authorize?sso_nonce=$nonce",
    "sender":"$aadlogin"
}
"@
        # Start the process
        $p.Start() | Out-Null
        $stdin =  $p.StandardInput
        $stdout = $p.StandardOutput

        # Write the input
        $stdin.BaseStream.Write([bitconverter]::GetBytes($body.Length),0,4) 
        $stdin.Write($body)
        $stdin.Close()

        # Read the output
        $response=""
        while(!$stdout.EndOfStream)
        {
            $response += $stdout.ReadLine()
        }

        $p.WaitForExit()

        # Strip the stuff from the beginning of the line
        $response = $response.Substring($response.IndexOf("{")) | ConvertFrom-Json

        # Check for error
        if($response.status -eq "Fail")
        {
            Write-verbose "Error getting PRT: $($response.code). $($response.description)"
        } else {

        $tokens = $response.response.data

        if ( ($tokens | Measure-Object).Count -eq 1) {
            Write-verbose "cannot get local prt token but has device creds in response"
            $parseddevicecreds= Read-Accesstoken  $tokens 
            write-verbose $parseddevicecreds

            if ($devicecreds) {
                return $tokens
            } else {
                return $null
            }
    

        } else {

            $prttoken = $tokens[0]
            $devicetoken =  $tokens[1].split(';')[0]

            $parsedprttoken= Read-Accesstoken  $prttoken
            $parseddevicecreds= Read-Accesstoken  $devicetoken
            Write-verbose "get prt token:"
            write-verbose $parsedprttoken

            Write-verbose "get device cert:"
            write-verbose $parseddevicecreds

            if($devicecreds)
            {
                return @($devicetoken, $prttoken)
            }
            else
            {
                return $prttoken  
            }

        }

    }    
   }
}

# get AccessToken with PRT
# PRT token will update x-ms-RefreshTokenCredential

function Get-AccessTokenWithPRT
{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Cookie,
        [Parameter(Mandatory=$True)]
        [String]$Resource,
        [Parameter(Mandatory=$True)]
        [String]$ClientId,
        [Parameter(Mandatory=$False)]
        [String]$RedirectUri="urn:ietf:wg:oauth:2.0:oob",
        [switch]$GetNonce,
        [bool]$IncludeRefreshToken=$false,
        [bool]$SaveToCache=$false
    )
    Process
    {

        $deviceinfo = Get-aadjdeviceinfo

        if ($deviceinfo.tenantId) {

          $cloud = get-tenantcloud -tenant $($deviceinfo.tenantId)

        } else {
           write-verbose "no device info found, please check if the current device is AAD joined or registered"
           return $null
        }

        $aadlogin = $script:AzureResources[$Cloud]["aad_login"] # get aad login url

        $parsedCookie = Read-Accesstoken $Cookie

        # Create parameters
        $mscrid =    (New-Guid).ToString()
        $requestId = $mscrid
        
        # Create url and headers
        $url = "$aadlogin/Common/oauth2/authorize?resource=$Resource&client_id=$ClientId&response_type=code&redirect_uri=$RedirectUri&client-request-id=$requestId&mscrid=$mscrid"

        # Add sso_nonce if exist
        if($parsedCookie.request_nonce)
        {
            $url += "&sso_nonce=$($parsedCookie.request_nonce)"
        }

        $headers = @{
            "User-Agent" = ""
            "x-ms-RefreshTokenCredential" = $Cookie
            }

        write-verbose "parse prt token in header: x-ms-RefreshTokenCredential"
        write-verbose "auth url: $url"
        
        # First, make the request to get the authorisation code (tries to redirect so throws an error)
        $response = Invoke-RestMethod -UseBasicParsing -Uri $url -Headers $headers -MaximumRedirection 0 -ErrorAction SilentlyContinue

        write-verbose "RESPONSE: $($response.OuterXml)"

        # Try to parse the code from the response
        if($response.html.body.script)
        {
            $values = $response.html.body.script.Split("?").Split("\")
            foreach($value in $values)
            {
                $row=$value.Split("=")
                if($row[0] -eq "code")
                {
                    $code = $row[1]
                    Write-Verbose "CODE: $code"
                    break
                }
            }
        }
        

        if(!$code)
        {
            if($response.html.body.h2.a.href -ne $null)
            {
                $values = $response.html.body.h2.a.href.Split("&")
                foreach($value in $values)
                {
                    $row=$value.Split("=")
                    if($row[0] -eq "sso_nonce")
                    {
                        $sso_nonce = $row[1]
                        if($GetNonce)
                        {
                            # Just return the nonce
                            return $sso_nonce
                        }
                        else
                        {
                            # Invalid PRT, nonce is required
                            Write-Warning "Nonce needed. Try New-AADIntUserPRTToken with -GetNonce switch or -Nonce $sso_nonce parameter"
                            break
                        }
                    }
                }
                
            }
            
            throw "Code not received!"
        }

        # Create the body
        $body = @{
            client_id =    $ClientId
            grant_type =   "authorization_code"
            code =         $code
            redirect_uri = $RedirectUri
        }

        # Make the second request to get the access token
        $response = Invoke-RestMethod -UseBasicParsing -Uri "$aadlogin/common/oauth2/token" -Body $body -ContentType "application/x-www-form-urlencoded" -Method Post

        # Save the tokens to cache
        if($SaveToCache -and ![string]::IsNullOrEmpty($response.access_token))
        {


            Write-Verbose "ACCESS TOKEN: SAVE TO CACHE"
            $Script:tokens["$cloud-$ClientId-$Resource"] =         $response.access_token

            if(![string]::IsNullOrEmpty($response.refresh_token)) {
                $Script:refresh_tokens["$cloud-$ClientId"] = $response.refresh_token
            }

            
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

# clear cloud app cache
function Clear-CloudApCache
{
    [CmdletBinding()]
    param(
    )
    Process
    {

        $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name

        if ($user -eq 'NT AUTHORITY\SYSTEM') {

            $cloudapguid = get-childitem C:\Windows\System32\config\systemprofile\AppData\Local\Microsoft\Windows\CloudAPCache\AzureAd
            foreach ($guid in $cloudapguid) {
                $cloudap = $guid.FullName           
                write-host "clear cloud app cache: $cloudap"
                Remove-Item -Path "$cloudap" -Recurse -Force
                
            }
        } else {
            Write-Error "cloud ap cache can only be deleted under local system"
            Write-Verbose "You can switch to local system context with command like: psexec.exe -i -s powershell.exe"
        }
  
    }

}

# Gets the AAD join info of the local device
function Get-aadjdeviceinfo
{
<#
    .SYNOPSIS
    Shows the Azure AD Join information of the local device.

    .DESCRIPTION
    Shows the Azure AD Join information of the local device.

    .Example
    PS C\:>Get-AADIntLocalDeviceJoinInfo

    JoinType           : Joined
    RegistryRoot       : HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin
    CertThumb          : CEC55C2566633AC8DA3D9E3EAD98A599084D0C4C
    CertPath           : Cert:\LocalMachine\My\CEC55C2566633AC8DA3D9E3EAD98A599084D0C4C
    TenantId           : afdb4be1-057f-4dc1-98a9-327ffa079cca
    DeviceId           : f4a4ea70-b196-4305-9531-018c3bcfc112
    ObjectId           : d625e2e9-8465-4513-b6c9-8d34a3735d41
    KeyName            : 8bff0b7f02f6256b521de95a77d4e70d_934bc9f7-04ef-43d8-a343-610b736a4030
    KeyFriendlyName    : Device Identity Key
    IdpDomain          : login.windows.net
    UserEmail          : JohnD@company.com
    AttestationLevel   : 0
    AikCertStatus      : 0
    TransportKeyStatus : 0
    DeviceDisplayName  : WIN-JohnD
    OsVersion          : 10.0.19044.1288
    DdidUpToDate       : 0
    LastSyncTime       : 1643370347

    .Example
    PS C\:>Get-AADIntLocalDeviceJoinInfo
    WARNING: This device has a TPM, exporting keys probably does not work!

    JoinType           : Joined
    RegistryRoot       : HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin
    CertThumb          : FFDABA36622C66F1F9104703D77603AE1964E92B
    CertPath           : Cert:\LocalMachine\My\FFDABA36622C66F1F9104703D77603AE1964E92B
    TenantId           : afdb4be1-057f-4dc1-98a9-327ffa079cca
    DeviceId           : e4c56ee8-419a-4421-bff4-1d3cb1c85ead
    ObjectId           : b62a31e9-8268-485f-aba8-69696cdf3048
    KeyName            : C:\ProgramData\Microsoft\Crypto\PCPKSP\[redacted]\[redacted].PCPKEY
    KeyFriendlyName    : Device Identity Key
    IdpDomain          : login.windows.net
    UserEmail          : package_c1b50acc-82f6-4a19-ba87-e62e5f7fbeee@company.com
    AttestationLevel   : 0
    AikCertStatus      : 0
    TransportKeyStatus : 3
    DeviceDisplayName  : cloudpc-80153
    OsVersion          : 10.0.19044.1469
    DdidUpToDate       : 0
    LastSyncTime       : 1643622945
#>
    [CmdletBinding()]
    param()
    Process
    {
        $AADJoinRoot       = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin"
        $AADRegisteredRoot = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin"

        # Check the join type and construct return value
        if(Test-Path -Path "$AADJoinRoot\JoinInfo")
        {

            write-verbose "device is joined to AAD tenant. Load join info from registery: $AADJoinRoot"
            $joinRoot = $AADJoinRoot
            $certRoot = "LocalMachine"
            $attributes = [ordered]@{
                "JoinType"     = "Joined"
                "RegistryRoot" = $AADJoinRoot
            }
        }
        elseif(Test-Path -Path "$AADRegisteredRoot\JoinInfo")
        {
            write-verbose "device is registered to AAD tenant. Load join info from registery: $AADRegisteredRoot"   
            $joinRoot = $AADRegisteredRoot
            $certRoot = "CurrentUser"
            $attributes = [ordered]@{
                "JoinType"     = "Registered"
                "RegistryRoot" = $AADRegisteredRoot
            }
        }
        else
        {
            write-verbose "cannot find device join info from registery"
            return $null
        }
        
        # Get the Device certificate thumbnail from registery (assuming the device can only be joined once)
        $regItem = (Get-ChildItem -Path "$joinRoot\JoinInfo\").Name
        $certThumbnail = $regItem.Substring($regItem.LastIndexOf("\")+1)
        $certificate   = Get-Item -Path "Cert:\$certRoot\My\$certThumbnail"


        if (!$certificate) {
           write-verbose "cannot load device certificate from certiicate store"
        } 

        $oids = Parse-CertificateOIDs -Certificate $certificate

        $attributes["CertThumb"      ] = "$certThumbnail"
        $attributes["CertPath"       ] = "Cert:\$certRoot\My\$certThumbnail"
        $attributes["TenantId"       ] = $oids.TenantId
        $attributes["DeviceId"       ] = $oids.DeviceId
        $attributes["ObjectId"       ] = $oids.ObjectId

        # This will fail for DeviceTransportKey because running as Local System
        try
        {
            $attributes["dkprivkeyName"        ] = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate).key.uniquename
            $attributes["dkprivKeyFriendlyName"] = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate).key.uipolicy.FriendlyName
        }
        catch
        {
            # Okay
        }

        if ( !(test-path $attributes["dkprivkeyName"])) { 
            write-verbose "cannot load device certificate private key from local file system"
        }

        # Read the join info
        $regItem = Get-Item -Path "$joinRoot\JoinInfo\$certThumbnail"
        $valueNames = $regItem.GetValueNames()
        foreach($name in $valueNames)
        {
            $attributes[$name] = $regItem.GetValue($name)
        }


        # Check the TPM
        if($attributes["TransportKeyStatus"] -eq 3)
        {
            Write-Verbose "Transport key stored in TPM"
            $transportKeys = Get-LocalDeviceTransportKeys -JoinType $attributes['JoinType'] -IdpDomain  $attributes['idpDomain'] -TenantId  $attributes['tenantId']  -UserEmail  $attributes['UserEmail'] -TPMKey

        } else {

            Write-Verbose "Transport key stored in local file system. Try to dump the transport key"
            $transportKeys = Get-LocalDeviceTransportKeys -JoinType $attributes['JoinType'] -IdpDomain  $attributes['idpDomain'] -TenantId  $attributes['tenantId']  -UserEmail  $attributes['UserEmail']

        }

        
        $attributes["tkprivname"] = $transportKeys.tkprivname
        $attributes["tkprivregpath"] = $transportKeys.tkprivregpath
        $attributes["tkprivstore"] = $transportKeys.tkprivstore      

        return New-Object psobject -Property $attributes
    }
}


# This file contains utility functions for local AAD Joined devices
function Get-LocalDeviceTransportKeys
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [ValidateSet('Joined','Registered')]
        [String]$JoinType,
        [Parameter(Mandatory=$True)]
        [String]$IdpDomain,
        [Parameter(Mandatory=$True)]
        [String]$TenantId,
        [Parameter(Mandatory=$True)]
        [String]$UserEmail,
        [Parameter(Mandatory=$false)]
        [switch]$TPMkey
    )
    Begin
    {
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
    }
    Process
    {
        # Calculate registry key parts
        $idp    = Convert-ByteArrayToHex -Bytes ($sha256.ComputeHash([text.encoding]::Unicode.GetBytes($IdpDomain)))
        $tenant = Convert-ByteArrayToHex -Bytes ($sha256.ComputeHash([text.encoding]::Unicode.GetBytes($TenantId)))
        $email  = Convert-ByteArrayToHex -Bytes ($sha256.ComputeHash([text.encoding]::Unicode.GetBytes($UserEmail)))
        $sid    = Convert-ByteArrayToHex -Bytes ($sha256.ComputeHash([text.encoding]::Unicode.GetBytes(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)))
        

        if($JoinType -eq "Joined")
        {
            $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Cryptography\Ngc\KeyTransportKey\PerDeviceKeyTransportKey\$Idp\$tenant"
        }
        else
        {
            $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Cryptography\Ngc\KeyTransportKey\$sid\$idp\$($tenant)_$($email)"
        }

        if((Test-Path -Path $registryPath) -eq $false)
        {
            Throw "The device seems not to be Azure AD joined or registered. Registry key not found: $registryPath"
        }

        # Get the Transport Key name from registry
        try
        {
            if ($TPMkey) {
               $sessionkey = "$registryPath\TpmKeyTransportKeyName"
                write-verbose "try to get session key name from registry: $sessionkey"
               $transPortKeyName = Get-ItemPropertyValue -Path "$registryPath" -Name "TpmKeyTransportKeyName"
 
            } else { 
                $sessionkey = "$registryPath\SoftwareKeyTransportKeyName"
                write-verbose "try to get session key name from registry: $sessionkey"
                $transPortKeyName = Get-ItemPropertyValue -Path "$registryPath" -Name "SoftwareKeyTransportKeyName"
              
            }
        }
        catch
        {            
            Throw "Unable to get the transport key name from registry: $registryPath"
        }

        Write-Verbose "TransportKey name: $transportKeyName`n"

        if ($TPMkey) {
            Write-Verbose "skip to get private key from TPM"

            $sessionkeyinfo = @{
                "tkprivname" = $transPortKeyName
                "tkprivregpath" = $sessionkey
                "tkprivstore" = "TPM"
            }

            return  $sessionkeyinfo

        } else {
            # Loop through the system keys 
            $haskey = $false
            $systemKeys = Get-ChildItem -Path "$env:ALLUSERSPROFILE\Microsoft\Crypto\SystemKeys"
            foreach($systemKey in $systemKeys)
            {
                Write-Verbose "Parsing $($systemKey.FullName)"
                $keyBlob = Get-Content $systemKey.FullName -Encoding byte

                # Parse the blob to get the name
                $key = Parse-CngBlob -Data $keyBlob
                if($key.name -eq $transPortKeyName)
                {
                    Write-Verbose "Transport Key found! "
                    write-Verbose $systemKey
                    $haskey = $true

                    $sessionkeyinfo = @{
                        "tkprivname" = $transPortKeyName
                        "tkprivregpath" = $sessionkey
                        "tkprivstore" = "$($systemKey.FullName)"
                    }
        
                    return  $sessionkeyinfo

                }
            }

            # no private key found from local file system
            if ($haskey -eq $false) {
                $sessionkeyinfo = @{
                    "tkprivname" = $transPortKeyName
                    "tkprivregpath" = $sessionkey
                    "tkprivstore" = "unable to get transport private key from key store $($env:ALLUSERSPROFILE)\Microsoft\Crypto\SystemKeys)"
                }
    
                return  $sessionkeyinfo
            }

        }
    }
    End
    {
        $sha256.Dispose()
    }
}

# Parses the oid values of the given certificate

function Parse-CertificateOIDs
{

    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    Process
    {
        function Get-OidRawValue
        {
            Param([byte[]]$RawValue)
            Process
            {
                # Is this DER value?
                if($RawValue.Length -gt 2 -and ($RawValue[2] -eq $RawValue.Length-3 ))
                {
                    return $RawValue[3..($RawValue.Length-1)] 
                }
                else
                {
                    return $RawValue
                }
            }
        }
        $retVal = New-Object psobject
        foreach($ext in $Certificate.Extensions)
        {
            switch($ext.Oid.Value)
            {
               "1.2.840.113556.1.5.284.2" {
                    $retVal | Add-Member -NotePropertyName "DeviceId" -NotePropertyValue ([guid][byte[]](Get-OidRawValue -RawValue $ext.RawData))
                
               } 
               "1.2.840.113556.1.5.284.3" {
                    $retVal | Add-Member -NotePropertyName "ObjectId" -NotePropertyValue ([guid][byte[]](Get-OidRawValue -RawValue $ext.RawData))
                
               } 
               "1.2.840.113556.1.5.284.5" {
                    $retVal | Add-Member -NotePropertyName "TenantId" -NotePropertyValue ([guid][byte[]](Get-OidRawValue -RawValue $ext.RawData))
                
               }
               "1.2.840.113556.1.5.284.8" {
                    # Tenant region
                    # AF = Africa
                    # AS = Asia
                    # AP = Australia/Pasific
                    # EU = Europe
                    # ME = Middle East
                    # NA = North America
                    # SA = South America
                    $retVal | Add-Member -NotePropertyName "Region"   -NotePropertyValue ([text.encoding]::UTF8.getString([byte[]](Get-OidRawValue -RawValue $ext.RawData)))
               }
               "1.2.840.113556.1.5.284.7" {
                    # JoinType
                    # 0 = Registered
                    # 1 = Joined
                    $retVal | Add-Member -NotePropertyName "JoinType" -NotePropertyValue ([int]([text.encoding]::UTF8.getString([byte[]](Get-OidRawValue -RawValue $ext.RawData))))
               }    
            }
        }

        return $retVal
    }
}
