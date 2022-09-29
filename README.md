# Introduction 
Powershell modules used for Azure AD oauth2 token generation

# Getting Started
you may download the module to local.
Then run in powershell admin shell with command like:
import-module ./aadtokens.psm1 -force

# Build and Test
Sample command (add -verbose will show ebug information which include all parameters when call Azure AD authorizationg and token endpoint):
# use authorization code flow and first party applcation -- wrong token demo
$authcodetoken1 = Get-AccessTokenForMSGraph -verbose

# try to call msgraph API with the access token generated
Connect-MgGraph -AccessToken $authcodetoken1 -Environment China
Get-Mguser
Disconnect-MgGraph


# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)