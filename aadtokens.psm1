
# verify dependent modules are loaded
#$DependentModules = ''

#$Installed = Import-Module $DependentModules -PassThru -ErrorAction SilentlyContinue | Where-Object { $_.name -In $DependentModules }
#$missing = $DependentModules | Where-Object { $_ -notin $Installed.name }
#if ($missing) {
#    Write-host "    [+] Module dependencies not found [$missing]. Attempting to install." -ForegroundColor Green
#    Install-Module $missing -Force -AllowClobber -Confirm:$false -Scope CurrentUser
#    Import-Module $missing
#}

$Path = [System.IO.Path]::Combine($PSScriptRoot, 'src')
Get-Childitem $Path -Filter *.ps1 -Recurse | Foreach-Object {
    write-host "import powershell modules for "$_.Fullname
    Import-Module $_.Fullname -Force
}

