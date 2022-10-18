$modulepath = $PSHOME+"\modules\aadtokens"
$path = $PSScriptRoot
Get-Childitem $Path -Filter *.ps* -Recurse | Foreach-Object {
    $targetFile = $modulepath + $_.FullName.SubString($path.Length);
    New-Item -ItemType File -Path $targetFile -Force;
    write-host "installing powershell modules $($_.Fullname)"
    Copy-Item $_.FullName -destination $targetFile  -force -Confirm:$false
}