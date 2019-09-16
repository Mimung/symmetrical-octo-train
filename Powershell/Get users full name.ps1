$user = whoami
$username = $env:UserName
$name = Get-WMIObject Win32_UserAccount | where caption -eq $user | Select-Object -ExpandProperty FullName
Write-Host "Your name is $name"