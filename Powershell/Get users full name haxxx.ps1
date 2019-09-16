New-Item -Path "c:\" -Name "Temp" -ItemType "directory" -EA SilentlyContinue
$user = whoami
$username = $env:UserName
$nameCreate = Get-WMIObject Win32_UserAccount | where caption -eq $user | select FullName | ft -hide | Out-File C:\Temp\Name.txt
$name = Get-Content -Path "C:\Temp\Name.txt" | Select -Index 1
Remove-Item -Path "C:\Temp\Name.txt"

Write-Host Your name is $name