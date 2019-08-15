<#

Title: Addins in office fix
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.0 - Script Created
Date created:  15 Aug 2019
Date Modified: 15 Aug 2019

Description:
KB article created into a script for easy troubleshooting

Log-file can be found on the current computer C:\Temp.txt

#>


$winword = [bool](Get-Process winword -EA SilentlyContinue)
$excel = [bool](Get-Process excel -EA SilentlyContinue)
$powerpoint = [bool](Get-Process POWERPNT -EA SilentlyContinue)
$outlook = [bool](Get-Process outlook -EA SilentlyContinue)
$program = [bool](Get-Process WinWget -EA SilentlyContinue)

$date = Get-Date -format "dd-MM-yyyy HH:mm"
$logPath = "C:\Temp\Log.txt"


if($winword -or $excel -or $powerpoint -or $outlook -or $program -eq "True")

{
Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::Ok
$MessageboxTitle = “Programs are open”
$Messageboxbody = “
Program(s) are running, please close programs marked True:
Word: $($winword)
Excel: $($excel)
PowerPoint: $($powerpoint)
Outlook: $($outlook)
Program: $($program)
”
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
[System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
return
}

$path = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
$name = "DelegateSentItemsStyle"

New-Item -Path "c:\" -Name "Temp" -ItemType "directory" -EA SilentlyContinue
Get-ItemProperty -path $path -Name $name | Select-Object DelegateSentItemsStyle | Out-File C:\Temp\Data.txt
$value = Get-Content -Path "C:\Temp\Data.txt" | Select -Index 3
Remove-Item -Path "C:\Temp\Data.txt"
$value = $value.Trim()

Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $path -Name $name
New-ItemProperty -Path $path -Name $name -Value ”3”  -PropertyType "DWORD"
