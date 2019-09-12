<#

Title: Office addins loadbehvior fix
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.0 - Script Created
Date created:  06 sept 2019
Date Modified: 06 sept 2019

Description:

Checks if Office applications is open, 
if they are closed it will proceede with checking the LoadBehavior values in registry and set them to 3 if they are not 3

Log-file can be found on the current computer C:\Temp.txt

This code is anonymized as this was created for a specific program that adds some addins to the different Office Applications

#>


$winword = [bool](Get-Process winword -EA SilentlyContinue)
$excel = [bool](Get-Process excel -EA SilentlyContinue)
$powerpoint = [bool](Get-Process POWERPNT -EA SilentlyContinue)
$outlook = [bool](Get-Process outlook -EA SilentlyContinue)
$date = Get-Date -format "dd-MM-yyyy HH:mm"
$logPath = "C:\Temp\Log.txt"
New-Item -Path "c:\" -Name "Temp" -ItemType "directory" -EA SilentlyContinue


if($winword -or $excel -or $powerpoint -or $outlook -eq "True")

{
Add-Content -Path $logPath -Value "$($date) - One of the following programs was running Outlook: $($outlook) - Excel: $($excel) - Word: $($winword) - Powerpoint: $($powerpoint)"
Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::Ok
$MessageboxTitle = “Program(s) are open”
$Messageboxbody = “
Please close program(s) marked 'True':
Word: $($winword)
Excel: $($excel)
PowerPoint: $($powerpoint)
Outlook: $($outlook)
”
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
[System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
return
}

#WORD

$keyname = "LoadBehavior"
$keylocation = "HKCU:\Software\Microsoft\Office\Word\Addins\"
$name = "LoadBehaviorWord"
$OutputFile = "C:\Temp\$($name).txt"
Get-ItemProperty -path $keylocation -Name LoadBehavior | Select-Object LoadBehavior | Out-File $OutputFile
$value = Get-Content -Path $OutputFile | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


if($value.Trim() -eq 3){
     Add-Content -Path $logPath -Value "$($date) - $($name) is $($value) (Default: 3)"
     Remove-Item -Path $OutputFile
}
else
{
Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $keylocation -Name $keyname
New-ItemProperty -Path $keylocation -Name $keyname -Value ”3”  -PropertyType "DWORD"
Remove-Item -Path $OutputFile
}

#Powerpoint

$keyname = "LoadBehavior"
$keylocation = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\"
$name = "LoadBehaviorPowerpoint"
$OutputFile = "C:\Temp\$($name).txt"
Get-ItemProperty -path $keylocation -Name LoadBehavior | Select-Object LoadBehavior | Out-File $OutputFile
$value = Get-Content -Path $OutputFile | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


if($value.Trim() -eq 3){
     Add-Content -Path $logPath -Value "$($date) - $($name) is $($value) (Default: 3)"
     Remove-Item -Path $OutputFile
}
else
{
Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $keylocation -Name $keyname
New-ItemProperty -Path $keylocation -Name $keyname -Value ”3”  -PropertyType "DWORD"
Remove-Item -Path $OutputFile
}

#Excel

$keyname = "LoadBehavior"
$keylocation = "HKCU:\Software\Microsoft\Office\Excel\Addins\"
$name = "LoadBehaviorExcel"
$OutputFile = "C:\Temp\$($name).txt"
Get-ItemProperty -path $keylocation -Name LoadBehavior | Select-Object LoadBehavior | Out-File $OutputFile
$value = Get-Content -Path $OutputFile | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


if($value.Trim() -eq 3){
     Add-Content -Path $logPath -Value "$($date) - $($name) is $($value) (Default: 3)"
     Remove-Item -Path $OutputFile
}
else
{
Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $keylocation -Name $keyname
New-ItemProperty -Path $keylocation -Name $keyname -Value ”3”  -PropertyType "DWORD"
Remove-Item -Path $OutputFile
}

#Outlook 1

$keyname = "LoadBehavior"
$keylocation = "HKCU:\Software\Microsoft\Office\Outlook\Addins\"
$name = "LoadBehaviorOutlook1"
$OutputFile = "C:\Temp\$($name).txt"
Get-ItemProperty -path $keylocation -Name LoadBehavior | Select-Object LoadBehavior | Out-File $OutputFile
$value = Get-Content -Path $OutputFile | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


if($value.Trim() -eq 3){
     Add-Content -Path $logPath -Value "$($date) - $($name) is $($value) (Default: 3)"
     Remove-Item -Path $OutputFile
}
else
{
Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $keylocation -Name $keyname
New-ItemProperty -Path $keylocation -Name $keyname -Value ”3”  -PropertyType "DWORD"
Remove-Item -Path $OutputFile
}

#Outlook 2

$keyname = "LoadBehavior"
$keylocation = "HKCU:\Software\Microsoft\Office\Outlook\Addins\"
$name = "LoadBehaviorOutlook2"
$OutputFile = "C:\Temp\$($name).txt"
Get-ItemProperty -path $keylocation -Name LoadBehavior | Select-Object LoadBehavior | Out-File $OutputFile
$value = Get-Content -Path $OutputFile | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


if($value.Trim() -eq 3){
     Add-Content -Path $logPath -Value "$($date) - $($name) is $($value) (Default: 3)"
     Remove-Item -Path $OutputFile
}
else
{
Add-Content -Path $logPath -Value "$($date) - $($name) was $($value) (Default: 3)"
Remove-ItemProperty -path $keylocation -Name $keyname
New-ItemProperty -Path $keylocation -Name $keyname -Value ”3”  -PropertyType "DWORD"
Remove-Item -Path $OutputFile
}
