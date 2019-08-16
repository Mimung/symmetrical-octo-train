<#

Title: Deleagete Sent Items Fix
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.0 - Script Created
Date created:  15 Aug 2019
Date Modified: 15 Aug 2019

Description:
This script checks the registry key DelegateSentItemsStyle value and changes it to 3
by checking the current value by dumping the information to a text-file if the 4 line of the file is not equal to 3
then it will delete the current DelegateSentItemsStyle key and create a new one with the DWORD value 3.
Then it deltes the file C:\Temp\DelegateSentItemsStyle.txt and starts outlook

!! The script will not run if the proccess Outlook is running, you will be prompted with a balloontip

It basicly does this:
https://docs.microsoft.com/en-us/exchange/troubleshoot/shared-mailboxes/sent-mail-is-not-saved

#>

# Checks if Outlook is running, stops the script if it is
if ( [bool](Get-Process outlook -EA SilentlyContinue) ) {
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
        $balloon.BalloonTipTitle = "Outlook Delegate Access"
        $balloon.BalloonTipText = 'Outlook is running                                Please close outlook and try again'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        $script:balloon.Dispose()
        return
}

<#
 Creates the folder C:\Temp and pipes out the value DelegateSentItemsStyle into a text-file that is then feeded back into the script
 and then it checks line 4 of that document and loads that into a value
 #>
New-Item -Path "c:\" -Name "Temp" -ItemType "directory" -EA SilentlyContinue
Get-ItemProperty -path HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences -Name DelegateSentItemsStyle | Select-Object DelegateSentItemsStyle | Out-File C:\Temp\DelegateSentItemsStyle.txt
$value = Get-Content -Path "C:\Temp\DelegateSentItemsStyle.txt" | Select -Index 3
# This part is to remove the extra spaces that is infront of the value it needs
$value = $value.Trim()
$value.Trim()


<#
If the value is equal to 3, then a balloon message will inform user about that and it then removes the data-file that was created earlier.

If the value is not equal to 3, then it will delete the current registry key and create a new one with the
correct value. it then removes the data-file that was created earlier and then starts Outlook. balloon-tip will be displayed

#>
if($value.Trim() -eq 3){
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $balloon.BalloonTipTitle = "Outlook Delegate Access"
        $balloon.BalloonTipText = 'Value is correct                                No changes have been done'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        $script:balloon.Dispose()
    Remove-Item -Path "C:\Temp\DelegateSentItemsStyle.txt"
return
}
else
{
Remove-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name "DelegateSentItemsStyle"
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name "DelegateSentItemsStyle" -Value ”3”  -PropertyType "DWORD"
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $balloon.BalloonTipTitle = "Outlook Delegate Access"
        $balloon.BalloonTipText = 'Value is now correct                                Changes have been done                   Starting outlook'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        $script:balloon.Dispose()
    Remove-Item -Path "C:\Temp\DelegateSentItemsStyle.txt"
Start-Process Outlook
return
}