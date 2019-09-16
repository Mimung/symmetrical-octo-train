<#

Title: Deleagete Sent Items Fix
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 2.0 - Script Created
Date created:  15 Aug 2019
Date Modified: 16 Sept 2019

Description:
This script checks the registry key DelegateSentItemsStyle value and changes it to 3

!! The script will not run if the proccess Outlook is running, you will be prompted with a balloontip

It basicly does this:
https://docs.microsoft.com/en-us/exchange/troubleshoot/shared-mailboxes/sent-mail-is-not-saved

V2.0: Removed the hacky part of the script and replaced that part with one string:
Get-ItemProperty -path KEYPATH -Name KEYNAME | Select-Object -ExpandProperty KEYNAME

Also made the script more easier to replicate for other purposes

Updated the description accordingly

#>

# Some configurable options

$process = 'Outlook'
$keypath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
$keyname = "DelegateSentItemsStyle"
$value = Get-ItemProperty -path $keypath -Name $keyname | Select-Object -ExpandProperty $keyname


# Checks if Outlook is running, stops the script if it is
if ( [bool](Get-Process $process -EA SilentlyContinue) ) {
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
        $balloon.BalloonTipTitle = "$($process) Delegate Access"
        $balloon.BalloonTipText = "$($process) is running                                Please close outlook and try again"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        $script:balloon.Dispose()
        return
}

<#
If the value is equal to 3, then a balloon message will inform user about that

If the value is not equal to 3, then it will delete the current registry key and create a new one with the
correct value. it then starts $process. balloon-tip will be displayed when its complete
#>
if($value -eq 3){
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
return
}
else
{
Remove-ItemProperty -path $keypath -Name $keyname
New-ItemProperty -Path $keypath -Name $keyname -Value ”3”  -PropertyType "DWORD"
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $balloon.BalloonTipTitle = "$($process) Delegate Access"
        $balloon.BalloonTipText = 'Value is now correct                                Changes have been done                   Starting outlook'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        $script:balloon.Dispose()
Start-Process ($process)
return
}