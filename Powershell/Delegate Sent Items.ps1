<#
Title: 		Delegate Sent Items
Version:	1.0
Author: 	Paul A. Kråkmo
Company:	Visolit
Changelog: 15.08.19 - 1.0 - Script created



#>
if ( [bool](Get-Process outlook -EA SilentlyContinue) ) {
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning  
        $balloon.BalloonTipText = 'Please close Outlook, and try again'
        $balloon.BalloonTipTitle = "Outlook is running"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        return
}

New-Item -Path 'C:\Temp' -ItemType Directory  -erroraction 'silentlycontinue'
$path = "C:\Temp\DelegateSentItemsStyle.txt"
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name DelegateSentItemsStyle | select-object DelegateSentItemsStyle | Out-File -FilePath $path
$readline = Get-Content -Path "$($path)" | Select -Index 3
$value = $readline.Trim()
Remove-Item -Path $path -Force
# write-host $value

if($value -eq 3)

{
  Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info  
        $balloon.BalloonTipText = 'Correct value present, no change was done'
        $balloon.BalloonTipTitle = "Delegate Access"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
        return
}
else
{
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name DelegateSentItemsStyle
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name DelegateSentItemsStyle -Value "3" -PropertyType "DWORD"
 Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info  
        $balloon.BalloonTipText = 'Correct value not present, change was done. Starting Outlook'
        $balloon.BalloonTipTitle = "Delegate Access"
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
    Start-Process outlook
        return
}
