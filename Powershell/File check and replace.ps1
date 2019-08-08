#Replaces file A with either B or C
#Paul Andreas Kråkmo - paul.krakmo@visolit.no - 08.08.2019

#Config options
#-----------
$path = "C:\Temp\"
$prodConfig = "ProdConfig.txt"
$testConfig = "TestConfig.txt"
$mainConfig = "MainConfig.txt"
$controlLine = "Test"
#Line number 0 = 1 1 = 2 ... 6 = 7 etc
$line = 0
#-----------


#Checks the 7th line of $mainConfig
$readline = Get-Content -Path "$($path)$($mainConfig)" | Select -Index $line

#Checks if the 7th line of $mainConfig if it is DatabaseName=KlientAdmin_TEST
if($readline -eq $controlLine)

#If the line is $controlLine then it will copy $prodConfig and replace the $mainConfig
{
    Copy-Item "$($path)$($prodConfig)" -Destination "$($path)$($mainConfig)"
#This part between the line can be removed entierly as this only creates the notification
#---------------
    Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info 
        $balloon.BalloonTipText = 'Test config present                   Switched to Prod Config'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
#---------------
}

elseif($readline -notmatch $controlLine) 

#If the line is not $controlLine then it will copy $testConfig and replace the $mainConfig
{
    Copy-Item "$($path)$($testConfig)" -Destination "$($path)$($mainConfig)"
#This part between the line can be removed entierly as this only creates the notification
#---------------
    Add-Type -AssemblyName System.Windows.Forms 
        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
        $path = (Get-Process -id $pid).Path
        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info 
        $balloon.BalloonTipText = 'Prod config present                   Switched to Test Config'
        $balloon.Visible = $true 
        $balloon.ShowBalloonTip(2000)
#---------------
}