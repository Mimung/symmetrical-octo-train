#Enter AD group name here:
$group = "Test AD group"
#Filepath to where you want to store the file Exmaple \\127.0.0.1\MyFolder\ or C:\MyFolder
$path = "\\127.0.0.1\MyFolder\"
#Enter Filename
$filename = "MyList"
#File Extension
$extension = ".csv"

#Do not edit below this line

Get-ADGroupMember -identity $group | select Name,@{Name="Office";Expression={ Get-ADUser $_.SamAccountName -Properties Office | Select -ExpandProperty Office }} | Export-csv -Delimiter ";" -path "$($path)$($filename)$($extension)" -NoTypeInformation