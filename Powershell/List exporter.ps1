$group = "AD groupname here"
$path = "file path here"
$filename = "filename + extension here"

#Do not edit below this line

Get-ADGroupMember -identity $gruppe | select Name,@{Name="Office";Expression={ Get-ADUser $_.SamAccountName -Properties Office | Select -ExpandProperty Office }} | Export-csv -Delimiter ";" -path "$($path)$($filename)" -NoTypeInformation