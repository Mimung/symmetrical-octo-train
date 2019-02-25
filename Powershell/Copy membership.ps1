# Account name that you want to copy AD membership from
$username1 = "enternamehere"

# Account name that you want to copy AD membership to, do not edit this line, you will be prompted to enter the username
$username2  = Read-host "Enter username to copy to: "
 
# copy-paste process. Get-ADuser membership     | then selecting membership                       | and add it to the second user
get-ADuser -identity $username1 -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $username2