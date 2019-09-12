<#

Title: View users AD Image
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.0 - Script Created
Date created:  12 sept 2019
Date Modified: 12 sept 2019

Description:

If a photo is set on the users AD-account, 
this script will then download that image as a JPG-file so that you can see it

#>

Import-Module ActiveDirectory
 
$UserName = "Username"
$date = Get-Date -format "dd-MM-yyyy HH:mm"
$PicturePath = "C:\temp\$($UserName)$($date).jpg"
 
$user = Get-ADUser $UserName -Properties thumbnailPhoto
$user.thumbnailPhoto | Set-Content $PicturePath -Encoding byte 