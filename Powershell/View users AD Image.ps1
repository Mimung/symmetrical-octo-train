<#

Title: View users AD Image
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.1 - Script Created
Date created:  12 sept 2019
Date Modified: 16 sept 2019

Description:

If a photo is set on the users AD-account, 
this script will then download that image as a JPG-file so that you can see it

1.1: Added minor edit, it will now create a new folder C:\Temp\AdPhotos
and store the pictures here

It is now possible to enter the name in the comand window after running

#>

Import-Module ActiveDirectory

New-Item -Path "c:\" -Name "Temp" -ItemType "directory" -EA SilentlyContinue
New-Item -Path "c:\Temp\" -Name "AdPhotos" -ItemType "directory" -EA SilentlyContinue
 
$UserName = Read-host "Enter username to get the picture from: "
$date = Get-Date -format "dd-MM-yyyy HH:mm"
$PicturePath = "C:\temp\AdPhotos\$($UserName)$($date).jpg"
 
$user = Get-ADUser $UserName -Properties thumbnailPhoto
$user.thumbnailPhoto | Set-Content $PicturePath -Encoding byte
Write-Host "Downloading of $UserName AD-photo is complete, image is saved here: $Picturepath "