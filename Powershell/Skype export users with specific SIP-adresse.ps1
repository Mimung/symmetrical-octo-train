<#

Title: Skype export users with specific SIP-adresse
 
Author:  Paul A. Kråkmo
Contact: paul.kraakmo@gmail.com
Github:  github.com/Mimung

Version: 1.0 - Script Created
Date created:  12 Sept 2019
Date Modified: 12 Sept 2019

Description:

This script will look up users with @domain.com at the end of their SIP adresse
Then export the users DisplayName, Accountname and LineURI to a CSV

NB: This script can be very slow because of the filtering on the SIP-adresse domain

This script will only work on a skype for business server

#>


Get-CsUser -Filter {SipAddress -like "*@domain.com"} |select-object DisplayName,SamAccountName,LineURI | export-csv -Delimiter ";" c:\temp\export.csv