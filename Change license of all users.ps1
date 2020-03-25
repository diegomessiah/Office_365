# Script:   Assign License Bulk Users (not hybrit)
# Purpose:  This script to change old license's users to new license (bulk users not hybrit 365)
# Author:   Diego Messiah | https://github.com/diegomessiah

Connect-MsolService

$Users = Import-csv ".\users.csv"
$AccountSkuId = "reseller-account:ENTERPRISEPACK"
$UsageLocation = "IL"

$ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "ENTERPRISEPACK"}

# LicenseOptions, prepare the new license #
# To disable sublicenses -DisabledPlans for example Flow365 ###
$LicenseOptions =  New-MsolLicenseOptions -AccountSkuId "reseller-account:ENTERPRISEPACK" -DisabledPlans FLOW_O365_P2


#Assigning the license and usage location ###
$Users | ForEach-Object {
Set-msoluser -UserPrincipalName $_.UserPrincipalName -UsageLocation $UsageLocation
Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuId -LicenseOptions $LicenseOptions 
}
