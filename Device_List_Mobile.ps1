# Script:   Get-Remote-LocalAdmins
# Purpose:  List of devices (Tablet & Mobile) connected to Tenant Mailbox
# Author:   Diego Messiah | https://github.com/diegomessiah	
 
$credentials = Get-Credential -Credential Admin@ACME.com
Write-Output "Getting the Exchange Online cmdlets"
 
$session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ -ConfigurationName Microsoft.Exchange -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $session

$csv = ".\MobileDevices.csv"
$results = @()
$mailboxUsers = get-mailbox -resultsize unlimited
$mobileDevice = @()
 
foreach($user in $mailboxUsers)
{
$UPN = $user.UserPrincipalName
#Get-MobileDevice
$mobileDevices = Get-MobileDevice -Mailbox $UPN
       
      foreach($mobileDevice in $mobileDevices)
      {
          Write-Output "Getting info about a device for $user"
          $properties = @{
          Name = $user.name
          UserDisplayName = $mobileDevice.UserDisplayName
          UPN = $UPN
          ClientType = $mobileDevice.ClientType
          DeviceModel = $mobileDevice.DeviceModel
          DeviceOS = $mobileDevice.DeviceOS
          DeviceTelephoneNumber = $mobileDevice.DeviceTelephoneNumber
          FirstSyncTime = $mobileDevice.FirstSyncTime
          IsValid = $mobileDevice.IsValid
          ExchangeObjectId = $mobileDevice.ExchangeObjectId 
          IsManaged = $mobileDevice.IsManaged
          IsCompliant = $mobileDevice.IsCompliant
          IsDisabled = $mobileDevice.IsDisabled
          }
          $results += New-Object psobject -Property $properties
      }
}

$results | Select-Object Name,UserDisplayName,UPN,ClientType,DeviceModel,DeviceOS,DeviceTelephoneNumber,FirstSyncTime,IsValid,ExchangeObjectId,IsManaged,IsCompliant,IsDisabled | Export-Csv -notypeinformation -Path $csv

Remove-PSSession $session
