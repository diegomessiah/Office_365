$credentials = Get-Credential -Credential UserAdmin@Company.com
Write-Output "Getting the Exchange Online cmdlets"
$session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -ConfigurationName Microsoft.Exchange -Credential $credentials `
    -Authentication Basic -AllowRedirection
    Import-PSSession $session
Import-Module MsOnline
Connect-MsolService -Credential $credential
$csv = ".\PCDevices.csv"
$results = @()
$mailboxUsers = get-mailbox -resultsize unlimited
$pcDevice = @()
 
foreach($user in $mailboxUsers)
{
$UPN = $user.UserPrincipalName
$displayName = $user.DisplayName
 
$pcDevices = Get-MsolDevice -RegisteredOwnerUpn  $UPN
       
      foreach($pcDevice in $pcDevices)
      {
          Write-Output "Getting info about a device for $displayName"
          $properties = @{
          Name = $user.name
          UPN = $UPN
          DisplayName = $displayName
          FriendlyName = $pcDevice.DisplayName
          OSType = $pcDevice.DeviceOSType
          ClientVersion = $pcDevice.DeviceOSVersion
          DeviceType = $pcDevice.DeviceTrustType
          DeviceTrustLevel = $pcDevice.DeviceTrustLevel
          LastLogon = $pcDevice.ApproximateLastLogonTimestamp
          Owner =  $pcDevice.RegisteredOwners
          }
          $results += New-Object psobject -Property $properties
      }
}
 
$results | Select-Object Name,UPN,DisplayName,FriendlyName,OSType,ClientVersion,DeviceType,DeviceTrustLevel,LastLogon,Owner | Export-Csv -notypeinformation -Path $csv 
Remove-PSSession $session
