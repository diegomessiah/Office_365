# Script:   Disabled Users
# Purpose:  Get create list of disabled users (Boolean) on the mailbox Tenant and export to csv file
# Author:   Diego Messiah | https://github.com/diegomessiah

#Accept input parameters  
Param(  
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]  
    [ValidateSet('AllUsers','LicensedUsers','UnLicensedUsers','BlockedUsers','LicensedAndBlockedUsers')]
    [string] $ReportType ,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $OutputFile 
) 

$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
Import-Module MSOnline 
Get-PSSession | Remove-PSSession ## Disconnect last connection to 365
Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
   
#Set default output file path if not passed.
if ([string]::IsNullOrEmpty($OutputFile) -eq $true) 
{ 
    $OutputFile = ".\Office365Users.csv"      
}

$o365Users;

switch ($ReportType) {
         "AllUsers"  
                     {
                         $o365Users= Get-MsolUser -All
                     }
         "LicensedUsers" 
                     {
                       $o365Users= Get-MsolUser -All | Where {$_.IsLicensed -eq $True}
                     }
         "UnLicensedUsers" 
                     {
                       $o365Users= Get-MsolUser -All | Where {$_.IsLicensed -eq $False}
                     }
         "BlockedUsers" 
                     {
                       $o365Users= Get-MsolUser -All  | Where {$_.BlockCredential -eq $True}
                     }
         "LicensedAndBlockedUsers" 
                     {
                       $o365Users= Get-MsolUser -All  | Where {$_.IsLicensed -eq $True -AND $_.BlockCredential -eq $True}
                     }
          default
                    {
                      $o365Users= Get-MsolUser -All
                    }
} 

#Export user details to CSV.
$o365Users | Select DisplayName,UserPrincipalName, IsLicensed, BlockCredential |
Export-CSV $OutputFile -NoTypeInformation -Encoding UTF8
