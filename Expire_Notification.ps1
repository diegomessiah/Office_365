# Script:   Password Expire Notification Office 365
# Purpose:  This script send notification about password expire when user password due to expire
# Author:   Diego Messiah | https://github.com/diegomessiah

# SMTP
$smtpServer="mail.company.com"
$expireindays = 15 
$from = "Notification Password <no-reply@company.com>" 
$logging = "Enabled" 
$logFile = ".\mylog.csv"
$testing = "disabled" 
$testRecipient = "mymailbox@company.com 
 
# LOG 
# Check Logging Settings 
$date = Get-Date -Format ddMMyyyy 
if (($logging) -eq "Enabled") 
{ 
# Test Log File Path 
$logfilePath = (Test-Path $logFile) 
if (($logFilePath) -ne "True") 
{ 
# Create CSV File and Headers 
New-Item $logfile -ItemType File 
Add-Content $logfile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn" 
} 
}
 
Import-Module MSOnline
Get-PSSession | Remove-PSSession # Disconnect last connection to 365
$AdminNames = globaladmin@company.com
$Pass = ConvertTo-SecureString 'Password' -AsPlainText -Force
$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
Connect-MSolService -credential $credential

# Get Users From MSOL where Passwords Expire 
$users = get-msoluser | where { $_.PasswordNeverExpires -eq $false } 
$domain = Get-MSOLDomain | where {$_.IsDefault -eq $true } 
$maxPasswordAge = ((Get-MsolPasswordPolicy -domain $domain.Name).ValidityPeriod).ToString() 
 
# Process Each User for Password Expiry 
# 
foreach ($user in $users) 
{ 
$Name = $user.DisplayName 
$emailaddress = $user.UserPrincipalName 
$passwordSetDate = $user.LastPasswordChangeTimestamp 
$expireson = $passwordsetdate + $maxPasswordAge 
$today = (get-date) 
$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days 

# Set Greeting based on Number of Days to Expiry.

# Check Number of Days to Expiry 
$messageDays = $daystoexpire

if (($messageDays) -ge "1") 
{ 
$messageDays = "in " + "$daystoexpire" + " days." 
} 
else 
{ 
$messageDays = "TODAY." 
}

# Email Subject Set Here 
$subject="Your password will expire $messageDays" 

# Email Body Set Here, Note: You can use HTML, including Images. 
$body =" 
Dear $name, 
<p> Your Office 365 e-mail Password will expire $messageDays.<br> 
You can change your password through the Office 365 web portal at <a href=https://www.office.com>www.office.com.</a></p>
<p> If you need instructions on how to access the portal, please contact NoName Support Team on +9555555.<br> 
Thank you, <br> </p> 
<p>NoName LTD<br></p>" 

# If Testing Is Enabled - Email Administrator 
if (($testing) -eq "Enabled") 
{ 
$emailaddress = $testRecipient 
} # End Testing

# If a user has no email address listed 
if (($emailaddress) -eq $null) 
{ 
$emailaddress = $testRecipient 
}# End No Valid Email

# Send Email Message 
if (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays)) 
{ 
# If Logging is Enabled Log Details 
if (($logging) -eq "Enabled") 
{ 
Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson" 
} 
# Send Email Message 
Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High
} 
} 
Remove-PSSession $Session 
}
