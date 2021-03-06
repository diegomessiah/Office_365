# Script:   Mailbox Integrity & Rules
# Purpose: Check Mailbox Rules (also hidden rules) on Tenant Mailbox
# Author:   Diego Messiah | https://github.com/diegomessiah	

# Parameters
Param(
    [Parameter(Mandatory=$false)]
    [alias("m")]
    [switch]$mfa,
	
    [Parameter(Mandatory=$false)]
    [alias("")]
    $user,

    [Parameter()]
    [alias("o")]
    $outputfile = "Rules" + (Get-Date -Format "MM-dd-yyyy_HHmmss").ToString() + ".json"	
)

$banner = @'

--------------------------------------------------------
            Get Mailbox Rule
--------------------------------------------------------

'@
# Function for printing out information in color
function color_out
{
	$mstring = $args[0]
	$mcolor = $args[1]
	$current_fc = $host.ui.RawUI.ForegroundColor
	$host.ui.RawUI.ForegroundColor = $mcolor 
	Write-Output $mstring
	$host.ui.RawUI.ForegroundColor = $current_fc
}

#Parse a rule description and add contents to provided custom PSObject
function rule_parser
{
	#get the rule description
	$description = $args[0]
	#get the PsObject
	$PSobject = $args[1]
	
	#break up the description into an array based on the newline character
	$description = ($description.Split([Environment]::NewLine) | ?{$_ -match "\S"})
	
	#track where we are in the rule description
	$ifSection = $False
	$takeSection = $False
	
	#keep track of what "if" and "take action" we're currently on
	$ifCount = 1
	$takeCount = 1
	
	#loop through description
	foreach($line in $description)
	{
		$line = $line.Trim()
		
		#enterting condition section 
		if($line.startswith("If"))
		{
			$ifSection = $True
			$takeSection = $False
		}
		elseif($line.startswith("Take")) 
		{
			$ifSection = $False
			$takeSection = $True
		}
		else
		{
			if($ifSection)
			{
				$name = "condition" + [string]$ifCount
				$ifcount += 1
				Add-Member -InputObject $PsObject -NotePropertyName $name -NotePropertyValue $line
			}
			elseif($takeSection)
			{
				$name = "action" + [string]$takeCount
				$takeCount += 1
				Add-Member -InputObject $PsObject  -NotePropertyName $name -NotePropertyValue $line
			}
		}  
	}
}

function O365_permission_check
{
	$session_import_result = $args[0]
	#ACCESS CHECK 1
	if(!($session_import_result.ExportedCommands.ContainsKey("Get-Mailbox") -and $session_import_result.ExportedCommands.ContainsKey("Get-InboxRule")))
	{
		color_out "[-] ERROR!: O365 PSSession failed admin check! [1/2]" "Red"
		color_out "[-] ERROR!: O365 PSSession failed imported command check!`n" "Red"
		Exit
	}

	#ACCESS CHECK 2
	if(((Get-Mailbox -ResultSize 2 -WarningAction "SilentlyContinue").Count) -lt 2)
	{
		color_out "[-] ERROR!: O365 PSSession failed admin check! [2/2]" "Red"	
		color_out "[-] ERROR!: O365 PSSession failed Get-Mailbox check!`n" "Red"
		Exit
	}

	color_out "[+] Passed O365 permission check" "Green"
}

#Start of Script

Write-Output $banner

if(!($mfa))
{
	do
	{
		do
		{
			try
			{
				$credObject = Get-Credential -Credential $null
			}
			catch
			{
				color_out "[-] ERROR!: Failed to provide O365 Credentials`n" "Red"
			}
		}While(!($credObject))

		try
		{
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credObject -Authentication Basic -AllowRedirection
		}
		catch
		{
			color_out "[-] ERROR!: Failed to create O365 PSSession`n" "Red"
		}
		If(!($Session))
		{
			color_out "[-] ERROR!: Failed to create O365 PSSession`n" "Red"
		}
	}While(!($Session))

	color_out "[+] Credentials are valid!" "Green"

	$Session_Import = Import-PSSession -AllowClobber $Session -DisableNameChecking -CommandName Get-Mailbox,Get-InboxRule

	O365_permission_check $Session_Import
}
else
{
	$cwd = Convert-Path .
	$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
	. "$CreateEXOPSSession\CreateExoPSSession.ps1"
	cd $cwd
	
		try
		{
			Connect-EXOPSSession
		}
		catch
		{
			color_out "[-] ERROR!: Failed to create O365 PSSession`n" "Red"
			exit
		}
	
		color_out "[+] Credentials are valid!" "Green"
	
}

if($user)
{
	
	$temp=""
	if(Test-Path $user)
	{
		color_out "[+] User parameter detected as a file" "Green"
		$u_array = (Get-Content $user | Where-Object {$_} | Foreach {$_.Trim()})
	}
	elseif($user.contains(","))
	{
		color_out "[+] User parameter detected as csv string" "Green"
		$u_array = $user.split(",")
	}
	else
	{
		color_out "[+] User parameter detected as single user string" "Green"
		$u_array = @($user)
	}	
}
else
{
	$u_array = Get-Mailbox -ResultSize Unlimited | foreach{$_.PrimarySmtpAddress}
}

color_out "[+] User list created!" "Green"

$userCount = $u_array.count

color_out "[+] Number of users in list: $userCount" "Green"

color_out "[+] Collecting any Forwarding Email Addresses" "Green"

$SMTP_Forwards = [System.Collections.ArrayList]@()

For ($i=0; $i -lt $userCount; $i++) 
{
	$currentAccount = $u_array[$i]
	Write-Progress -Id 1 -Activity $("Working on mailbox: " + $currentAccount) -PercentComplete (($i / $u_array.count) * 100) 
	$mb = Get-Mailbox $currentAccount 
	if($mb.ForwardingSmtpAddress -ne $null)
	{
		$SMTP_Forwards.Add(($mb | Select UserPrincipalName,ForwardingSmtpAddress,DelivertoMailboxAndForward)) | Out-Null
	}
}

if($SMTP_Forwards.Count -gt 0)
{
	$SMTP_Forwards |  ConvertTo-Csv -NoTypeInformation | Out-File "EmailForwarding.csv" -Encoding UTF8 
}

color_out "[+] Collecting Mailbox Rules" "Green"	

For ($i=0; $i -lt $userCount; $i++) 
{
	$currentAccount = $u_array[$i]
	Write-Progress -Id 1 -Activity $("Working on mailbox: " + $currentAccount) -PercentComplete (($i / $u_array.count) * 100) 
	$ErrorActionPreference = "Stop"
	While($True)
	{
		#Small Sleep
		Start-Sleep -m 500
		try
		{
			if(!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" -And $_.State -eq "Opened"}))
			{
				While(!(Test-Connection outlook.office365.com -Count 1 -Quiet -ErrorAction SilentlyContinue))
				{
					color_out "[-] ERROR!: Unable to ping outlook.office365.com, will retry in 30 seconds..." "Red"		
					Start-Sleep -s 30
				}
				
				color_out "[-] ERROR!: Microsoft.Exchange PSSession is broken" "Red"
				
				if(!($mfa))
				{
					if($Session)
					{
						Remove-PSSession $Session
					}
				
					color_out "[-] Creating new Microsoft.Exchange PSSession" "Magenta"
					$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credObject -Authentication Basic -AllowRedirection
					if(!($Session))
					{
						color_out "[-] ERROR!: Failed to create O365 PSSession`n" "Red"
					}
					else
					{
						color_out "[+] New Microsoft.Exchange PSSession created" "Green"
						$session_import_result = Import-PSSession -AllowClobber $Session -DisableNameChecking -CommandName Get-Mailbox,Get-InboxRule
					}
				}
				else
				{
					Connect-EXOPSSession -SessionOption $SessionOptions -DisableNameChecking  | Out-Null
				}
			}
	
			#Garbage Collection
			[System.GC]::Collect()
			
			$rules = Get-InboxRule -Mailbox $u_array[$i] -WarningAction silentlyContinue
		}
		catch
		{
			continue
		}
		
		if ($rules) 
		{
			foreach($rule in $rules)
			{
				$tempPsObject = New-Object PsObject -property @{
					'user' = $u_array[$i]
					'name' = $rule.name
					'priority' = $rule.priority
					}	
				rule_parser $rule.description $tempPSobject
				$tempPsObject | ConvertTo-Json | Out-File $outputfile -Encoding UTF8 -Append
			}
		}
		break
	}
	$ErrorActionPreference = "Continue"
}

#script ending
if(!($mfa))
{
	Remove-PSSession $Session
	color_out "[*] Removing Created Microsoft.Exchange PSSession" "Green"
}
else
{
	Get-PSSession | Remove-PSSession
}
[System.GC]::Collect()
color_out "[+] Script Complete!" "Green"
color_out "[+] Goodbye!`n" "Green"
