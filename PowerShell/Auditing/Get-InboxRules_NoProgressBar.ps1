##### Uncomment below dependancies if required on machine #####
#Install-module -name MSOnline
#Install-Module -Name AzureAD
#Install-Module ExchangeOnlineManagement

Import-Module ExchangeOnlineManagement


"Attempting connection to MSOL."
"Please sign in when promted..."
Connect-MsolService
if($?)
	{
		"Connected to MSOL Successfully.`n"
	}
else
{
	"Unable to establish connection to MSOL.`n"
}


"Attempting to retrive users from MSOL..."
$users = @()
foreach ($user in (Get-MsolUser -All | select -expand 'UserPrincipalName'))
{$users+=$user}
if($?)
	{
		"Users retrieved succesfully`n"
	}
else
{
	"Could not retrieve users`n"
}


"Attempting connection to Exchange Online."
"Please sign in when promted..."
Connect-ExchangeOnline
if($?)
	{
		"Connected to Exchange Online Successfully`n"
	}
else
{
	"Unable to establish connection to Exchange Online`n"
}


"Attempting to retrive mailbox rules..."
$rulez = @()
$user_rules=@()
if ($users)
{
	foreach ($user in $users)
	{
	$user_rules= Get-InboxRule -Mailbox $user -includehidden | Select-object *
	if($?)
		{
			"User Inbox Rules Retrieved Successfully`n"
			$rulez+=$user_rules
		}
	}
}


if ($rulez)
{
"Rules retrieved succesfully.`n"
$rulez | Export-Csv -path .\user_mailbox_rules.csv -NoTypeInformation
"Outputting rules to C:\user_mailbox_rules.csv"
}
