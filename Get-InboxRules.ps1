##### Uncomment below dependancies if required on machine #####
#Install-module -name MSOnline
#Install-Module -Name AzureAD
#Install-Module ExchangeOnlineManagement

Import-Module ExchangeOnlineManagement

$users = @()
$counter = 0
$rulez = @()
$user_rules=@()
$rulescounter = 0

Write-Progress -id 0 -Activity "Running Inbox Rules retrieval script"
Start-Sleep -Milliseconds 200

Write-Progress -id 1 -ParentId 0 -Activity "Attempting connection to MSOL..." -Status "Please sign in when prompted."
Start-Sleep -Milliseconds 200
Connect-MsolService
if($?)
	{
		Write-Progress -id 1 -ParentId 0 -Activity "Connected to MSOL Successfully."
		Start-Sleep -Milliseconds 200

		Write-Progress -id 2 -ParentId 0 -Activity "Attempting to retrieve users from MSOL..."
		Start-Sleep -Milliseconds 200
		$msol_get = Get-MsolUser -All | select -expand 'UserPrincipalName'
		foreach ($user in $msol_get)
			{
				$counter++
				Write-Progress -id 2 -ParentId 0 -Activity "Retrieving users..." -Status "Retrieved user $user successfully!"  -PercentComplete (($counter / $msol_get.count) * 100)
				$users+=$user
			}
		if($?)
			{
			}
		else
			{
				Write-Progress -id 2 -ParentId 0 -Activity "Could not retrieve users"
			}


		Write-Progress -id 3 -ParentId 0 -Activity "Attempting connection to Exchange Online." -Status "Please sign in when prompted."
		Start-Sleep -Milliseconds 200
		Connect-ExchangeOnline -ShowBanner:$false
		if($?)
			{
				Write-Progress -id 3 -ParentId 0 -Activity "Connected to Exchange Online Successfully"
				Start-Sleep -Milliseconds 200
				if ($users)
					{
						Write-Progress -id 4 -ParentId 0 -Activity "Attempting to retrieve mailbox rules..."
						Start-Sleep -Milliseconds 200
						foreach ($user in $users)
						{
						$rulescounter++
						$user_rules=Get-InboxRule -Mailbox $user -includehidden 2> $null | Select-object *
						if($?)
							{
								$rulez+=$user_rules
								Write-Progress -id 4 -ParentId 0 -Activity "Retrieving $user mailbox rules..." -PercentComplete (($rulescounter / $users.count) * 100)
							}
						}
					}
			}
		else
			{
				"Unable to establish connection to Exchange Online"
			}
	}
else
	{
		"Unable to establish connection to MSOL."
	}
if ($rulez)
	{
		$rulez | Export-Csv -path c:\user_mailbox_rules.csv -NoTypeInformation
		"Exporting User Mailbox Rules to C:\user_mailbox_rules.csv"
	}