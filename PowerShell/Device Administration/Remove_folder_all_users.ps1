#Gathers list of users
$users = Get-ChildItem c:\users|where{$_.name -notmatch 'Public|Default'}

# Iterates through each user, checks if folder is present and recursively deletes folder and contents
Foreach($user in $users) {
	$ofc = Test-Path -Path "C:\Users\$($user.Name)\AppData\Local\Microsoft\Office\16.0\OfficeFileCache"
	If ($ofc -eq $True)
	{
		Write-Output "OfficeFileCache present. Deleting..."
		Remove-Item -Path "C:\Users\$($user.Name)\AppData\Local\Microsoft\Office\16.0\OfficeFileCache" -Force -Recurse
		Write-Output "Done."
	}
	else
	{
		Write-Output "OfficeFileCache not found for user $($user.Name)."
	}
}