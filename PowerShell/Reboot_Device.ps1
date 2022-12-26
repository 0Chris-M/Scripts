#>
# Checks whether or not the system requires a reboot, and remediates accordingly.
#shutdown.exe -d switches documentation:
#https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/shutdown
#https://docs.microsoft.com/en-us/windows/win32/shutdown/system-shutdown-reason-codes

$sysInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"

if($sysInfo.RebootRequired)
{ 
    shutdown.exe /r /t 120 /c "Installing Patches" /f /d p:4:1
}
else
{
    "Reboot not required"
	exit 0
}


