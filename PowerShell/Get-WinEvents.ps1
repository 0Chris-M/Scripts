mkdir C:\IncidentResponse -ErrorAction SilentlyContinue
cd C:\IncidentResponse
mkdir C:\IncidentResponse\WinLogs_$env:computername -ErrorAction SilentlyContinue
copy C:\Windows\System32\winevt\Logs\* C:\IncidentResponse\WinLogs_$env:computerName
Compress-Archive -Path C:\IncidentResponse\WinLogs_$env:computername -DestinationPath C:\IncidentResponse\WinLogs_$env:computername.zip