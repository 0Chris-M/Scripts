#CVE-2022-30190:
#https://msrc-blog.microsoft.com/2022/05/30/guidance-for-cve-2022-30190-microsoft-support-diagnostic-tool-vulnerability/

# If registry key exists proceed to Remediation Code, if not then exit.
$msdtkey = Test-Path -Path registry::HKEY_CLASSES_ROOT\ms-msdt
if ($msdtkey -eq $True) {
    Write-Output "ms-msdt key present."
	Write-Output "Exporting key to C:\Windows\Temp\ms-msdt-backup\ and deleting key..."
    # Create backup directory in C:\Windows\Temp\ms-msdt-backup\
	# This can be changed to a directory of your choosing.
	$testpath = "C:\Windows\Temp\ms-msdt-backup\"

	If (Test-Path $testpath)
	{
		# Export the current registry key into the previously created directory
		reg export "HKEY_CLASSES_ROOT\ms-msdt" "C:\Windows\Temp\ms-msdt-backup\ms-msdt.reg" /y
		# Remove the vulnerable registry entry.
		reg delete "HKEY_CLASSES_ROOT\ms-msdt" /f
	}
	else
	{
		mkdir "C:\Windows\Temp\ms-msdt-backup\"
		# Export the current registry key into the previously created directory
		reg export "HKEY_CLASSES_ROOT\ms-msdt" "C:\Windows\Temp\ms-msdt-backup\ms-msdt.reg" /y
		# Remove the vulnerable registry entry.
		reg delete "HKEY_CLASSES_ROOT\ms-msdt" /f
	}
} else {
    Write-Output "ms-msdt key is not present. Exit"
    exit 0
}


