# Pulls list of any Local Admin account present on the machine

# If PowerShell < v5.1, then exit 
if ($PSVersionTable.PSVersion -lt [version]'5.1') { Exit 0 }
$adminNames = Get-LocalGroupMember -Group Administrators
If ($adminNames) { Return $adminNames } else { Exit 0 }
