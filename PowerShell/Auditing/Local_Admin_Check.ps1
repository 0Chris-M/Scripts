# Pulls list of any Local Admin account present on the machine

# If Windows OS isn't at least Win10 v1607 or Server 2016 or PowerShell < v5.1, then exit 
if ($PSVersionTable.PSVersion -lt [version]'5.1') { Exit 0 }
$adminNames = Get-LocalGroupMember -Group Administrators
If ($adminNames) { Return $adminNames } else { Exit 0 }
