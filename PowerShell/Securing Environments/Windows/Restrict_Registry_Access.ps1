#####################################################################
# Prevent user access to registry editing tools (regedit, etc)
# https://admx.help/?Category=Windows_10_2016&Policy=Microsoft.Policies.WindowsTools::DisableRegedit

# Regex pattern for SIDs
$PatternSID = 'S-1-\d+-\d+-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users
$ProfileList = gp 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
    Select  @{name="SID";expression={$_.PSChildName}}, 
            @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
 
# Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
$LoadedHives = gci Registry::HKEY_USERS | ? {$_.PSChildname -match $PatternSID} | Select @{name="SID";expression={$_.PSChildName}}
 
# Get all users that are not currently logged
$UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select @{name="SID";expression={$_.InputObject}}, UserHive, Username
 
# Loop through each profile on the machine
Foreach ($item in $ProfileList) {
    # Load User ntuser.dat if it's not already loaded
    IF ($item.SID -in $UnloadedHives.SID) {
        reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
    }
     
     #####################################################################
    # Modify registry entries below:
  
    ##################################################################### 
	$regedloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Microsoft\Windows\CurrentVersion\Policies\System
	If ($regedloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Microsoft\Windows\CurrentVersion\Policies\System -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name "DisableRegistryTools" -Value 2 -Type DWord -Force

    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}