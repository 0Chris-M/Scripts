#####################################################################
# Gather requirements from client. Set Trusted Locations for Applications that require Macros, remove Trusted Locations in applications that don't require Macros.
##### Need to gather Trusted Locations network/UNC paths and update script accordingly. Make sure to remove Trusted Locations for applications that don't require Macros.


# Pattern to cycle through SIDs on machine
$PatternSID = 'S-1-\d+-\d+-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users on device
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
    # Office Common
	$offreg = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\Common\Security
	If ($offreg -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\Common\Security -Force
	}
	$off16reg = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security\'Trusted Locations'
	If ($off16reg -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security\'Trusted Locations' -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\Common\Security -Name automationsecurity -Value 2 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security -Name vbaoff -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security -Name macroruntimescanscope -Value 2 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security\'Trusted Locations' -Name 'allow user locations' -Value 0 -Type DWord -Force

    # Access
	$acctrusdoc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents'
	If ($acctrusdoc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents' -Force
	}
	$acctrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\Location1
	If ($acctrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\Location1 -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security -Name blockcontentexecutionfrominternet -Value 1 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents' -Name disabletrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force

    # Excel
 	$extrusdoc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents'
	If ($extrusdoc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents' -Force
	}
	$extrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\Location1
	If ($extrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\Location1 -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name excelbypassencryptedmacroscan -Value 0 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name blockcontentexecutionfrominternet -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name accessvbom -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents' -Name disabletrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name xl4macrowarningfollowvba -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force
    
    # Outlook
 	$outlksec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Outlook\Security\
	If ($outlksec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Outlook\Security\ -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Outlook\Security -Name level -Value 4 -Type DWord -Force
    
    # Powerpoint
	$pwrptrusdoc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents'
	If ($pwrptrusdoc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents' -Force
	}
	$pwrptrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\Location1
	If ($pwrptrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\Location1 -Force
	} 
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name powerpointbypassencryptedmacroscan -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name blockcontentexecutionfrominternet -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name accessvbom -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents' -Name disabletrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force
 
    # MS Project
 	$prjtsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\
	If ($prjtsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\ -Force
	}
	$prjttrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\Location1
	If ($prjttrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\Location1 -Force
	}  
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force
 
    # MS Publisher
 	$pubsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security\
	If ($pubsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security\ -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security -Name automationsecuritypublisher -Value 3 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security -Name vbawarnings -Value 4 -Type DWord -Force

    # Visio
 	$visapp = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application
	If ($visapp -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application -Force
	}
	$vistrusdoc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents'
	If ($vistrusdoc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents' -Force
	}
	$vistrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\Location1
	If ($vistrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\Location1 -Force
	}   
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application -Name createvbaprojects -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application -Name loadvbaprojectsfromtext -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security -Name blockcontentexecutionfrominternet -Value 1 -Type DWord -Force        
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents' -Name disabletrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force

    # Word
	$wordtrusdoc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents'
	If ($wordtrusdoc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents' -Force
	}
	$wordtrusloc = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\Location1
	If ($wordtrusloc -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\Location1 -Force
	}   
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name wordbypassencryptedmacroscan -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name blockcontentexecutionfrominternet -Value 1 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name accessvbom -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents' -Name disabletrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name vbawarnings -Value 4 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations' -Name allownetworklocations -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations' -Name alllocationsdisabled -Value 0 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\Location1 -Name path -Value 'UNC PATH' -Type ExpandString -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\Location1 -Name description -Value 'Trusted Location' -Type String -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\Location1 -Name allowsubfolders -Value 1 -Type DWord -Force



    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}