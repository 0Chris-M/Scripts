#####################################################################
# Harden Office environments



 #####################################################################
                   ## Code for HKEY_USERS ##
 ##################################################################### 
     
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
    # Disable ActiveX from initializing in all Office Applications
	$cmnsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security
	If ($cmnsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Name disableallactivex -Value 1 -Type DWord -Force
    
	
    # Disable add-ins for excel, project, powerpoint, visio, word, access, publisher, and OneNote.
	$exclsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security
	If ($exclsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Force
	}
	$msprojsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'ms project'\security
	If ($msprojsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'ms project'\security -Force
	}
	$pwrpntsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security
	If ($pwrpntsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Force
	}
	$vissec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security
	If ($vissec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security -Force
	}
	$wrdsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security
	If ($wrdsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security -Force
	}
	$acsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security
	If ($acsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security -Force
	}
	$pubsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security
	If ($pubsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security -Force
	}
	$onesec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security
	If ($onesec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Name disablealladdins -Value 1 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\"ms project"\security -Name disablealladdins -Value 1 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Name disablealladdins -Value 1 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security -Name disablealladdins -Value 1 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security -Name disablealladdins -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security -Name disablealladdins -Value 1 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security -Name disablealladdins -Value 1 -Type DWord -Force    
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Name disablealladdins -Value 1 -Type DWord -Force    
    
    ### Disable DDE
    # Common
	$cmnsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security
	If ($cmnsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Name blockedextensions -Value 'ps1;hta;wsf;js;py;rb;rtf' -Type String -Force          
      
	# Excel
	$exclext = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content"
	If ($exclext -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content"
	}
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name disableddeserverlaunch -Value 1 -Type DWord -Force      
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name disableddeserverlookup -Value 1 -Type DWord -Force
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name enableblockunsecurequeryfiles -Value 1 -Type DWord -Force
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name DataConnectionWarnings -Value 2 -Type DWord -Force      
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name RichDataConnectionWarnings -Value 2 -Type DWord -Force
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name WorkbookLinkWarnings -Value 2 -Type DWord -Force

	# Word
	$wrdopt = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options\wordmail
	If ($wrdopt -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options\wordmail -Force
	}	
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options -Name dontupdatelinks -Value 1 -Type DWord -Force      
	Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options\wordmail -Name dontupdatelinks -Value 1 -Type DWord -Force      


    # Block OLE
	$exclsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security
	If ($exclsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Force
	}
	$msprojsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'ms project'\security
	If ($msprojsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'ms project'\security -Force
	}
	$pwrpntsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security
	If ($pwrpntsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Force
	}
	$vissec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security
	If ($vissec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security -Force
	}
	$wrdsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security
	If ($wrdsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security -Force
	}
	$acsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security
	If ($acsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security -Force
	}
	$pubsec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security
	If ($pubsec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security -Force
	}
	$onesec = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security
	If ($onesec -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\PowerPoint\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\security -Name PackagerPrompt -Value 2 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Name PackagerPrompt -Value 2 -Type DWord -Force      

    # Excel extension hardening
	$exclext = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security
	If ($exclext -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security  -Force
	}	
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Name extensionhardening -Value 2 -Type DWord -Force      

    # Office File Validation (OFV) for Excel, PowerPoint, and Word
	$excflval = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security\filevalidation
	If ($excflval -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security\filevalidation -Force
	}
	$wrdflval = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security\filevalidation
	If ($wrdflval -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security\filevalidation -Force
	}
	$powpflval = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security\filevalidation
	If ($powpflval -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security\filevalidation -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security\filevalidation -Name enableonload -Value 1 -Type DWord -Force      
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security\filevalidation -Name enableonload -Value 1 -Type DWord -Force
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security\filevalidation -Name enableonload -Value 1 -Type DWord -Force

    # Run external programs in PowerPoint
	$powpflval = Test-Path -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security
	If ($powpflval -eq $False)
	{
		New-Item -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Force
	}
    Set-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Name runprograms -Value 0 -Type DWord -Force      


    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}